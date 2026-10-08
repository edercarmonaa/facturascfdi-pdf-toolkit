"""
Extrae datos de CFDI desde PDFs, consulta el estatus en el servicio SOAP del SAT
y genera un Excel con los datos extraidos y la respuesta del SAT.

Requisitos:
    python3 -m pip install pymupdf opencv-python numpy pandas openpyxl tqdm
    Opcional para paginas escaneadas con QR danado: instalar tesseract en el PATH.

Uso:
    python3 consultar_estatus_sat.py
    python3 consultar_estatus_sat.py /ruta/a/pdfs -o resultado_sat_cfdi.xlsx
    python3 consultar_estatus_sat.py --sin-consultar
"""

import argparse
import html
import logging
import re
import shutil
import subprocess
import sys
import tempfile
import urllib.error
import urllib.request
import xml.etree.ElementTree as ET
from dataclasses import dataclass
from decimal import Decimal, InvalidOperation
from pathlib import Path
from urllib.parse import parse_qs, urlencode, urlparse

try:
    import cv2
    import fitz  # PyMuPDF
    import numpy as np
    import pandas as pd
except ImportError as e:
    print(
        "Falta instalar una dependencia. Ejecuta:\n"
        "python3 -m pip install pymupdf opencv-python numpy pandas openpyxl tqdm\n"
        f"\nDetalle: {e}",
        file=sys.stderr,
    )
    sys.exit(1)

try:
    from tqdm import tqdm
except ImportError:
    def tqdm(iterable, **kwargs):
        return iterable


SAT_ENDPOINT = "https://consultaqr.facturaelectronica.sat.gob.mx/ConsultaCFDIService.svc"
SOAP_ACTION = "http://tempuri.org/IConsultaCFDIService/Consulta"
ARCHIVO_SALIDA = Path("resultado_sat_cfdi.xlsx")
DPIS_A_INTENTAR = [200, 350, 500]
OCR_DPI = 300

COLUMNAS_EXCEL = [
    "nombre_archivo",
    "ruta",
    "pagina",
    "uuid",
    "qr",
    "re",
    "rr",
    "tt",
    "fe",
    "expresion_impresa",
    "metodo_extraccion",
    "estado_extraccion",
    "sat_codigo_estatus",
    "sat_estado",
    "sat_es_cancelable",
    "sat_estatus_cancelacion",
    "sat_validacion_efos",
    "sat_error",
    "sat_respuesta_xml",
]

UUID_RE = re.compile(
    r"[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}"
)
RFC_RE = re.compile(r"\b[A-Z&Ñ]{3,4}\d{6}[A-Z0-9]{3}\b")
RFC_CANDIDATO_RE = re.compile(r"\b[A-Z&Ñ0-9]{3,4}[0-9OILSB]{6}[A-Z0-9]{3}\b")
# Si tu caso lo requiere, agrega localmente RFCs de PAC para excluirlos de la
# deteccion de emisor/receptor. No agregues identificadores reales al repo.
RFC_PAC_CONOCIDOS: set[str] = set()

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%H:%M:%S",
)
logger = logging.getLogger(__name__)
detector = cv2.QRCodeDetector()


def ruta_para_salida(pdf_path: Path, base: Path) -> str:
    """Devuelve una ruta relativa cuando es posible para no exponer rutas locales."""
    try:
        return str(pdf_path.resolve().relative_to(base.resolve()))
    except ValueError:
        return pdf_path.name


@dataclass
class DatosCfdi:
    nombre_archivo: str
    ruta: str
    pagina: int | str
    uuid: str = ""
    re: str = ""
    rr: str = ""
    tt: str = ""
    fe: str = ""
    metodo_extraccion: str = ""
    estado_extraccion: str = ""
    qr_texto: str = ""

    def expresion_impresa(self) -> str:
        if not all([self.re, self.rr, self.tt, self.uuid, self.fe]):
            return ""
        return "?" + urlencode({
            "re": self.re,
            "rr": self.rr,
            "tt": self.tt,
            "id": self.uuid,
            "fe": self.fe,
        })

    def fila_base(self) -> dict:
        return {
            "nombre_archivo": self.nombre_archivo,
            "ruta": self.ruta,
            "pagina": self.pagina,
            "uuid": self.uuid,
            "qr": self.qr_texto,
            "re": self.re,
            "rr": self.rr,
            "tt": self.tt,
            "fe": self.fe,
            "expresion_impresa": self.expresion_impresa(),
            "metodo_extraccion": self.metodo_extraccion,
            "estado_extraccion": self.estado_extraccion,
            "sat_codigo_estatus": "",
            "sat_estado": "",
            "sat_es_cancelable": "",
            "sat_estatus_cancelacion": "",
            "sat_validacion_efos": "",
            "sat_error": "",
            "sat_respuesta_xml": "",
        }


def pagina_a_imagen(pagina: "fitz.Page", dpi: int) -> np.ndarray:
    pix = pagina.get_pixmap(dpi=dpi)
    img = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.width, pix.n)

    if pix.n == 4:
        return cv2.cvtColor(img, cv2.COLOR_RGBA2BGR)
    if pix.n == 1:
        return cv2.cvtColor(img, cv2.COLOR_GRAY2BGR)
    return cv2.cvtColor(img, cv2.COLOR_RGB2BGR)


def detectar_qrs_en_imagen(img: np.ndarray) -> list[str]:
    try:
        ok, datos, _, _ = detector.detectAndDecodeMulti(img)
        if ok:
            return [dato for dato in datos if dato]
    except cv2.error:
        pass

    dato, _, _ = detector.detectAndDecode(img)
    return [dato] if dato else []


def limpiar_total(valor: str) -> str:
    valor = re.sub(r"[^0-9.,]", "", valor or "")
    if not valor:
        return ""
    if "," in valor and "." in valor:
        valor = valor.replace(",", "")
    elif "," in valor and "." not in valor:
        valor = valor.replace(",", ".")
    try:
        return f"{Decimal(valor):.6f}"
    except InvalidOperation:
        return valor


def corregir_uuid_por_ocr(uuid: str) -> str:
    """Corrige confusiones OCR comunes en UUIDs, especialmente O/0."""
    if not uuid:
        return ""
    uuid = uuid.upper().replace("O", "0")
    return uuid if UUID_RE.fullmatch(uuid) else ""


def normalizar_uuid(texto: str) -> str:
    if not texto:
        return ""

    query = parse_qs(urlparse(texto).query)
    for llave in ("id", "uuid", "UUID"):
        valores = query.get(llave)
        if valores and UUID_RE.fullmatch(valores[0].strip()):
            return valores[0].strip().upper()

    texto_compacto = re.sub(r"\s+", "", texto)
    candidatos = [
        corregir_uuid_por_ocr(match.group(0))
        for origen in (texto, texto_compacto)
        for match in UUID_RE.finditer(origen)
    ]
    patron_ocr = re.compile(
        r"[0-9A-FO]{8}-[0-9A-FO]{4}-[0-9A-FO]{4}-[0-9A-FO]{4}-[0-9A-FO]{12}",
        re.IGNORECASE,
    )
    candidatos.extend(
        corregir_uuid_por_ocr(match.group(0))
        for match in patron_ocr.finditer(texto_compacto)
    )
    candidatos = [uuid for uuid in candidatos if uuid]
    if not candidatos:
        return ""

    frecuencia: dict[str, int] = {}
    for uuid in candidatos:
        frecuencia[uuid] = frecuencia.get(uuid, 0) + 1
    return max(frecuencia, key=frecuencia.get)


def extraer_desde_qr(qr_texto: str, pdf_path: Path, base: Path, pagina: int) -> DatosCfdi:
    query = parse_qs(urlparse(qr_texto).query)

    def obtener(llave: str) -> str:
        valores = query.get(llave)
        return valores[0].strip() if valores else ""

    return DatosCfdi(
        nombre_archivo=pdf_path.name,
        ruta=ruta_para_salida(pdf_path, base),
        pagina=pagina,
        uuid=normalizar_uuid(qr_texto),
        re=obtener("re").upper(),
        rr=obtener("rr").upper(),
        tt=limpiar_total(obtener("tt")),
        fe=obtener("fe"),
        metodo_extraccion="QR",
        estado_extraccion="Datos completos desde QR",
        qr_texto=qr_texto,
    )


def ocr_imagen(img: np.ndarray, psm: str = "6") -> str:
    if not shutil.which("tesseract"):
        return ""

    gris = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    variantes = [
        gris,
        cv2.resize(gris, None, fx=2, fy=2, interpolation=cv2.INTER_CUBIC),
    ]
    textos: list[str] = []

    for variante in variantes:
        _, binaria = cv2.threshold(variante, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
        with tempfile.NamedTemporaryFile(suffix=".png") as archivo_temporal:
            cv2.imwrite(archivo_temporal.name, binaria)
            try:
                proceso = subprocess.run(
                    ["tesseract", archivo_temporal.name, "stdout", "--psm", psm],
                    text=True,
                    capture_output=True,
                    check=False,
                    timeout=30,
                )
            except subprocess.TimeoutExpired:
                continue
        if proceso.stdout:
            textos.append(proceso.stdout)

    return "\n".join(textos)


def extraer_texto_ocr_pagina(pagina: "fitz.Page") -> str:
    img = pagina_a_imagen(pagina, OCR_DPI)
    alto, ancho = img.shape[:2]
    recortes = [
        img,
        img[: int(alto * 0.45), :],
        img[: int(alto * 0.35), int(ancho * 0.45) :],
        img[int(alto * 0.38) : int(alto * 0.72), :],
        img[int(alto * 0.45) : int(alto * 0.58), :],
    ]
    textos = [ocr_imagen(recorte) for recorte in recortes if recorte.size]
    return "\n".join(texto for texto in textos if texto)


def extraer_total_ocr(texto: str) -> str:
    patrones = [
        r"\bTotal\b\s*\$?\s*([0-9][0-9,]*\.?\d{0,6})",
        r"\bTOTAL\b\s*\$?\s*([0-9][0-9,]*\.?\d{0,6})",
    ]
    for patron in patrones:
        matches = re.findall(patron, texto, flags=re.IGNORECASE)
        if matches:
            return limpiar_total(matches[-1])
    return ""


def extraer_rfc_ocr(texto: str) -> tuple[str, str]:
    def normalizar_rfc_candidato(valor: str) -> str:
        valor = re.sub(r"[^A-Z&Ñ0-9]", "", valor.upper())
        if len(valor) not in (12, 13):
            return ""

        letras = 4 if len(valor) == 13 else 3
        inicio = valor[:letras]
        fecha = (
            valor[letras : letras + 6]
            .replace("O", "0")
            .replace("I", "1")
            .replace("L", "1")
            .replace("S", "5")
            .replace("B", "8")
        )
        fin = valor[letras + 6 :].replace("O", "0")
        normalizado = inicio + fecha + fin
        return normalizado if RFC_RE.fullmatch(normalizado) else ""

    texto_mayus = texto.upper()
    candidatos: list[str] = []
    for match in RFC_CANDIDATO_RE.findall(texto_mayus):
        rfc = normalizar_rfc_candidato(match)
        if rfc:
            candidatos.append(rfc)

    unicos = [
        rfc
        for rfc in dict.fromkeys(candidatos)
        if rfc not in RFC_PAC_CONOCIDOS
    ]

    receptor = ""
    for linea in texto_mayus.splitlines():
        if "RECEPTOR" not in linea:
            continue
        rfcs_linea = [
            normalizar_rfc_candidato(match)
            for match in RFC_CANDIDATO_RE.findall(linea)
        ]
        rfcs_linea = [rfc for rfc in rfcs_linea if rfc and rfc not in RFC_PAC_CONOCIDOS]
        if rfcs_linea:
            receptor = rfcs_linea[-1]
            break

    if not receptor and len(unicos) >= 2:
        receptor = unicos[1]

    emisor = ""
    for linea in texto_mayus.splitlines():
        if "CERTIFIC" in linea:
            continue
        if "RFC" not in linea and "R.F.C" not in linea:
            continue
        rfcs_linea = [
            normalizar_rfc_candidato(match)
            for match in RFC_CANDIDATO_RE.findall(linea)
        ]
        for rfc in rfcs_linea:
            if rfc and rfc != receptor and rfc not in RFC_PAC_CONOCIDOS:
                emisor = rfc
                break
        if emisor:
            break

    if not emisor:
        for rfc in unicos:
            if rfc != receptor:
                emisor = rfc
                break

    return emisor, receptor


def extraer_sello_ocr(texto: str, uuid: str) -> str:
    if not uuid:
        return ""

    texto_lineal = re.sub(r"\s+", "", texto)
    uuid_sin_guiones = re.escape(uuid.upper())
    match = re.search(
        uuid_sin_guiones + r"\|[^|]*\|[^|]*\|([A-Za-z0-9+/=]{20,})\|",
        texto_lineal.upper(),
    )
    if match:
        return match.group(1)

    return ""


def extraer_desde_ocr(
    pagina: "fitz.Page",
    pdf_path: Path,
    base: Path,
    pagina_num: int,
) -> DatosCfdi | None:
    texto = extraer_texto_ocr_pagina(pagina)
    uuid = normalizar_uuid(texto)
    if not uuid:
        return None

    re_rfc, rr_rfc = extraer_rfc_ocr(texto)
    sello = extraer_sello_ocr(texto, uuid)
    fe = sello[-8:] if len(sello) >= 8 else ""

    datos = DatosCfdi(
        nombre_archivo=pdf_path.name,
        ruta=ruta_para_salida(pdf_path, base),
        pagina=pagina_num,
        uuid=uuid,
        re=re_rfc,
        rr=rr_rfc,
        tt=extraer_total_ocr(texto),
        fe=fe,
        metodo_extraccion="OCR",
        estado_extraccion="Datos extraidos por OCR",
    )

    faltantes = [campo for campo in ("re", "rr", "tt", "fe") if not getattr(datos, campo)]
    if faltantes:
        datos.estado_extraccion = "OCR incompleto: faltan " + ", ".join(faltantes)

    return datos


def buscar_cfdis_en_pdf(pdf_path: Path, base: Path) -> list[DatosCfdi]:
    resultados: list[DatosCfdi] = []
    vistos: set[tuple[str, int]] = set()

    with fitz.open(pdf_path) as doc:
        if doc.needs_pass:
            return [DatosCfdi(
                pdf_path.name,
                ruta_para_salida(pdf_path, base),
                "",
                estado_extraccion="PDF protegido",
            )]

        for idx in range(doc.page_count):
            pagina = doc[idx]
            pagina_num = idx + 1
            qrs_encontrados: list[str] = []

            for dpi in DPIS_A_INTENTAR:
                img = pagina_a_imagen(pagina, dpi)
                qrs_encontrados = detectar_qrs_en_imagen(img)
                if qrs_encontrados:
                    break

            if qrs_encontrados:
                for qr_texto in qrs_encontrados:
                    datos = extraer_desde_qr(qr_texto, pdf_path, base, pagina_num)
                    llave = (datos.uuid, pagina_num)
                    if datos.uuid and llave not in vistos:
                        resultados.append(datos)
                        vistos.add(llave)
                continue

            datos_ocr = extraer_desde_ocr(pagina, pdf_path, base, pagina_num)
            if datos_ocr and datos_ocr.uuid:
                llave = (datos_ocr.uuid, pagina_num)
                if llave not in vistos:
                    resultados.append(datos_ocr)
                    vistos.add(llave)

    if not resultados:
        resultados.append(DatosCfdi(
            nombre_archivo=pdf_path.name,
            ruta=ruta_para_salida(pdf_path, base),
            pagina="",
            estado_extraccion="No encontrado",
        ))

    return resultados


def consultar_sat(
    expresion_impresa: str,
    timeout: int = 30,
    incluir_xml: bool = False,
) -> dict:
    if not expresion_impresa:
        return {"sat_error": "No se puede consultar: faltan re, rr, tt, id o fe"}

    body = f"""<?xml version="1.0" encoding="utf-8"?>
<soapenv:Envelope xmlns:soapenv="http://schemas.xmlsoap.org/soap/envelope/" xmlns:tem="http://tempuri.org/">
  <soapenv:Header/>
  <soapenv:Body>
    <tem:Consulta>
      <tem:expresionImpresa>{html.escape(expresion_impresa)}</tem:expresionImpresa>
    </tem:Consulta>
  </soapenv:Body>
</soapenv:Envelope>""".encode("utf-8")

    request = urllib.request.Request(
        SAT_ENDPOINT,
        data=body,
        headers={
            "Content-Type": "text/xml; charset=utf-8",
            "SOAPAction": f'"{SOAP_ACTION}"',
            "User-Agent": "cfdi-status-checker/1.0",
        },
        method="POST",
    )

    try:
        with urllib.request.urlopen(request, timeout=timeout) as response:
            xml_text = response.read().decode("utf-8", errors="replace")
    except urllib.error.URLError as e:
        return {"sat_error": f"Error de conexion: {e.reason}"}
    except urllib.error.HTTPError as e:
        return {"sat_error": f"HTTP {e.code}: {e.read().decode('utf-8', errors='replace')[:500]}"}
    except TimeoutError:
        return {"sat_error": "Timeout consultando SAT"}

    resultado = {"sat_respuesta_xml": xml_text if incluir_xml else "", "sat_error": ""}
    try:
        root = ET.fromstring(xml_text)
    except ET.ParseError as e:
        resultado["sat_error"] = f"Respuesta XML invalida: {e}"
        return resultado

    nombres = {
        "CodigoEstatus": "sat_codigo_estatus",
        "Estado": "sat_estado",
        "EsCancelable": "sat_es_cancelable",
        "EstatusCancelacion": "sat_estatus_cancelacion",
        "ValidacionEFOS": "sat_validacion_efos",
    }
    for elem in root.iter():
        tag = elem.tag.split("}", 1)[-1]
        if tag in nombres:
            resultado[nombres[tag]] = elem.text or ""

    return resultado


def parsear_argumentos() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Extrae datos CFDI de PDFs y consulta el estatus SOAP en SAT."
    )
    parser.add_argument(
        "carpeta",
        nargs="?",
        default=".",
        help="Carpeta con PDFs a procesar.",
    )
    parser.add_argument(
        "-o",
        "--salida",
        default=str(ARCHIVO_SALIDA),
        help=f"Excel de salida. Por defecto: {ARCHIVO_SALIDA}",
    )
    parser.add_argument(
        "--sin-consultar",
        action="store_true",
        help="Solo extrae datos y genera Excel, sin llamar al servicio del SAT.",
    )
    parser.add_argument(
        "--incluir-xml",
        action="store_true",
        help="Incluye la respuesta XML cruda del SAT en el Excel generado.",
    )
    return parser.parse_args()


def main() -> None:
    args = parsear_argumentos()
    carpeta = Path(args.carpeta)
    salida = Path(args.salida)

    pdfs = sorted(p for p in carpeta.glob("*.pdf") if not p.name.startswith("~$"))
    if not pdfs:
        logger.warning("No se encontraron PDFs en '%s'.", carpeta)
        return

    filas: list[dict] = []
    for pdf_path in tqdm(pdfs, desc="Extrayendo CFDI"):
        for datos in buscar_cfdis_en_pdf(pdf_path, carpeta):
            fila = datos.fila_base()
            if not args.sin_consultar:
                respuesta_sat = consultar_sat(
                    fila["expresion_impresa"],
                    incluir_xml=args.incluir_xml,
                )
                fila.update(respuesta_sat)
            filas.append(fila)

    df = pd.DataFrame(filas, columns=COLUMNAS_EXCEL)
    df.to_excel(salida, index=False)

    consultados = df["sat_estado"].fillna("").ne("").sum() if "sat_estado" in df else 0
    logger.info(
        "Listo. %s filas generadas, %s con respuesta SAT. Archivo generado: %s",
        len(df),
        consultados,
        salida,
    )


if __name__ == "__main__":
    main()
