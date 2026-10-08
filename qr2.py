"""
Extrae códigos QR (y su UUID de factura/CFDI) de todos los PDFs en una carpeta
y guarda los resultados en Excel.

Requisitos:
    pip install pymupdf opencv-python pandas openpyxl tqdm
    (opcional, para el fallback por OCR: instalar Tesseract y tenerlo en el PATH)

Uso:
    python extraer_qr_pdfs.py [carpeta_con_pdfs]
"""

import argparse
import logging
import re
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path
from urllib.parse import parse_qs, urlparse

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
except ImportError:  # tqdm es opcional
    def tqdm(iterable, **kwargs):
        return iterable

# ---------------------------------------------------------------------------
# Configuración
# ---------------------------------------------------------------------------
ARCHIVO_SALIDA = Path("resultado_qr_pdfs.xlsx")
DPIS_A_INTENTAR = [200, 350, 500]  # sube la resolución si no encuentra nada
COLUMNAS_EXCEL = [
    "nombre_archivo",
    "ruta",
    "pagina",
    "uuid",
    "metodo_qr",
    "estado",
]

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%H:%M:%S",
)
logger = logging.getLogger(__name__)

detector = cv2.QRCodeDetector()
UUID_RE = re.compile(
    r"[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}"
)


def ruta_para_salida(pdf_path: Path, base: Path) -> str:
    """Devuelve una ruta relativa cuando es posible para no exponer rutas locales."""
    try:
        return str(pdf_path.resolve().relative_to(base.resolve()))
    except ValueError:
        return pdf_path.name


# ---------------------------------------------------------------------------
# Funciones auxiliares
# ---------------------------------------------------------------------------
def pagina_a_imagen(pagina: "fitz.Page", dpi: int) -> np.ndarray:
    """Convierte una página de PDF en una imagen BGR (formato OpenCV)."""
    pix = pagina.get_pixmap(dpi=dpi)
    img = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.width, pix.n)

    if pix.n == 4:
        img = cv2.cvtColor(img, cv2.COLOR_RGBA2BGR)
    elif pix.n == 1:
        img = cv2.cvtColor(img, cv2.COLOR_GRAY2BGR)
    else:
        img = cv2.cvtColor(img, cv2.COLOR_RGB2BGR)
    return img


def detectar_qrs_en_imagen(img: np.ndarray) -> list[str]:
    """Devuelve la lista de textos decodificados de todos los QR encontrados en la imagen."""
    try:
        ok, datos, _, _ = detector.detectAndDecodeMulti(img)
        if ok:
            return [d for d in datos if d]
    except cv2.error:
        pass

    # Fallback a detección de un solo QR (por si detectAndDecodeMulti falla)
    dato, _, _ = detector.detectAndDecode(img)
    return [dato] if dato else []


def extraer_uuid(texto_qr: str) -> str:
    """Extrae el UUID desde un QR del SAT o desde cualquier texto con formato UUID."""
    if not texto_qr:
        return ""

    query = parse_qs(urlparse(texto_qr).query)
    for llave in ("id", "uuid", "UUID"):
        valores = query.get(llave)
        if valores:
            uuid = valores[0].strip()
            if UUID_RE.fullmatch(uuid):
                return uuid.upper()

    match = UUID_RE.search(texto_qr)
    if match:
        return match.group(0).upper()

    # OCR a veces inserta espacios dentro del folio: "20F 1-4DAA".
    texto_compacto = re.sub(r"\s+", "", texto_qr)
    match = UUID_RE.search(texto_compacto)
    return match.group(0).upper() if match else ""


def extraer_uuid_por_ocr(pagina: "fitz.Page", dpi: int = 300) -> str:
    """Intenta leer el folio fiscal con Tesseract cuando el QR no se puede decodificar."""
    if not shutil.which("tesseract"):
        return ""

    img = pagina_a_imagen(pagina, dpi)
    alto, ancho = img.shape[:2]
    recortes = [
        img[: int(alto * 0.35), int(ancho * 0.45) :],  # Folio Fiscal suele estar arriba a la derecha
        img[: int(alto * 0.45), :],
        img[int(alto * 0.45) : int(alto * 0.58), :],  # Cadena original suele incluir el UUID
        img[int(alto * 0.45) : int(alto * 0.72), :],
        img[int(alto * 0.38) : int(alto * 0.70), :],
    ]

    for recorte in recortes:
        if recorte.size == 0:
            continue
        gris = cv2.cvtColor(recorte, cv2.COLOR_BGR2GRAY)
        variantes = [
            gris,
            cv2.resize(gris, None, fx=2, fy=2, interpolation=cv2.INTER_CUBIC),
        ]

        for variante in variantes:
            _, binaria = cv2.threshold(variante, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)

            with tempfile.NamedTemporaryFile(suffix=".png") as archivo_temporal:
                cv2.imwrite(archivo_temporal.name, binaria)
                try:
                    proceso = subprocess.run(
                        ["tesseract", archivo_temporal.name, "stdout", "--psm", "6"],
                        text=True,
                        capture_output=True,
                        check=False,
                        timeout=30,
                    )
                except subprocess.TimeoutExpired:
                    continue

            uuid = extraer_uuid(proceso.stdout)
            if uuid:
                return uuid

    return ""


def crear_resultado(
    pdf_path: Path,
    base: Path,
    pagina: int | str,
    metodo_qr: str,
    estado: str,
) -> dict:
    """Crea una fila para el Excel con nombres consistentes.

    Si se marca 'Encontrado' pero no se pudo extraer un UUID válido del
    contenido del QR, se reclasifica el estado para no dar información
    engañosa (el QR existe, pero no trae un folio fiscal reconocible).
    """
    uuid = extraer_uuid(metodo_qr)
    if estado == "Encontrado" and not uuid:
        estado = "QR encontrado sin UUID reconocible"

    return {
        "nombre_archivo": pdf_path.name,
        "ruta": ruta_para_salida(pdf_path, base),
        "pagina": pagina,
        "uuid": uuid,
        "metodo_qr": metodo_qr,
        "estado": estado,
    }


def buscar_qrs_en_pdf(pdf_path: Path, base: Path) -> list[dict]:
    """
    Busca QR en todas las páginas de un PDF (probando varias resoluciones).

    Siempre revisa el documento completo, incluso si ya encontró un UUID,
    por si hay más de una factura en el mismo archivo. El OCR de respaldo
    solo se activa si NINGUNA página del PDF tuvo un QR detectable, para no
    pagar el costo del OCR cuando el QR normal ya funcionó.
    """
    resultados_pdf: list[dict] = []

    with fitz.open(pdf_path) as doc:
        if doc.needs_pass:
            return [crear_resultado(pdf_path, base, "", "", "PDF protegido con contraseña")]

        if doc.page_count == 0:
            return [crear_resultado(pdf_path, base, "", "", "PDF sin páginas")]

        paginas_con_qr: list[tuple[int, list[str]]] = []

        for num_pagina in range(doc.page_count):
            pagina = doc[num_pagina]
            qrs_encontrados: list[str] = []

            for dpi in DPIS_A_INTENTAR:
                img = pagina_a_imagen(pagina, dpi)
                qrs_encontrados = detectar_qrs_en_imagen(img)
                if qrs_encontrados:
                    break  # ya encontramos algo en esta página, no subir más el DPI

            if qrs_encontrados:
                paginas_con_qr.append((num_pagina, qrs_encontrados))

        if paginas_con_qr:
            for num_pagina, qrs in paginas_con_qr:
                for qr in qrs:
                    resultados_pdf.append(
                        crear_resultado(pdf_path, base, num_pagina + 1, qr, "Encontrado")
                    )
        else:
            # Ninguna página tuvo QR detectable: se intenta OCR en todo el documento.
            for num_pagina in range(doc.page_count):
                uuid_ocr = extraer_uuid_por_ocr(doc[num_pagina])
                if uuid_ocr:
                    resultados_pdf.append({
                        "nombre_archivo": pdf_path.name,
                        "ruta": ruta_para_salida(pdf_path, base),
                        "pagina": num_pagina + 1,
                        "uuid": uuid_ocr,
                        "metodo_qr": "",
                        "estado": "Encontrado por OCR",
                    })

    if not resultados_pdf:
        resultados_pdf.append(crear_resultado(pdf_path, base, "", "", "No encontrado"))

    return resultados_pdf


# ---------------------------------------------------------------------------
# Programa principal
# ---------------------------------------------------------------------------
def parsear_argumentos() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Extrae QR/UUID de PDFs en una carpeta.")
    parser.add_argument(
        "carpeta",
        nargs="?",
        default=".",
        help="Carpeta con los PDFs a procesar (por defecto: carpeta actual).",
    )
    parser.add_argument(
        "-o", "--salida",
        default=str(ARCHIVO_SALIDA),
        help=f"Ruta del Excel de salida (por defecto: {ARCHIVO_SALIDA}).",
    )
    return parser.parse_args()


def main() -> None:
    args = parsear_argumentos()
    carpeta_pdfs = Path(args.carpeta)
    archivo_salida = Path(args.salida)

    if not carpeta_pdfs.exists():
        logger.error("La carpeta '%s' no existe.", carpeta_pdfs)
        return

    pdfs = sorted(carpeta_pdfs.glob("*.pdf"))
    if not pdfs:
        logger.warning("No se encontraron PDFs en '%s'.", carpeta_pdfs)
        return

    resultados: list[dict] = []

    for pdf_path in tqdm(pdfs, desc="Procesando PDFs"):
        try:
            resultados.extend(buscar_qrs_en_pdf(pdf_path, carpeta_pdfs))
        except Exception as e:
            logger.exception("Error procesando %s", pdf_path.name)
            resultados.append(crear_resultado(pdf_path, carpeta_pdfs, "", "", f"Error: {e}"))

    df = pd.DataFrame(resultados, columns=COLUMNAS_EXCEL)
    df.to_excel(archivo_salida, index=False)

    estados_con_dato = ["Encontrado", "Encontrado por OCR"]
    filas_con_uuid = df["estado"].isin(estados_con_dato)
    archivos_con_uuid = df.loc[filas_con_uuid, "nombre_archivo"].nunique()
    uuids_unicos = df.loc[filas_con_uuid, "uuid"].nunique()

    logger.info(
        "Listo. %s/%s archivos con al menos un UUID (%s UUID únicos, %s filas totales). "
        "Archivo generado: %s",
        archivos_con_uuid, len(pdfs), uuids_unicos, len(df), archivo_salida,
    )


if __name__ == "__main__":
    main()
