"""
Extrae códigos QR de todos los PDFs en una carpeta y guarda los resultados en Excel.

Requisitos:
    pip install pymupdf opencv-python pandas openpyxl tqdm
"""

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
CARPETA_PDFS = Path(".")
ARCHIVO_SALIDA = Path("resultado_qr_pdfs.xlsx")
DPIS_A_INTENTAR = [200, 350, 500]  # sube la resolución si no encuentra nada
DETENER_EN_PRIMER_QR = False  # False = revisa todas las páginas del PDF
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
        gris = cv2.cvtColor(recorte, cv2.COLOR_BGR2GRAY)
        variantes = [
            gris,
            cv2.resize(gris, None, fx=2, fy=2, interpolation=cv2.INTER_CUBIC),
        ]

        for variante in variantes:
            _, binaria = cv2.threshold(variante, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)

            with tempfile.NamedTemporaryFile(suffix=".png") as archivo_temporal:
                cv2.imwrite(archivo_temporal.name, binaria)
                proceso = subprocess.run(
                    ["tesseract", archivo_temporal.name, "stdout", "--psm", "6"],
                    text=True,
                    capture_output=True,
                    check=False,
                )

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
    """Crea una fila para el Excel con nombres consistentes."""
    return {
        "nombre_archivo": pdf_path.name,
        "ruta": ruta_para_salida(pdf_path, base),
        "pagina": pagina,
        "uuid": extraer_uuid(metodo_qr),
        "metodo_qr": metodo_qr,
        "estado": estado,
    }


def buscar_qrs_en_pdf(pdf_path: Path, base: Path) -> list[dict]:
    """Busca QR en todas las páginas de un PDF, probando varias resoluciones."""
    resultados_pdf = []

    with fitz.open(pdf_path) as doc:
        if doc.page_count == 0:
            return [crear_resultado(pdf_path, base, "", "", "PDF sin páginas")]

        for num_pagina in range(doc.page_count):
            pagina = doc[num_pagina]
            qrs_encontrados: list[str] = []

            for dpi in DPIS_A_INTENTAR:
                img = pagina_a_imagen(pagina, dpi)
                qrs_encontrados = detectar_qrs_en_imagen(img)
                if qrs_encontrados:
                    break  # ya encontramos algo, no hace falta subir más la resolución

            if qrs_encontrados:
                for qr in qrs_encontrados:
                    resultados_pdf.append(
                        crear_resultado(pdf_path, base, num_pagina + 1, qr, "Encontrado")
                    )
                if DETENER_EN_PRIMER_QR:
                    break
            else:
                uuid_ocr = extraer_uuid_por_ocr(pagina)
                if uuid_ocr:
                    resultados_pdf.append({
                        "nombre_archivo": pdf_path.name,
                        "ruta": ruta_para_salida(pdf_path, base),
                        "pagina": num_pagina + 1,
                        "uuid": uuid_ocr,
                        "metodo_qr": "OCR Folio Fiscal",
                        "estado": "Encontrado por OCR",
                    })

    if not resultados_pdf:
        resultados_pdf.append(crear_resultado(pdf_path, base, "", "", "No encontrado"))

    return resultados_pdf


# ---------------------------------------------------------------------------
# Programa principal
# ---------------------------------------------------------------------------
def main() -> None:
    if not CARPETA_PDFS.exists():
        logger.error("La carpeta '%s' no existe.", CARPETA_PDFS)
        return

    pdfs = sorted(CARPETA_PDFS.glob("*.pdf"))
    if not pdfs:
        logger.warning("No se encontraron PDFs en '%s'.", CARPETA_PDFS)
        return

    resultados: list[dict] = []

    for pdf_path in tqdm(pdfs, desc="Procesando PDFs"):
        try:
            resultados.extend(buscar_qrs_en_pdf(pdf_path, CARPETA_PDFS))
        except Exception as e:
            logger.exception("Error procesando %s", pdf_path.name)
            resultados.append(crear_resultado(pdf_path, CARPETA_PDFS, "", "", f"Error: {e}"))

    df = pd.DataFrame(resultados, columns=COLUMNAS_EXCEL)
    df.to_excel(ARCHIVO_SALIDA, index=False)

    encontrados = df["estado"].isin(["Encontrado", "Encontrado por OCR"])
    archivos_con_resultados = df.loc[encontrados, "nombre_archivo"].nunique()
    facturas_encontradas = encontrados.sum()
    qrs_decodificados = (df["estado"] == "Encontrado").sum()
    recuperados_por_ocr = (df["estado"] == "Encontrado por OCR").sum()
    logger.info(
        "Listo. %s/%s archivos con resultados, %s facturas: %s por QR y %s por OCR. Archivo generado: %s",
        archivos_con_resultados,
        len(pdfs),
        facturas_encontradas,
        qrs_decodificados,
        recuperados_por_ocr,
        ARCHIVO_SALIDA,
    )


if __name__ == "__main__":
    main()
