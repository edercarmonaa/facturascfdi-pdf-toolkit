"""
Extrae UUIDs de CFDI desde PDFs y genera un Excel simple con:
nombre_archivo, uuid, pagina.

Requisitos:
    python3 -m pip install pymupdf opencv-python numpy pandas openpyxl tqdm
    Opcional para QR danados o paginas escaneadas: instalar tesseract en el PATH.

Uso:
    python3 extraer_uuid_paginas.py
    python3 extraer_uuid_paginas.py /ruta/a/pdfs -o resultado_uuid_paginas.xlsx
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
except ImportError:
    def tqdm(iterable, **kwargs):
        return iterable


ARCHIVO_SALIDA = Path("resultado_uuid_paginas.xlsx")
DPIS_A_INTENTAR = [200, 350, 500]
OCR_DPI = 300
COLUMNAS_EXCEL = ["nombre_archivo", "uuid", "pagina"]

UUID_RE = re.compile(
    r"[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}"
)
UUID_OCR_RE = re.compile(
    r"[0-9A-FO]{8}-[0-9A-FO]{4}-[0-9A-FO]{4}-[0-9A-FO]{4}-[0-9A-FO]{12}",
    re.IGNORECASE,
)

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%H:%M:%S",
)
logger = logging.getLogger(__name__)
detector = cv2.QRCodeDetector()


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


def corregir_uuid_ocr(uuid: str) -> str:
    uuid = (uuid or "").upper().replace("O", "0")
    return uuid if UUID_RE.fullmatch(uuid) else ""


def extraer_uuid(texto: str) -> str:
    if not texto:
        return ""

    query = parse_qs(urlparse(texto).query)
    for llave in ("id", "uuid", "UUID"):
        valores = query.get(llave)
        if valores:
            uuid = corregir_uuid_ocr(valores[0].strip())
            if uuid:
                return uuid

    texto_compacto = re.sub(r"\s+", "", texto)
    candidatos = [
        corregir_uuid_ocr(match.group(0))
        for origen in (texto, texto_compacto)
        for match in UUID_RE.finditer(origen)
    ]
    candidatos.extend(
        corregir_uuid_ocr(match.group(0))
        for match in UUID_OCR_RE.finditer(texto_compacto)
    )
    candidatos = [uuid for uuid in candidatos if uuid]
    if not candidatos:
        return ""

    frecuencia: dict[str, int] = {}
    for uuid in candidatos:
        frecuencia[uuid] = frecuencia.get(uuid, 0) + 1
    return max(frecuencia, key=frecuencia.get)


def ocr_imagen(img: np.ndarray) -> str:
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
                    ["tesseract", archivo_temporal.name, "stdout", "--psm", "6"],
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


def extraer_uuid_por_ocr(pagina: "fitz.Page") -> str:
    img = pagina_a_imagen(pagina, OCR_DPI)
    alto, ancho = img.shape[:2]
    recortes = [
        img,
        img[: int(alto * 0.45), :],
        img[: int(alto * 0.35), int(ancho * 0.45) :],
        img[int(alto * 0.38) : int(alto * 0.72), :],
        img[int(alto * 0.45) : int(alto * 0.58), :],
    ]
    texto = "\n".join(ocr_imagen(recorte) for recorte in recortes if recorte.size)
    return extraer_uuid(texto)


def buscar_uuids_en_pdf(pdf_path: Path) -> list[dict]:
    resultados: list[dict] = []
    vistos: set[tuple[str, int]] = set()

    with fitz.open(pdf_path) as doc:
        if doc.needs_pass:
            return resultados

        for indice in range(doc.page_count):
            pagina = doc[indice]
            pagina_num = indice + 1
            qrs_encontrados: list[str] = []

            for dpi in DPIS_A_INTENTAR:
                img = pagina_a_imagen(pagina, dpi)
                qrs_encontrados = detectar_qrs_en_imagen(img)
                if qrs_encontrados:
                    break

            uuids = [extraer_uuid(qr) for qr in qrs_encontrados]
            uuids = [uuid for uuid in uuids if uuid]

            if not uuids:
                uuid_ocr = extraer_uuid_por_ocr(pagina)
                uuids = [uuid_ocr] if uuid_ocr else []

            for uuid in uuids:
                llave = (uuid, pagina_num)
                if llave in vistos:
                    continue
                resultados.append({
                    "nombre_archivo": pdf_path.name,
                    "uuid": uuid,
                    "pagina": pagina_num,
                })
                vistos.add(llave)

    return resultados


def parsear_argumentos() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Extrae nombre de archivo, UUID y pagina de PDFs.")
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
    return parser.parse_args()


def main() -> None:
    args = parsear_argumentos()
    carpeta = Path(args.carpeta)
    salida = Path(args.salida)

    pdfs = sorted(carpeta.glob("*.pdf"))
    if not pdfs:
        logger.warning("No se encontraron PDFs en '%s'.", carpeta)
        return

    filas: list[dict] = []
    for pdf_path in tqdm(pdfs, desc="Extrayendo UUIDs"):
        try:
            filas.extend(buscar_uuids_en_pdf(pdf_path))
        except Exception:
            logger.exception("Error procesando %s", pdf_path.name)

    df = pd.DataFrame(filas, columns=COLUMNAS_EXCEL)
    df.to_excel(salida, index=False)
    logger.info("Listo. %s UUIDs encontrados. Archivo generado: %s", len(df), salida)


if __name__ == "__main__":
    main()
