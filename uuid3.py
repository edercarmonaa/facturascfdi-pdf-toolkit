"""
Extrae UUIDs de CFDI desde PDFs y genera un Excel simple con:
nombre_archivo, uuid, pagina.

La lógica de detección es EXACTAMENTE la misma que el original (todas las
páginas se revisan con QR y, si falla, con OCR sobre 5 recortes x 2
variantes cada uno) -- no se omite ninguna página por heurísticas.

Cambios de esta versión, todos orientados a bajar consumo de CPU/RAM y
subir velocidad, sin tocar QUÉ se calcula:

    1. OMP_THREAD_LIMIT=1 al invocar tesseract. Tesseract usa OpenMP y por
       defecto intenta usar TODOS los núcleos en cada llamada. Si ya estás
       corriendo varias llamadas en paralelo (hilos) dentro de varios
       procesos en paralelo (PDFs), sin este límite terminas con decenas
       de hilos peleando por los mismos núcleos: más lento y más RAM, no
       menos. Con el límite, cada llamada usa 1 hilo y el paralelismo real
       lo controla tu --workers/--ocr-workers.
    2. Los valores por defecto de --workers y --ocr-workers ahora se
       calculan para que su producto ronde el número de núcleos
       disponibles, en vez de 10 hilos fijos por página multiplicados por
       N procesos.
    3. Los procesos worker se reciclan cada cierto número de PDFs
       (max_tasks_per_child, si tu Python lo soporta) para que la RAM
       acumulada por PyMuPDF/OpenCV en PDFs muy grandes se libere en vez
       de quedarse creciendo durante toda la corrida.
    4. Se liberan explícitamente los buffers de imagen grandes (500 dpi
       puede ser >20 megapíxeles) en cuanto ya no se necesitan, en vez de
       esperar a que el recolector de basura decida hacerlo.
    5. Tesseract sigue recibiendo la imagen por stdin (sin disco) y el
       render de página sigue pidiéndose sin canal alfa.

Requisitos:
    python3 -m pip install pymupdf opencv-python numpy pandas openpyxl tqdm
    Opcional para QR dañados o páginas escaneadas: instalar tesseract en el PATH.

Uso:
    python3 extraer_uuid_paginas.py
    python3 extraer_uuid_paginas.py /ruta/a/pdfs -o resultado_uuid_paginas.xlsx
    python3 extraer_uuid_paginas.py /ruta/a/pdfs --workers 4 --ocr-workers 2

Nota sobre --workers/--ocr-workers: su PRODUCTO es cuánta CPU real vas a
usar a la vez. Súbelos por separado solo si sabes que te sobran núcleos;
si no, vas a ser más lento por contención, no más rápido.
"""

import argparse
import gc
import logging
import os
import re
import shutil
import subprocess
import sys
from concurrent.futures import ProcessPoolExecutor, ThreadPoolExecutor, as_completed
from functools import lru_cache
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
    def tqdm(iterable=None, **kwargs):
        return iterable if iterable is not None else _DummyBar()

    class _DummyBar:
        def update(self, *a, **k):
            pass

        def close(self):
            pass


ARCHIVO_SALIDA = Path("resultado_uuid_paginas3.xlsx")
DPIS_A_INTENTAR = [200, 350, 500]
OCR_DPI = 300
COLUMNAS_EXCEL = ["nombre_archivo", "uuid", "pagina"]

NUCLEOS = os.cpu_count() or 4
# Producto workers x ocr_workers pensado para rondar NUCLEOS sin pasarse.
DEFAULT_WORKERS = max(1, NUCLEOS // 2)
DEFAULT_OCR_WORKERS = max(1, NUCLEOS // DEFAULT_WORKERS)

# Cada proceso worker se recicla tras procesar esta cantidad de PDFs
# (evita que la RAM de PyMuPDF/OpenCV crezca sin límite en corridas largas).
PDFS_POR_RECICLO_DE_WORKER = 25

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

# Variables de entorno para las llamadas a tesseract: 1 hilo interno por
# llamada, el paralelismo real lo maneja nuestro ThreadPoolExecutor.
_ENV_TESSERACT = {**os.environ, "OMP_THREAD_LIMIT": "1", "OMP_NUM_THREADS": "1"}


@lru_cache(maxsize=1)
def hay_tesseract() -> bool:
    """Se evalúa una sola vez por proceso en vez de en cada recorte."""
    return shutil.which("tesseract") is not None


def pagina_a_imagen(pagina: "fitz.Page", dpi: int) -> np.ndarray:
    # alpha=False evita traer un canal alfa que de todos modos se descartaba
    # en la conversión de color (mismo contenido visual, menos trabajo).
    pix = pagina.get_pixmap(dpi=dpi, alpha=False)
    img = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.width, pix.n).copy()
    pix = None  # libera el buffer nativo de MuPDF apenas se copió a numpy

    if img.shape[2] == 1:
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


# ---------------------------------------------------------------------------
# OCR
# ---------------------------------------------------------------------------
def _tesseract_stdin(png_bytes: bytes) -> str:
    """Corre tesseract pasando la imagen por stdin (sin archivo temporal en disco)."""
    try:
        proceso = subprocess.run(
            ["tesseract", "stdin", "stdout", "--psm", "6"],
            input=png_bytes,
            capture_output=True,
            check=False,
            timeout=30,
            env=_ENV_TESSERACT,
        )
    except subprocess.TimeoutExpired:
        return ""
    return proceso.stdout.decode("utf-8", errors="ignore") if proceso.stdout else ""


def _preparar_variantes_binarizadas(img: np.ndarray) -> list[bytes]:
    """Genera las mismas 2 variantes (gris y gris x2) ya binarizadas y como PNG en memoria."""
    gris = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    variantes = [
        gris,
        cv2.resize(gris, None, fx=2, fy=2, interpolation=cv2.INTER_CUBIC),
    ]
    pngs = []
    for variante in variantes:
        _, binaria = cv2.threshold(variante, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
        ok, buffer = cv2.imencode(".png", binaria)
        if ok:
            pngs.append(buffer.tobytes())
    return pngs


def extraer_uuid_por_ocr(pagina: "fitz.Page", ocr_workers: int) -> str:
    """
    Idéntico resultado al original: 5 recortes x 2 variantes = hasta 10
    llamadas a tesseract, cuyo texto se concatena y se vota por UUID más
    frecuente. La diferencia es que las llamadas corren en paralelo con un
    límite de concurrencia (ocr_workers) y cada una usa 1 solo hilo interno.
    """
    if not hay_tesseract():
        return ""

    img = pagina_a_imagen(pagina, OCR_DPI)
    alto, ancho = img.shape[:2]
    recortes = [
        img,
        img[: int(alto * 0.45), :],
        img[: int(alto * 0.35), int(ancho * 0.45):],
        img[int(alto * 0.38): int(alto * 0.72), :],
        img[int(alto * 0.45): int(alto * 0.58), :],
    ]

    tareas_png: list[bytes] = []
    for recorte in recortes:
        if recorte.size:
            tareas_png.extend(_preparar_variantes_binarizadas(recorte))
    img = None  # ya se extrajeron los recortes que hacían falta

    if not tareas_png:
        return ""

    textos: list[str] = []
    with ThreadPoolExecutor(max_workers=max(1, min(ocr_workers, len(tareas_png)))) as pool:
        futuros = [pool.submit(_tesseract_stdin, png) for png in tareas_png]
        for futuro in futuros:
            texto = futuro.result()
            if texto:
                textos.append(texto)

    return extraer_uuid("\n".join(textos))


# ---------------------------------------------------------------------------
# Por PDF
# ---------------------------------------------------------------------------
def buscar_uuids_en_pdf(pdf_path: Path, ocr_workers: int) -> list[dict]:
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
                img = None  # libera cada intento de render apenas se evaluó
                if qrs_encontrados:
                    break

            uuids = [extraer_uuid(qr) for qr in qrs_encontrados]
            uuids = [uuid for uuid in uuids if uuid]

            if not uuids:
                uuid_ocr = extraer_uuid_por_ocr(pagina, ocr_workers=ocr_workers)
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

        # PDFs grandes pueden acumular bastante memoria en OpenCV/MuPDF
        # antes de que el GC normal la reclame; se fuerza al terminar el PDF.
        gc.collect()

    return resultados


def _procesar_pdf_worker(args: tuple[Path, int]) -> tuple[Path, list[dict], str | None]:
    """Wrapper picklable para correr en un proceso aparte."""
    pdf_path, ocr_workers = args
    try:
        filas = buscar_uuids_en_pdf(pdf_path, ocr_workers=ocr_workers)
        return pdf_path, filas, None
    except Exception as e:  # noqa: BLE001 - se reporta en el proceso principal
        return pdf_path, [], str(e)


# ---------------------------------------------------------------------------
# Programa principal
# ---------------------------------------------------------------------------
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
    parser.add_argument(
        "--workers",
        type=int,
        default=DEFAULT_WORKERS,
        help=(
            f"Cuántos PDFs procesar en paralelo (procesos). "
            f"Por defecto: {DEFAULT_WORKERS} (para {NUCLEOS} núcleos detectados)."
        ),
    )
    parser.add_argument(
        "--ocr-workers",
        type=int,
        default=DEFAULT_OCR_WORKERS,
        help=(
            f"Cuántas llamadas a tesseract lanzar en paralelo por página (hilos). "
            f"Por defecto: {DEFAULT_OCR_WORKERS}. Evita que --workers x --ocr-workers "
            f"supere tus núcleos disponibles."
        ),
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

    if args.workers * args.ocr_workers > NUCLEOS * 2:
        logger.warning(
            "--workers (%s) x --ocr-workers (%s) = %s procesos/hilos activos a la vez, "
            "muy por encima de tus %s núcleos. Puede ser más lento por contención, no más rápido.",
            args.workers, args.ocr_workers, args.workers * args.ocr_workers, NUCLEOS,
        )

    filas: list[dict] = []
    tareas = [(pdf_path, args.ocr_workers) for pdf_path in pdfs]

    if args.workers <= 1 or len(pdfs) == 1:
        # Un solo PDF o un solo worker: no vale la pena la sobrecarga de procesos.
        for pdf_path, ocr_workers in tqdm(tareas, desc="Extrayendo UUIDs"):
            _, resultado, error = _procesar_pdf_worker((pdf_path, ocr_workers))
            if error:
                logger.error("Error procesando %s: %s", pdf_path.name, error)
            filas.extend(resultado)
    else:
        kwargs_pool = {"max_workers": args.workers}
        try:
            # Recicla procesos worker cada N tareas (Python 3.11+). Si tu
            # versión no lo soporta, se ignora y se sigue igual.
            ProcessPoolExecutor(max_workers=1, max_tasks_per_child=1).shutdown()
            kwargs_pool["max_tasks_per_child"] = PDFS_POR_RECICLO_DE_WORKER
        except TypeError:
            pass

        resultados_por_pdf: dict[Path, list[dict]] = {}
        with ProcessPoolExecutor(**kwargs_pool) as pool:
            futuros = {pool.submit(_procesar_pdf_worker, tarea): tarea[0] for tarea in tareas}
            barra = tqdm(total=len(futuros), desc="Extrayendo UUIDs")
            for futuro in as_completed(futuros):
                pdf_path, resultado, error = futuro.result()
                if error:
                    logger.error("Error procesando %s: %s", pdf_path.name, error)
                resultados_por_pdf[pdf_path] = resultado
                barra.update(1)
            barra.close()

        # Se reordena según el orden original de la carpeta (glob ordenado),
        # igual que si se hubiera procesado secuencialmente.
        for pdf_path in pdfs:
            filas.extend(resultados_por_pdf.get(pdf_path, []))

    df = pd.DataFrame(filas, columns=COLUMNAS_EXCEL)
    df.to_excel(salida, index=False)
    logger.info("Listo. %s UUIDs encontrados. Archivo generado: %s", len(df), salida)


if __name__ == "__main__":
    main()