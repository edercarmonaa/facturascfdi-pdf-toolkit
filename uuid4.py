"""
Para cada PDF de una carpeta, genera un PDF con solo las páginas donde se
detectó un código QR. No hace nada más (no extrae UUID, no hace OCR, no
genera Excel).

Ejemplo: factura.pdf -> pdfs_filtrados/factura_filtrado.pdf

Requisitos:
    python3 -m pip install pymupdf opencv-python numpy tqdm
    (tqdm es opcional: si no está instalado, el script funciona igual pero
    sin barra de progreso)

Uso:
    python3 filtrar_paginas_con_qr.py
    python3 filtrar_paginas_con_qr.py /ruta/a/pdfs -o /ruta/de/salida
"""

import argparse
import time
from pathlib import Path

import cv2
import fitz  # PyMuPDF
import numpy as np

try:
    from tqdm import tqdm
except ImportError:
    def tqdm(iterable, **kwargs):
        return iterable

DPIS_A_INTENTAR = [200, 350, 500]  # sube la resolución si no encuentra nada

detector = cv2.QRCodeDetector()


def pagina_a_imagen(pagina: "fitz.Page", dpi: int) -> np.ndarray:
    pix = pagina.get_pixmap(dpi=dpi, alpha=False)
    img = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.width, pix.n)
    if img.shape[2] == 1:
        return cv2.cvtColor(img, cv2.COLOR_GRAY2BGR)
    return cv2.cvtColor(img, cv2.COLOR_RGB2BGR)


def pagina_tiene_qr(pagina: "fitz.Page") -> bool:
    for dpi in DPIS_A_INTENTAR:
        img = pagina_a_imagen(pagina, dpi)

        ok, datos, _, _ = detector.detectAndDecodeMulti(img)
        if ok and any(datos):
            return True

        dato, _, _ = detector.detectAndDecode(img)
        if dato:
            return True

    return False


def filtrar_pdf(pdf_path: Path, carpeta_salida: Path) -> None:
    inicio = time.perf_counter()

    with fitz.open(pdf_path) as doc:
        paginas_con_qr = [
            i
            for i in tqdm(
                range(doc.page_count),
                desc=pdf_path.name,
                unit="pág",
                leave=False,
            )
            if pagina_tiene_qr(doc[i])
        ]

        segundos = time.perf_counter() - inicio

        if not paginas_con_qr:
            print(f"{pdf_path.name}: ninguna página con QR, se omite ({segundos:.1f}s)")
            return

        doc.select(paginas_con_qr)
        carpeta_salida.mkdir(parents=True, exist_ok=True)
        ruta_salida = carpeta_salida / f"{pdf_path.stem}_filtrado.pdf"
        doc.save(ruta_salida)
        print(
            f"{pdf_path.name}: {len(paginas_con_qr)} página(s) con QR -> {ruta_salida} "
            f"({segundos:.1f}s)"
        )


def main() -> None:
    parser = argparse.ArgumentParser(description="Separa las páginas con QR de cada PDF en un PDF filtrado.")
    parser.add_argument("carpeta", nargs="?", default=".", help="Carpeta con los PDFs de entrada.")
    parser.add_argument("-o", "--salida", default="pdfs_filtrados", help="Carpeta de salida.")
    args = parser.parse_args()

    carpeta = Path(args.carpeta)
    carpeta_salida = Path(args.salida)

    pdfs = sorted(carpeta.glob("*.pdf"))
    if not pdfs:
        print(f"No se encontraron PDFs en '{carpeta}'.")
        return

    inicio_total = time.perf_counter()

    for pdf_path in tqdm(pdfs, desc="PDFs", unit="pdf"):
        filtrar_pdf(pdf_path, carpeta_salida)

    total_segundos = time.perf_counter() - inicio_total
    minutos, segundos = divmod(total_segundos, 60)
    print(f"\nListo: {len(pdfs)} PDF(s) procesados en {int(minutos)}m {segundos:.1f}s.")


if __name__ == "__main__":
    main()