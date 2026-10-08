# CFDI PDF Toolkit

Herramientas de linea de comandos para procesar lotes de facturas CFDI en PDF. El proyecto resuelve una tarea comun en conciliacion fiscal: localizar codigos QR o folios fiscales dentro de muchos PDF, extraer UUID/RFC/total cuando es posible y generar archivos Excel para revision o consulta del estatus ante el SAT.

## Que problema resuelve

Cuando se reciben decenas o cientos de facturas en PDF, validar manualmente cada UUID es lento y propenso a errores. Estos scripts automatizan tres pasos:

- Detectar paginas con QR dentro de PDFs grandes.
- Extraer UUID de CFDI desde QR o, como respaldo, mediante OCR con Tesseract.
- Consultar el servicio publico del SAT con la expresion impresa del CFDI y guardar el resultado en Excel.

## Privacidad

Los PDFs de facturas, Excels generados y archivos filtrados pueden contener RFC, UUID, importes, nombres fiscales y otros datos sensibles. Por eso el repositorio ignora por defecto `*.pdf`, `*.xlsx`, `resultado_*.xlsx`, `*_filtrado.pdf` y `pdfs_filtrados/`.

No subas facturas reales ni resultados generados a un repositorio publico. Usa datos ficticios si necesitas agregar ejemplos.

## Requisitos

- Python 3.10 o superior.
- Dependencias de Python listadas en `requirements.txt`.
- Opcional: Tesseract OCR instalado y disponible en `PATH` para leer PDFs escaneados o QR danados.

Instalacion sugerida:

```bash
python3 -m venv .venv
source .venv/bin/activate
python3 -m pip install -r requirements.txt
```

## Uso

Extraer QR y UUID de PDFs:

```bash
python3 extraer_qr_pdfs.py /ruta/a/pdfs -o resultado_qr_pdfs.xlsx
```

Extraer solo UUID y pagina:

```bash
python3 extraer_uuid_paginas.py /ruta/a/pdfs -o resultado_uuid_paginas.xlsx
```

Version paralela para lotes grandes:

```bash
python3 uuid3.py /ruta/a/pdfs -o resultado_uuid_paginas.xlsx --workers 4 --ocr-workers 2
```

Consultar estatus CFDI ante SAT:

```bash
python3 consultar_estatus_sat.py /ruta/a/pdfs -o resultado_sat_cfdi.xlsx
```

Solo extraer datos, sin llamar al SAT:

```bash
python3 consultar_estatus_sat.py /ruta/a/pdfs -o resultado_sat_cfdi.xlsx --sin-consultar
```

Generar PDFs filtrados con paginas que contienen QR:

```bash
python3 uuid4.py /ruta/a/pdfs -o pdfs_filtrados
```

## Salidas

Los Excels generados pueden incluir:

- nombre de archivo;
- pagina;
- UUID;
- RFC emisor y receptor;
- total;
- fragmento `fe`;
- estado de extraccion;
- respuesta normalizada del SAT.

Por privacidad, los archivos de salida estan ignorados por Git.

## Notas de seguridad

- No hay credenciales requeridas para usar el servicio publico de consulta CFDI del SAT.
- No se deben versionar `.env`, PDFs reales, Excels generados, claves privadas ni credenciales.
- Si algun PDF o Excel real ya fue publicado por error, eliminalo del historial Git y considera los datos expuestos.

## Estado del proyecto

Este repositorio esta preparado para revision antes de publicacion: incluye `.gitignore`, `.env.example`, `requirements.txt` y documentacion. Antes de hacerlo publico, revisa manualmente que no agregues archivos fiscales reales con `git add`.

## Licencia

Este proyecto se publica bajo la licencia MIT. Consulta [LICENSE](LICENSE).
