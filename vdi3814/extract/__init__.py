"""Extraktions-Backends, nach Genauigkeit gestaffelt.

1. excel    - .xlsx/.xlsm/.xls: exakte Zellwerte, kein Erkennungsrisiko
2. pdf_text - PDF mit Textebene (aus Excel/CAD gedruckt): exakte Wortkoordinaten
3. vision   - Scans, PNG/JPEG, bildbasierte PDFs: lokales Vision-Language-Modell

Alle drei liefern dasselbe Zwischenformat (RawTable), das die Pipeline
anschliessend gegen die Spalten-Profile mappt.
"""

from .base import RawTable, ExtractionMode
from .excel_extractor import extract_excel, is_excel
from .ocr_extractor import classify_image, extract_ocr
from .pdftext_extractor import extract_pdf_text, pdf_has_text_layer

__all__ = [
    "RawTable",
    "ExtractionMode",
    "extract_excel",
    "is_excel",
    "extract_pdf_text",
    "pdf_has_text_layer",
    "extract_ocr",
    "classify_image",
]
