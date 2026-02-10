
from pathlib import Path
import fitz  # PyMuPDF

document_dir = Path(r"C:\Users\GreenhouseProduction\Downloads\us customs test")
pdf_in  = document_dir / "Blank Template copy - Copy.pdf"
pdf_out = document_dir / "output.pdf"

with fitz.open(pdf_in) as doc:
    target_page = doc[0]

    font_rgb = (0, 137, 210)
    font_colour = tuple(value / 255 for value in font_rgb)

    for page_num, page in enumerate(doc):
        for indx, field in enumerate(page.widgets()):
            # index checking
            if field.field_type == fitz.PDF_WIDGET_TYPE_TEXT: #if field is text field
                field.field_value = '{0}'.format(indx) #insert text into the field
                field.update()

    doc.save(pdf_out)