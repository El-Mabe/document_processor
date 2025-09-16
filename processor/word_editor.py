from docx import Document
from docx.shared import Cm, Inches


def replace_text(doc: Document, replacements: dict, image_replacements: dict):
    def replace_in_paragraphs(paragraphs):
        for paragraph in paragraphs:
            for run in paragraph.runs:
                text = run.text

                # Reemplazos de texto manteniendo el estilo
                for key, value in replacements.items():
                    if key in text:
                        run.text = text.replace(key, value)

                # Reemplazos de imágenes
                for img_key, img_info in image_replacements.items():
                    if img_key in run.text:
                        run.text = ""  # limpiar solo el marcador
                        insert_image_in_paragraph(paragraph, img_info)

    def replace_in_tables(tables):
        for table in tables:
            for row in table.rows:
                for cell in row.cells:
                    replace_in_paragraphs(cell.paragraphs)

    # Revisar headers y footers
    for section in doc.sections:
        for header in [section.header, section.first_page_header, section.even_page_header]:
            replace_in_paragraphs(header.paragraphs)
            replace_in_tables(header.tables)

        for footer in [section.footer, section.first_page_footer, section.even_page_footer]:
            replace_in_paragraphs(footer.paragraphs)
            replace_in_tables(footer.tables)

    # Revisar contenido normal del documento
    replace_in_paragraphs(doc.paragraphs)
    replace_in_tables(doc.tables)


def insert_image_in_paragraph(paragraph, image_info):
    if isinstance(image_info, str):
        image_info = {
            "path": image_info,
            "width_cm": 3.5,
            "height_cm": 1.5
        }
    run = paragraph.add_run()
    width = Cm(image_info.get("width_cm", 3.5))
    height = Cm(image_info.get("height_cm", 1.5))
    run.add_picture(image_info["path"], width=width, height=height)


def process_word_file(input_path: str, output_path: str, replacements: dict, image_replacements: dict):
    doc = Document(input_path)
    replace_text(doc, replacements, image_replacements)
    doc.save(output_path)
