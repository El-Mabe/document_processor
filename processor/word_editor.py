from docx import Document
from docx.shared import Cm
from docx.shared import Inches
import os


# def replace_text(doc: Document, replacements: dict):
#     def replace_in_paragraphs(paragraphs):
#         for paragraph in paragraphs:
#             for run in paragraph.runs:
#                 for key, value in replacements.items():
#                     if key in run.text:
#                         run.text = run.text.replace(key, value)

#     def replace_in_tables(tables):
#         for table in tables:
#             for row in table.rows:
#                 for cell in row.cells:
#                     replace_in_paragraphs(cell.paragraphs)

#     # 1. Cuerpo del documento
#     replace_in_paragraphs(doc.paragraphs)
#     replace_in_tables(doc.tables)

#     # 2. Encabezados y pies de página (todas las variantes)
#     for section in doc.sections:
#         # --- Encabezados ---
#         for header in [section.header, section.first_page_header, section.even_page_header]:
#             replace_in_paragraphs(header.paragraphs)
#             replace_in_tables(header.tables)

#         # --- Pies de página ---
#         for footer in [section.footer, section.first_page_footer, section.even_page_footer]:
#             replace_in_paragraphs(footer.paragraphs)
#             replace_in_tables(footer.tables)

def replace_text(doc: Document, replacements: dict, image_replacements: dict):
    def replace_in_paragraphs(paragraphs):
        for paragraph in paragraphs:
            full_text = "".join(run.text for run in paragraph.runs)

        # Reemplazar texto
            for key, value in replacements.items():
                if key in full_text:
                    full_text = full_text.replace(key, value)

        # Detectar e insertar imagen
            for img_key, img_info in image_replacements.items():
                if img_key in full_text:
                    # Borrar el contenido del párrafo
                    for run in paragraph.runs:
                        run.text = ""
                    # Insertar imagen
                    insert_image_in_paragraph(paragraph, img_info)
                    full_text = full_text.replace(img_key, "")

        # Finalmente, establecer el nuevo texto (si no es imagen)
            if full_text.strip():
                paragraph.clear()  # Borra el contenido
                paragraph.add_run(full_text)
    #     for paragraph in paragraphs:
    #         for i, run in enumerate(paragraph.runs):
    #             for key, value in replacements.items():
    #                 if key in run.text:
    #                     run.text = run.text.replace(key, value)
    #             for img_key, img_path in image_replacements.items():
    #                 if img_key in run.text and os.path.exists(img_path):
    #                     insert_image_in_run(paragraph, i, img_path)
    
  

    def replace_in_tables(tables):
        for table in tables:
            for row in table.rows:
                for cell in row.cells:
                    replace_in_paragraphs(cell.paragraphs)

    for section in doc.sections:
        for header in [section.header, section.first_page_header, section.even_page_header]:
            replace_in_paragraphs(header.paragraphs)
            replace_in_tables(header.tables)

        for footer in [section.footer, section.first_page_footer, section.even_page_footer]:
            replace_in_paragraphs(footer.paragraphs)
            replace_in_tables(footer.tables)

    replace_in_paragraphs(doc.paragraphs)
    replace_in_tables(doc.tables)


def insert_image_in_run(paragraph, run_index, image_path):
    run = paragraph.runs[run_index]
    # Elimina el texto de la etiqueta
    run.text = ""
    # Inserta la imagen justo después de ese run
    run = paragraph.add_run()
    run.add_picture(image_path, width=Inches(1.2))  # Puedes ajustar el tamaño
                                
# def insert_image_at_paragraph(doc: Document, keyword: str, image_path: str):
#     for i, paragraph in enumerate(doc.paragraphs):
#         if keyword in paragraph.text:
#             # Reemplazar palabra clave por imagen
#             paragraph.clear()
#             run = paragraph.add_run()
#             run.add_picture(image_path, width=Inches(2.0))  # Cambia el tamaño según necesites
#             break

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
    # insert_image_at_paragraph(doc, image_keyword, image_path)
    doc.save(output_path)
