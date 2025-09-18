from docx import Document
from docx.shared import Cm, Inches
import re

def replace_text(doc: Document, replacements: dict, image_replacements: dict):
    """
    Reemplaza texto e imágenes en un documento Word, manejando texto fragmentado en runs.
    """
    def replace_in_paragraphs(paragraphs):
        for paragraph in paragraphs:
            # Método 1: Consolidar texto completo del párrafo
            full_text = ''.join(run.text for run in paragraph.runs)
            
            # Verificar si hay reemplazos de texto necesarios
            text_replacements_needed = any(key in full_text for key in replacements.keys())
            image_replacements_needed = any(key in full_text for key in image_replacements.keys())
            
            if text_replacements_needed or image_replacements_needed:
                # Aplicar reemplazos de texto
                modified_text = full_text
                for key, value in replacements.items():
                    modified_text = modified_text.replace(key, value)
                
                # Manejar reemplazos de imágenes
                for img_key, img_info in image_replacements.items():
                    if img_key in modified_text:
                        # Dividir el texto donde está la imagen
                        parts = modified_text.split(img_key, 1)
                        
                        # Limpiar todos los runs existentes
                        for run in paragraph.runs[:]:
                            run._element.getparent().remove(run._element)
                        
                        # Agregar texto antes de la imagen
                        if parts[0]:
                            paragraph.add_run(parts[0])
                        
                        # Insertar imagen
                        insert_image_in_paragraph(paragraph, img_info)
                        
                        # Agregar texto después de la imagen
                        if len(parts) > 1 and parts[1]:
                            paragraph.add_run(parts[1])
                        
                        return  # Salir para evitar procesamiento adicional
                
                # Si solo hay reemplazos de texto, actualizar conservando formato
                if text_replacements_needed and not image_replacements_needed:
                    update_paragraph_text_preserving_format(paragraph, modified_text)

    def update_paragraph_text_preserving_format(paragraph, new_text):
        """
        Actualiza el texto del párrafo conservando el formato del primer run.
        """
        if not paragraph.runs:
            paragraph.add_run(new_text)
            return
        
        # Obtener el formato del primer run
        first_run = paragraph.runs[0]
        font_format = {
            'name': first_run.font.name,
            'size': first_run.font.size,
            'bold': first_run.font.bold,
            'italic': first_run.font.italic,
            'underline': first_run.font.underline,
        }
        
        # Limpiar todos los runs
        for run in paragraph.runs[:]:
            run._element.getparent().remove(run._element)
        
        # Crear nuevo run con el texto y formato
        new_run = paragraph.add_run(new_text)
        new_run.font.name = font_format['name']
        if font_format['size']:
            new_run.font.size = font_format['size']
        new_run.font.bold = font_format['bold']
        new_run.font.italic = font_format['italic']
        new_run.font.underline = font_format['underline']

    def replace_in_tables(tables):
        for table in tables:
            for row in table.rows:
                for cell in row.cells:
                    replace_in_paragraphs(cell.paragraphs)

    # Revisar headers y footers
    for section in doc.sections:
        for header in [section.header, section.first_page_header, section.even_page_header]:
            if header:  # Verificar que existe
                replace_in_paragraphs(header.paragraphs)
                replace_in_tables(header.tables)
        
        for footer in [section.footer, section.first_page_footer, section.even_page_footer]:
            if footer:  # Verificar que existe
                replace_in_paragraphs(footer.paragraphs)
                replace_in_tables(footer.tables)

    # Revisar contenido normal del documento
    replace_in_paragraphs(doc.paragraphs)
    replace_in_tables(doc.tables)

def insert_image_in_paragraph(paragraph, image_info):
    """
    Inserta una imagen en un párrafo.
    """
    if isinstance(image_info, str):
        image_info = {
            "path": image_info,
            "width_cm": 3.5,
            "height_cm": 1.5
        }
    
    try:
        run = paragraph.add_run()
        width = Cm(image_info.get("width_cm", 3.5))
        height = Cm(image_info.get("height_cm", 1.5))
        run.add_picture(image_info["path"], width=width, height=height)
    except Exception as e:
        print(f"Error al insertar imagen: {e}")
        # Agregar texto alternativo si la imagen falla
        paragraph.add_run(f"[ERROR: No se pudo cargar imagen {image_info.get('path', 'desconocida')}]")

def replace_text_advanced(doc: Document, replacements: dict, image_replacements: dict):
    """
    Versión avanzada que maneja patrones regex y texto fragmentado.
    """
    def process_paragraph_advanced(paragraph):
        # Obtener todo el texto del párrafo
        full_text = ''.join(run.text for run in paragraph.runs)
        original_text = full_text
        
        # Aplicar reemplazos usando regex para mayor flexibilidad
        for pattern, replacement in replacements.items():
            # Usar regex para patrones más complejos
            if isinstance(pattern, str) and pattern.startswith('{{') and pattern.endswith('}}'):
                # Patrón simple de marcador
                full_text = full_text.replace(pattern, replacement)
            else:
                # Usar regex
                full_text = re.sub(pattern, replacement, full_text)
        
        # Solo actualizar si hubo cambios
        if full_text != original_text:
            # Preservar el formato del párrafo
            if paragraph.runs:
                # Mantener el estilo del párrafo
                paragraph_style = paragraph.style
                
                # Limpiar runs existentes
                for run in paragraph.runs[:]:
                    run._element.getparent().remove(run._element)
                
                # Agregar nuevo contenido
                new_run = paragraph.add_run(full_text)
                
                # Restaurar estilo del párrafo si es necesario
                paragraph.style = paragraph_style

    # Procesar todo el documento
    for paragraph in doc.paragraphs:
        process_paragraph_advanced(paragraph)
    
    # Procesar tablas
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    process_paragraph_advanced(paragraph)

def process_word_file(input_path: str, output_path: str, replacements: dict, image_replacements: dict = None):
    """
    Procesa un archivo Word aplicando reemplazos de texto e imágenes.
    
    Args:
        input_path: Ruta del archivo de entrada
        output_path: Ruta del archivo de salida
        replacements: Diccionario con reemplazos de texto {patrón: reemplazo}
        image_replacements: Diccionario con reemplazos de imágenes {marcador: info_imagen}
    """
    if image_replacements is None:
        image_replacements = {}
    
    try:
        doc = Document(input_path)
        replace_text(doc, replacements, image_replacements)
        doc.save(output_path)
        print(f"Documento procesado exitosamente: {output_path}")
    except Exception as e:
        print(f"Error al procesar el documento: {e}")

# Ejemplo de uso
if __name__ == "__main__":
    # Ejemplo de reemplazos
    replacements = {
        "{{NOMBRE}}": "Juan Pérez",
        "{{FECHA}}": "15 de Septiembre 2024",
        "{{EMPRESA}}": "Mi Empresa S.A."
    }
    
    # Ejemplo de reemplazos de imágenes
    image_replacements = {
        "{{LOGO}}": {
            "path": "logo.png",
            "width_cm": 4.0,
            "height_cm": 2.0
        },
        "{{FIRMA}}": "firma.png"  # Formato simple
    }
    
    # Procesar archivo
    process_word_file(
        "template.docx",
        "output.docx", 
        replacements,
        image_replacements
    )