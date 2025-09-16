from openpyxl import load_workbook
from openpyxl.drawing.image import Image
import os


def replace_text_and_images(workbook, replacements: dict, image_replacements: dict):
    for ws in workbook.worksheets:  # recorrer todas las hojas
        for row in ws.iter_rows():
            for cell in row:
                if cell.value and isinstance(cell.value, str):
                    text = cell.value

                    # Reemplazo de texto
                    for key, value in replacements.items():
                        if key in text:
                            text = text.replace(key, value)

                    # Reemplazo de imágenes
                    for img_key, img_info in image_replacements.items():
                        if img_key in text:
                            text = text.replace(img_key, "")
                            if isinstance(img_info, str):
                                img_info = {
                                    "path": img_info,
                                    "width": None,
                                    "height": None
                                }
                            if os.path.exists(img_info["path"]):
                                img = Image(img_info["path"])
                                # Opcional: ajustar tamaño si se pasó en config
                                if img_info.get("width"):
                                    img.width = img_info["width"]
                                if img_info.get("height"):
                                    img.height = img_info["height"]

                                ws.add_image(img, cell.coordinate)

                    cell.value = text if text else None


def process_excel_file(input_path: str, output_path: str, replacements: dict, image_replacements: dict):
    wb = load_workbook(input_path)
    replace_text_and_images(wb, replacements, image_replacements)
    wb.save(output_path)
