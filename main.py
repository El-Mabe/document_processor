import os
import json
import time
from processor.word_editor import process_word_file

def main():
    start_time = time.time()  # ⏱️ Comienza el conteo
    with open('config.json', 'r', encoding='utf-8') as f:
        config = json.load(f)

    input_dir = "input_docs"
    output_dir = "output_docs"
    os.makedirs(output_dir, exist_ok=True)

    for filename in os.listdir(input_dir):
        if filename.endswith(".docx"):
            input_path = os.path.join(input_dir, filename)
            output_path = os.path.join(output_dir, filename)
            print(f"Procesando {filename}...")

            process_word_file(
                input_path=input_path,
                output_path=output_path,
                replacements = config.get("replacements", {}),
                image_replacements=config.get("image_replacements", {})
                # image_path=config["image_path"],
                # image_keyword=config["image_keyword"]
            )

    end_time = time.time()  # ⏱️ Fin del conteo
    duration = end_time - start_time

    print("✔️ Proceso completado.")
    print(f"✔️ Todos los documentos han sido procesados.")
    print(f"🕒 Tiempo total de ejecución: {duration:.2f} segundos.")

if __name__ == "__main__":
    main()
