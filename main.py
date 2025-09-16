import os
import json
import time
import tkinter as tk
from tkinter import filedialog, messagebox
from processor.word_editor import process_word_file
from processor.excel_editor import process_excel_file 


from processor.word_editor import process_word_file
from processor.excel_editor import process_excel_file  # 👈 importar


def ejecutar_proceso(valores_usuario, imagenes_usuario, carpeta_entrada_var, carpeta_salida_var):
    start_time = time.time()

    # Obtener valores de reemplazo
    replacements = {key: entrada.get() for key, entrada in valores_usuario.items()}

    # Obtener imágenes
    image_replacements = {}
    for key, info in imagenes_usuario.items():
        ruta = info["path_var"].get()
        if ruta:
            image_replacements[key] = ruta

    carpeta_entrada = carpeta_entrada_var.get()
    carpeta_salida = carpeta_salida_var.get()

    if not carpeta_entrada:
        messagebox.showwarning("Advertencia", "Debe seleccionar una carpeta de entrada.")
        return
    if not carpeta_salida:
        messagebox.showwarning("Advertencia", "Debe seleccionar una carpeta de salida.")
        return

    procesados = 0
    errores = []  # 👈 lista de errores

    for dirpath, _, filenames in os.walk(carpeta_entrada):
        for filename in filenames:
            if filename.lower().endswith((".docx", ".xlsx")):
                input_path = os.path.join(dirpath, filename)

                # Mantener estructura de carpetas
                relative_path = os.path.relpath(dirpath, carpeta_entrada)
                output_dir = os.path.join(carpeta_salida, relative_path)
                os.makedirs(output_dir, exist_ok=True)

                output_path = os.path.join(output_dir, filename)
                relative_input = os.path.relpath(input_path, carpeta_entrada)

                try:
                    print(f"Procesando {relative_input} -> {output_path}")

                    if filename.lower().endswith(".docx"):
                        process_word_file(
                            input_path=input_path,
                            output_path=output_path,
                            replacements=replacements,
                            image_replacements=image_replacements
                        )

                    elif filename.lower().endswith(".xlsx"):
                        process_excel_file(
                            input_path=input_path,
                            output_path=output_path,
                            replacements=replacements,
                            image_replacements=image_replacements
                        )

                    procesados += 1

                except Exception as e:
                    tipo = "Word" if filename.lower().endswith(".docx") else "Excel"
                    error_msg = f"[{tipo}] {relative_input} -> {str(e)}"
                    print(f"⚠️ Error en {error_msg}")
                    errores.append(error_msg)

    end_time = time.time()
    duration = end_time - start_time

    # Mensaje final
    if errores:
        log_path = os.path.join(carpeta_salida, "errores.log")
        with open(log_path, "w", encoding="utf-8") as f:
            f.write("=== Archivos con errores ===\n\n")
            for err in errores:
                f.write(err + "\n")
            f.write(f"\nTotal errores: {len(errores)}\n")
        mensaje = (
            f"✔️ {procesados} documentos procesados.\n"
            f"❌ {len(errores)} documentos con errores.\n"
            f"🕒 Tiempo: {duration:.2f} segundos.\n\n"
            f"📄 Revisa el archivo 'errores.log' en la carpeta de salida."
        )
    else:
        mensaje = (
            f"✔️ {procesados} documentos procesados.\n"
            f"🕒 Tiempo: {duration:.2f} segundos.\n\n"
            "🎉 Todos fueron procesados correctamente."
        )

    messagebox.showinfo("Proceso completado", mensaje)



def seleccionar_imagen(info):
    ruta = filedialog.askopenfilename(
        title="Seleccionar imagen",
        filetypes=[("Imágenes", "*.png;*.jpg;*.jpeg;*.bmp")]
    )
    if ruta:
        info["path_var"].set(ruta)


def seleccionar_documentos(archivos_var):
    rutas = filedialog.askopenfilenames(
        title="Seleccionar documentos",
        filetypes=[("Word", "*.docx")]
    )
    if rutas:
        archivos_var.set(";".join(rutas))


def seleccionar_carpeta(carpeta_var):
    carpeta = filedialog.askdirectory(title="Seleccionar carpeta de salida")
    if carpeta:
        carpeta_var.set(carpeta)


def main():
    # Cargar config.json SOLO para las claves
    with open('config.json', 'r', encoding='utf-8') as f:
        config = json.load(f)

    replacements_config = config.get("replacements", {})
    image_replacements_config = config.get("image_replacements", {})

    # Crear ventana
    ventana = tk.Tk()
    ventana.title("Generador de Documentos")
    ventana.geometry("600x600")

    tk.Label(ventana, text="Ingrese los valores de texto:", font=("Arial", 12, "bold")).pack(pady=10)

    # Crear entradas dinámicamente para texto
    valores_usuario = {}
    for key in replacements_config.keys():
        frame = tk.Frame(ventana)
        frame.pack(pady=5, fill="x", padx=20)
        tk.Label(frame, text=key, width=20, anchor="w").pack(side="left")
        entrada = tk.Entry(frame, width=30)
        entrada.pack(side="left", expand=True, fill="x")
        valores_usuario[key] = entrada

    # Crear botones para selección de imágenes
    tk.Label(ventana, text="Seleccione las imágenes:", font=("Arial", 12, "bold")).pack(pady=10)
    imagenes_usuario = {}
    for key in image_replacements_config.keys():
        frame = tk.Frame(ventana)
        frame.pack(pady=5, fill="x", padx=20)

        tk.Label(frame, text=key, width=20, anchor="w").pack(side="left")

        path_var = tk.StringVar(value="")  # SIEMPRE en blanco
        entry = tk.Entry(frame, textvariable=path_var, width=30)
        entry.pack(side="left", expand=True, fill="x")

        boton = tk.Button(frame, text="Seleccionar", command=lambda i={"path_var": path_var}: seleccionar_imagen(i))
        boton.pack(side="left", padx=5)

        imagenes_usuario[key] = {"path_var": path_var}

    # Variables para carpetas
    carpeta_entrada_var = tk.StringVar(value="")
    carpeta_salida_var = tk.StringVar(value="")

    # Selección de carpeta de entrada
    frame_in = tk.Frame(ventana)
    frame_in.pack(pady=10, fill="x", padx=20)
    tk.Label(frame_in, text="Carpeta entrada:", width=20, anchor="w").pack(side="left")
    tk.Entry(frame_in, textvariable=carpeta_entrada_var, width=30).pack(side="left", expand=True, fill="x")
    tk.Button(frame_in, text="Seleccionar", command=lambda: seleccionar_carpeta(carpeta_entrada_var)).pack(side="left", padx=5)

    # Selección de carpeta de salida
    frame_out = tk.Frame(ventana)
    frame_out.pack(pady=10, fill="x", padx=20)
    tk.Label(frame_out, text="Carpeta salida:", width=20, anchor="w").pack(side="left")
    tk.Entry(frame_out, textvariable=carpeta_salida_var, width=30).pack(side="left", expand=True, fill="x")
    tk.Button(frame_out, text="Seleccionar", command=lambda: seleccionar_carpeta(carpeta_salida_var)).pack(side="left", padx=5)

    # Botón de procesar
    boton = tk.Button(
        ventana,
        text="Procesar Documentos",
        command=lambda: ejecutar_proceso(valores_usuario, imagenes_usuario, carpeta_entrada_var, carpeta_salida_var),
        bg="blue", fg="white"
    )
    boton.pack(pady=20)

    ventana.mainloop()


if __name__ == "__main__":
    main()
