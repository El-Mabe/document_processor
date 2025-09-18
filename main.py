import os
import json
import time
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from processor.word_editor import process_word_file
from processor.excel_editor import process_excel_file 


def ejecutar_proceso(valores_usuario, imagenes_usuario, carpeta_entrada_var, carpeta_salida_var, progress_var, progress_bar, status_label):
    """Ejecuta el proceso de reemplazo con barra de progreso y mejor manejo de errores."""
    start_time = time.time()

    # Obtener valores de reemplazo
    replacements = {key: entrada.get() for key, entrada in valores_usuario.items()}
    
    # Filtrar reemplazos vacíos (opcional)
    replacements = {k: v for k, v in replacements.items() if v.strip()}

    # Obtener imágenes con información completa
    image_replacements = {}
    for key, info in imagenes_usuario.items():
        ruta = info["path_var"].get()
        if ruta and os.path.exists(ruta):  # 👈 Verificar que existe
            # Usar formato completo de imagen si hay información adicional
            if "width_cm" in info or "height_cm" in info:
                image_replacements[key] = {
                    "path": ruta,
                    "width_cm": info.get("width_cm", 3.5),
                    "height_cm": info.get("height_cm", 1.5)
                }
            else:
                image_replacements[key] = ruta

    carpeta_entrada = carpeta_entrada_var.get()
    carpeta_salida = carpeta_salida_var.get()

    if not carpeta_entrada:
        messagebox.showwarning("Advertencia", "Debe seleccionar una carpeta de entrada.")
        return
    if not carpeta_salida:
        messagebox.showwarning("Advertencia", "Debe seleccionar una carpeta de salida.")
        return

    # Contar archivos total para barra de progreso
    total_files = 0
    for dirpath, _, filenames in os.walk(carpeta_entrada):
        for filename in filenames:
            if filename.lower().endswith((".docx", ".xlsx")):
                total_files += 1

    if total_files == 0:
        messagebox.showinfo("Información", "No se encontraron archivos .docx o .xlsx en la carpeta seleccionada.")
        return

    # Configurar barra de progreso
    progress_var.set(0)
    progress_bar["maximum"] = total_files

    procesados = 0
    errores = []
    archivos_procesados = []

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

                # Actualizar status
                status_label.config(text=f"Procesando: {relative_input}")
                status_label.update()

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
                    archivos_procesados.append(relative_input)

                except Exception as e:
                    tipo = "Word" if filename.lower().endswith(".docx") else "Excel"
                    error_msg = f"[{tipo}] {relative_input} -> {str(e)}"
                    print(f"⚠️ Error en {error_msg}")
                    errores.append(error_msg)

                # Actualizar barra de progreso
                progress_var.set(procesados + len(errores))
                progress_bar.update()

    end_time = time.time()
    duration = end_time - start_time

    # Limpiar status
    status_label.config(text="Proceso completado")

    # Crear log detallado
    log_path = os.path.join(carpeta_salida, "proceso_log.txt")
    with open(log_path, "w", encoding="utf-8") as f:
        f.write("=== REPORTE DE PROCESAMIENTO ===\n\n")
        f.write(f"Fecha: {time.strftime('%Y-%m-%d %H:%M:%S')}\n")
        f.write(f"Tiempo total: {duration:.2f} segundos\n")
        f.write(f"Archivos procesados: {procesados}\n")
        f.write(f"Archivos con errores: {len(errores)}\n\n")
        
        if replacements:
            f.write("=== REEMPLAZOS DE TEXTO APLICADOS ===\n")
            for key, value in replacements.items():
                f.write(f"{key} -> {value}\n")
            f.write("\n")
        
        if image_replacements:
            f.write("=== REEMPLAZOS DE IMÁGENES APLICADOS ===\n")
            for key, value in image_replacements.items():
                if isinstance(value, dict):
                    f.write(f"{key} -> {value['path']}\n")
                else:
                    f.write(f"{key} -> {value}\n")
            f.write("\n")

        if archivos_procesados:
            f.write("=== ARCHIVOS PROCESADOS CORRECTAMENTE ===\n")
            for archivo in archivos_procesados:
                f.write(f"✓ {archivo}\n")
            f.write("\n")

        if errores:
            f.write("=== ARCHIVOS CON ERRORES ===\n")
            for err in errores:
                f.write(f"✗ {err}\n")

    # Mensaje final mejorado
    if errores:
        mensaje = (
            f"✔️ {procesados} documentos procesados correctamente.\n"
            f"❌ {len(errores)} documentos con errores.\n"
            f"🕒 Tiempo: {duration:.2f} segundos.\n\n"
            f"📄 Revisa el archivo 'proceso_log.txt' en la carpeta de salida para detalles."
        )
        messagebox.showwarning("Proceso completado con errores", mensaje)
    else:
        mensaje = (
            f"🎉 ¡Proceso completado exitosamente!\n\n"
            f"✔️ {procesados} documentos procesados.\n"
            f"🕒 Tiempo: {duration:.2f} segundos.\n\n"
            f"📄 Log detallado guardado en 'proceso_log.txt'"
        )
        messagebox.showinfo("Proceso completado", mensaje)


def seleccionar_imagen(info):
    """Selecciona una imagen con validación mejorada."""
    ruta = filedialog.askopenfilename(
        title="Seleccionar imagen",
        filetypes=[
            ("Todas las imágenes", "*.png *.jpg *.jpeg *.bmp *.gif *.tiff"),
            ("PNG", "*.png"),
            ("JPEG", "*.jpg *.jpeg"),
            ("BMP", "*.bmp"),
            ("Todos los archivos", "*.*")
        ]
    )
    if ruta:
        if os.path.exists(ruta):
            info["path_var"].set(ruta)
        else:
            messagebox.showerror("Error", "El archivo seleccionado no existe.")


def seleccionar_carpeta(carpeta_var, titulo="Seleccionar carpeta"):
    """Selecciona una carpeta con título personalizable."""
    carpeta = filedialog.askdirectory(title=titulo)
    if carpeta:
        carpeta_var.set(carpeta)


def validar_configuracion():
    """Valida que existe el archivo de configuración."""
    if not os.path.exists('config.json'):
        messagebox.showerror(
            "Error de configuración", 
            "No se encontró el archivo 'config.json'.\n"
            "Por favor, asegúrate de que existe en el mismo directorio que este programa."
        )
        return False
    return True


def main():
    """Función principal con interfaz mejorada."""
    # Validar configuración
    if not validar_configuracion():
        return

    # Cargar config.json
    try:
        with open('config.json', 'r', encoding='utf-8') as f:
            config = json.load(f)
    except Exception as e:
        messagebox.showerror("Error", f"Error al cargar config.json: {str(e)}")
        return

    replacements_config = config.get("replacements", {})
    image_replacements_config = config.get("image_replacements", {})

    # Crear ventana principal
    ventana = tk.Tk()
    ventana.title("Generador de Documentos v2.0")
    ventana.geometry("700x700")
    ventana.configure(bg="#f0f0f0")

    # Crear un canvas con scrollbar para contenido largo
    canvas = tk.Canvas(ventana, bg="#f0f0f0")
    scrollbar = ttk.Scrollbar(ventana, orient="vertical", command=canvas.yview)
    scrollable_frame = ttk.Frame(canvas)

    scrollable_frame.bind(
        "<Configure>",
        lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
    )

    canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
    canvas.configure(yscrollcommand=scrollbar.set)

    # Título principal
    titulo = tk.Label(scrollable_frame, text="🔄 Generador de Documentos", 
                      font=("Arial", 16, "bold"), bg="#f0f0f0", fg="#2c3e50")
    titulo.pack(pady=15)

    # Sección de texto
    if replacements_config:
        frame_texto = ttk.LabelFrame(scrollable_frame, text="📝 Valores de Texto", padding=10)
        frame_texto.pack(pady=10, padx=20, fill="x")

        valores_usuario = {}
        for key in replacements_config.keys():
            frame = ttk.Frame(frame_texto)
            frame.pack(pady=3, fill="x")
            
            ttk.Label(frame, text=key, width=25).pack(side="left")
            entrada = ttk.Entry(frame, width=40)
            entrada.pack(side="left", expand=True, fill="x", padx=(10, 0))
            valores_usuario[key] = entrada
    else:
        valores_usuario = {}

    # Sección de imágenes
    if image_replacements_config:
        frame_imagenes = ttk.LabelFrame(scrollable_frame, text="🖼️ Selección de Imágenes", padding=10)
        frame_imagenes.pack(pady=10, padx=20, fill="x")

        imagenes_usuario = {}
        for key in image_replacements_config.keys():
            frame = ttk.Frame(frame_imagenes)
            frame.pack(pady=3, fill="x")

            ttk.Label(frame, text=key, width=25).pack(side="left")

            path_var = tk.StringVar(value="")
            entry = ttk.Entry(frame, textvariable=path_var, width=35, state="readonly")
            entry.pack(side="left", expand=True, fill="x", padx=(10, 5))

            boton = ttk.Button(frame, text="Buscar", width=10,
                             command=lambda i={"path_var": path_var}: seleccionar_imagen(i))
            boton.pack(side="left")

            imagenes_usuario[key] = {"path_var": path_var}
    else:
        imagenes_usuario = {}

    # Sección de carpetas
    frame_carpetas = ttk.LabelFrame(scrollable_frame, text="📁 Carpetas de Trabajo", padding=10)
    frame_carpetas.pack(pady=10, padx=20, fill="x")

    carpeta_entrada_var = tk.StringVar(value="")
    carpeta_salida_var = tk.StringVar(value="")

    # Carpeta de entrada
    frame_in = ttk.Frame(frame_carpetas)
    frame_in.pack(pady=3, fill="x")
    ttk.Label(frame_in, text="Carpeta entrada:", width=25).pack(side="left")
    ttk.Entry(frame_in, textvariable=carpeta_entrada_var, width=35, state="readonly").pack(side="left", expand=True, fill="x", padx=(10, 5))
    ttk.Button(frame_in, text="Buscar", width=10, 
               command=lambda: seleccionar_carpeta(carpeta_entrada_var, "Seleccionar carpeta de entrada")).pack(side="left")

    # Carpeta de salida
    frame_out = ttk.Frame(frame_carpetas)
    frame_out.pack(pady=3, fill="x")
    ttk.Label(frame_out, text="Carpeta salida:", width=25).pack(side="left")
    ttk.Entry(frame_out, textvariable=carpeta_salida_var, width=35, state="readonly").pack(side="left", expand=True, fill="x", padx=(10, 5))
    ttk.Button(frame_out, text="Buscar", width=10,
               command=lambda: seleccionar_carpeta(carpeta_salida_var, "Seleccionar carpeta de salida")).pack(side="left")

    # Sección de progreso
    frame_progreso = ttk.LabelFrame(scrollable_frame, text="📊 Progreso", padding=10)
    frame_progreso.pack(pady=10, padx=20, fill="x")

    progress_var = tk.DoubleVar()
    progress_bar = ttk.Progressbar(frame_progreso, variable=progress_var, length=400)
    progress_bar.pack(pady=5, fill="x")

    status_label = ttk.Label(frame_progreso, text="Listo para procesar", foreground="green")
    status_label.pack(pady=5)

    # Botón de procesar
    boton_procesar = tk.Button(
        scrollable_frame,
        text="🚀 Procesar Documentos",
        command=lambda: ejecutar_proceso(
            valores_usuario, imagenes_usuario, 
            carpeta_entrada_var, carpeta_salida_var,
            progress_var, progress_bar, status_label
        ),
        bg="#3498db", fg="white", font=("Arial", 12, "bold"),
        padx=20, pady=10
    )
    boton_procesar.pack(pady=20)

    # Configurar scroll
    canvas.pack(side="left", fill="both", expand=True)
    scrollbar.pack(side="right", fill="y")

    # Configurar scroll con rueda del mouse
    def _on_mousewheel(event):
        canvas.yview_scroll(int(-1*(event.delta/120)), "units")
    
    canvas.bind_all("<MouseWheel>", _on_mousewheel)

    # Centrar ventana
    ventana.update_idletasks()
    x = (ventana.winfo_screenwidth() // 2) - (700 // 2)
    y = (ventana.winfo_screenheight() // 2) - (700 // 2)
    ventana.geometry(f"700x700+{x}+{y}")

    ventana.mainloop()


if __name__ == "__main__":
    main()