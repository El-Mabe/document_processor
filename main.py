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
    
    # 🔥 MEJORA: Filtrar reemplazos vacíos pero conservar los que el usuario quiere vacíos explícitamente
    replacements = {k: v for k, v in replacements.items() if v is not None}
    
    # 🔥 NUEVA FUNCIONALIDAD: Detectar si hay reemplazos que requieren preservación especial de estilos
    special_style_keys = []
    for key in replacements.keys():
        # Detectar patrones que usualmente tienen formato especial
        if any(word in key.upper() for word in ['NOMBRE', 'TITULO', 'CARGO', 'EMPRESA', 'DIRECTOR']):
            special_style_keys.append(key)
    
    if special_style_keys:
        print(f"🎨 Detectados {len(special_style_keys)} campos con posible formato especial: {special_style_keys}")

    # 🔥 NUEVO: Separar placeholders de reemplazos de texto
    image_replacements = {}  # Para reemplazos por texto {{LOGO}}
    placeholder_replacements = {}  # Para placeholders de imagen
    
    for key, info in imagenes_usuario.items():
        ruta = info["path_var"].get()
        if ruta and os.path.exists(ruta):
            # Determinar si es placeholder o reemplazo por texto
            if key.startswith("placeholder_") or key.endswith(".png"):
                # Es un placeholder de imagen
                placeholder_config = {
                    "path": ruta,
                    "maintain_aspect": info.get("maintain_aspect", False)
                }
                # Agregar configuraciones adicionales si están disponibles en el config
                if "width_cm" in info:
                    placeholder_config["width_cm"] = info["width_cm"]
                if "height_cm" in info:
                    placeholder_config["height_cm"] = info["height_cm"]
                
                placeholder_replacements[key] = placeholder_config
                print(f"🖼️ Placeholder configurado: {key} -> {os.path.basename(ruta)}")
            else:
                # Es reemplazo por texto tradicional
                image_config = {
                    "path": ruta,
                    "width_cm": info.get("width_cm", 3.5),
                    "height_cm": info.get("height_cm", 1.5)
                }
                
                # Agregar configuraciones adicionales si están disponibles
                if "resize_mode" in info:
                    image_config["resize_mode"] = info["resize_mode"]
                if "maintain_aspect" in info:
                    image_config["maintain_aspect"] = info["maintain_aspect"]
                if "alignment" in info:
                    image_config["alignment"] = info["alignment"]
                    
                image_replacements[key] = image_config
        elif ruta:  # Si hay ruta pero el archivo no existe
            print(f"⚠️ Advertencia: La imagen {ruta} no existe, se omitirá el reemplazo para {key}")

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
    supported_extensions = (".docx", ".xlsx")
    for dirpath, _, filenames in os.walk(carpeta_entrada):
        for filename in filenames:
            if filename.lower().endswith(supported_extensions):
                total_files += 1

    if total_files == 0:
        messagebox.showinfo("Información", 
                           f"No se encontraron archivos {', '.join(supported_extensions)} en la carpeta seleccionada.")
        return

    # Configurar barra de progreso
    progress_var.set(0)
    progress_bar["maximum"] = total_files

    procesados = 0
    errores = []
    archivos_procesados = []
    archivos_con_estilos_preservados = []  # 🔥 NUEVO: Rastrear preservación de estilos

    for dirpath, _, filenames in os.walk(carpeta_entrada):
        for filename in filenames:
            if filename.lower().endswith(supported_extensions):
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
                        # 🔥 MEJORA: Usar versión con preservación de estilos Y placeholders
                        process_word_file(
                            input_path=input_path,
                            output_path=output_path,
                            replacements=replacements,
                            image_replacements=image_replacements,
                            placeholder_replacements=placeholder_replacements  # 🔥 NUEVO
                        )
                        
                        # 🔥 NUEVO: Marcar si este archivo tenía campos con estilos especiales
                        if special_style_keys:
                            archivos_con_estilos_preservados.append(relative_input)

                    elif filename.lower().endswith(".xlsx"):
                        process_excel_file(
                            input_path=input_path,
                            output_path=output_path,
                            replacements=replacements,
                            image_replacements=image_replacements,
                            placeholder_replacements=placeholder_replacements
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

    # 🔥 MEJORADO: Log más detallado con información de estilos
    log_path = os.path.join(carpeta_salida, "proceso_detallado_log.txt")
    with open(log_path, "w", encoding="utf-8") as f:
        f.write("=== REPORTE DETALLADO DE PROCESAMIENTO ===\n\n")
        f.write(f"Fecha: {time.strftime('%Y-%m-%d %H:%M:%S')}\n")
        f.write(f"Tiempo total: {duration:.2f} segundos\n")
        f.write(f"Archivos procesados: {procesados}\n")
        f.write(f"Archivos con errores: {len(errores)}\n")
        f.write(f"Archivos con estilos preservados: {len(archivos_con_estilos_preservados)}\n\n")
        
        if replacements:
            f.write("=== REEMPLAZOS DE TEXTO APLICADOS ===\n")
            for key, value in replacements.items():
                style_note = " (⭐ ESTILO PRESERVADO)" if key in special_style_keys else ""
                f.write(f"{key} -> {value}{style_note}\n")
            f.write("\n")
        
        if image_replacements:
            f.write("=== REEMPLAZOS DE IMÁGENES POR TEXTO APLICADOS ===\n")
            for key, value in image_replacements.items():
                if isinstance(value, dict):
                    f.write(f"{key} -> {value['path']} ")
                    f.write(f"[{value.get('width_cm', 'auto')}x{value.get('height_cm', 'auto')} cm]\n")
                else:
                    f.write(f"{key} -> {value}\n")
            f.write("\n")
        
        if placeholder_replacements:
            f.write("=== REEMPLAZOS DE PLACEHOLDERS APLICADOS ===\n")
            for key, value in placeholder_replacements.items():
                if isinstance(value, dict):
                    f.write(f"{key} -> {value['path']} (maintain_aspect: {value.get('maintain_aspect', False)})\n")
                else:
                    f.write(f"{key} -> {value}\n")
            f.write("\n")

        if archivos_procesados:
            f.write("=== ARCHIVOS PROCESADOS CORRECTAMENTE ===\n")
            for archivo in archivos_procesados:
                style_marker = " 🎨" if archivo in archivos_con_estilos_preservados else ""
                f.write(f"✓ {archivo}{style_marker}\n")
            f.write("\n")
        
        if archivos_con_estilos_preservados:
            f.write("=== ARCHIVOS CON PRESERVACIÓN DE ESTILOS ===\n")
            f.write("Los siguientes archivos tenían texto con formato especial que fue preservado:\n")
            for archivo in archivos_con_estilos_preservados:
                f.write(f"🎨 {archivo}\n")
            f.write("\n")

        if errores:
            f.write("=== ARCHIVOS CON ERRORES ===\n")
            for err in errores:
                f.write(f"✗ {err}\n")
        
        # 🔥 NUEVO: Estadísticas adicionales
        f.write("\n=== ESTADÍSTICAS ADICIONALES ===\n")
        f.write(f"Promedio de tiempo por archivo: {duration/max(procesados, 1):.2f} segundos\n")
        f.write(f"Tasa de éxito: {(procesados/(procesados + len(errores))*100):.1f}%\n")
        f.write(f"Campos con preservación de estilos: {len(special_style_keys)}\n")

    # 🔥 MEJORADO: Mensaje final más informativo
    if errores:
        mensaje = (
            f"✔️ {procesados} documentos procesados correctamente.\n"
            f"❌ {len(errores)} documentos con errores.\n"
            f"🎨 {len(archivos_con_estilos_preservados)} documentos con estilos preservados.\n"
            f"🕒 Tiempo: {duration:.2f} segundos.\n\n"
            f"📄 Revisa 'proceso_detallado_log.txt' para más detalles."
        )
        messagebox.showwarning("Proceso completado con errores", mensaje)
    else:
        mensaje = (
            f"🎉 ¡Proceso completado exitosamente!\n\n"
            f"✔️ {procesados} documentos procesados.\n"
            f"🎨 {len(archivos_con_estilos_preservados)} con estilos preservados.\n"
            f"🕒 Tiempo: {duration:.2f} segundos\n"
            f"⚡ Promedio: {duration/max(procesados, 1):.1f}s por archivo\n\n"
            f"📄 Log detallado: 'proceso_detallado_log.txt'"
        )
        messagebox.showinfo("Proceso completado", mensaje)


def seleccionar_imagen(info):
    """Selecciona una imagen con validación mejorada y preview opcional."""
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
            
            # 🔥 NUEVO: Mostrar información de la imagen seleccionada
            try:
                from PIL import Image
                with Image.open(ruta) as img:
                    width, height = img.size
                    file_size = os.path.getsize(ruta) / 1024  # KB
                    print(f"📷 Imagen seleccionada: {os.path.basename(ruta)} "
                          f"({width}x{height}px, {file_size:.1f}KB)")
            except Exception:
                print(f"📷 Imagen seleccionada: {os.path.basename(ruta)}")
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


def crear_config_ejemplo():
    """Crea un archivo de configuración de ejemplo si no existe."""
    config_ejemplo = {
        "replacements": {
            "{{NOMBRE}}": "",
            "{{APELLIDO}}": "",
            "{{EMPRESA}}": "",
            "{{CARGO}}": "",
            "{{FECHA}}": ""
        },
        "image_replacements": {
            "{{LOGO}}": {
                "width_cm": 4.0,
                "height_cm": 2.5,
                "resize_mode": "fit_placeholder",
                "maintain_aspect": True
            },
            "{{FIRMA}}": {
                "width_cm": 3.0,
                "height_cm": 1.5,
                "resize_mode": "auto_detect"
            }
        }
    }
    
    try:
        with open('config.json', 'w', encoding='utf-8') as f:
            json.dump(config_ejemplo, f, indent=4, ensure_ascii=False)
        return True
    except Exception as e:
        print(f"Error creando config.json: {e}")
        return False


def main():
    """Función principal con interfaz mejorada."""
    # Validar configuración o crear ejemplo
    if not validar_configuracion():
        respuesta = messagebox.askyesno(
            "Crear configuración", 
            "¿Quieres crear un archivo config.json de ejemplo?\n\n"
            "Esto te permitirá personalizar los campos de texto e imágenes."
        )
        if respuesta:
            if crear_config_ejemplo():
                messagebox.showinfo("Configuración creada", 
                                   "Se ha creado config.json con ejemplos.\n"
                                   "Puedes editarlo según tus necesidades.")
            else:
                messagebox.showerror("Error", "No se pudo crear el archivo de configuración.")
                return
        else:
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
    ventana.title("🎨 Generador de Documentos v2.1 - Con Preservación de Estilos")
    ventana.geometry("750x750")
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
    titulo = tk.Label(scrollable_frame, text="🎨 Generador de Documentos v2.1", 
                      font=("Arial", 16, "bold"), bg="#f0f0f0", fg="#2c3e50")
    titulo.pack(pady=10)
    
    # 🔥 NUEVO: Subtítulo informativo
    subtitulo = tk.Label(scrollable_frame, text="✨ Con preservación automática de estilos (negritas, cursivas, etc.)", 
                        font=("Arial", 9), bg="#f0f0f0", fg="#7f8c8d")
    subtitulo.pack(pady=(0, 15))

    # Sección de texto
    if replacements_config:
        frame_texto = ttk.LabelFrame(scrollable_frame, text="📝 Valores de Texto (Estilos Preservados)", padding=10)
        frame_texto.pack(pady=10, padx=20, fill="x")

        valores_usuario = {}
        for key in replacements_config.keys():
            frame = ttk.Frame(frame_texto)
            frame.pack(pady=3, fill="x")
            
            # 🔥 MEJORA: Indicador visual para campos con formato especial
            label_text = key
            if any(word in key.upper() for word in ['NOMBRE', 'TITULO', 'CARGO', 'EMPRESA']):
                label_text += " 🎨"
            
            ttk.Label(frame, text=label_text, width=25).pack(side="left")
            entrada = ttk.Entry(frame, width=40)
            entrada.pack(side="left", expand=True, fill="x", padx=(10, 0))
            valores_usuario[key] = entrada
    else:
        valores_usuario = {}

    # Sección de imágenes mejorada
    if image_replacements_config:
        frame_imagenes = ttk.LabelFrame(scrollable_frame, text="🖼️ Selección de Imágenes (Redimensionado Inteligente)", padding=10)
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

            # 🔥 NUEVO: Agregar información de configuración de imagen
            image_info = {"path_var": path_var}
            if isinstance(image_replacements_config[key], dict):
                image_info.update(image_replacements_config[key])
            
            imagenes_usuario[key] = image_info
    else:
        imagenes_usuario = {}

    # Resto del código igual...
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
        text="🚀 Procesar Documentos (Con Preservación de Estilos)",
        command=lambda: ejecutar_proceso(
            valores_usuario, imagenes_usuario, 
            carpeta_entrada_var, carpeta_salida_var,
            progress_var, progress_bar, status_label
        ),
        bg="#27ae60", fg="white", font=("Arial", 12, "bold"),
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
    x = (ventana.winfo_screenwidth() // 2) - (750 // 2)
    y = (ventana.winfo_screenheight() // 2) - (750 // 2)
    ventana.geometry(f"750x750+{x}+{y}")

    ventana.mainloop()


if __name__ == "__main__":
    main()