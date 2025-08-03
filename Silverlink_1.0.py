import pandas as pd
from fuzzywuzzy import process
import customtkinter as ctk
from tkinter import ttk, messagebox, scrolledtext

# Cargar el archivo Excel principal
archivo_excel_principal = 'buscadorb.xlsx'  # Reemplaza con el nombre de tu archivo Excel
df_principal = pd.read_excel(archivo_excel_principal)

# Cargar el archivo Excel de desglose
archivo_excel_desglose = 'desglose_productos.xlsx'  # Reemplaza con el nombre de tu archivo Excel
df_desglose = pd.read_excel(archivo_excel_desglose)

# Lista de columnas adicionales que quieres mostrar
columnas_adicionales = ["1661", "1662", "1670", "1672", "1675", "1678", "1681", "1686",
                        "1691", "1696", "1708", "1726", "1730", "1739", "Entrearmadas"]  # Añadir "Entrearmadas"

# Variables globales
detalles_globales = None
opciones_seleccion = []

# Función para formatear la columna "Valor"
def formatear_valor(valor):
    if pd.isna(valor):  # Manejar celdas vacías
        return "0"
    return f"{int(valor):,}".replace(",", ".")  # Formatear sin decimales y con punto para miles

# Función para buscar múltiples coincidencias en la columna "Nombre"
def buscar_coincidencias(nombre, umbral=80, limite_resultados=5):
    nombres = df_principal['Nombre'].tolist()
    resultados = process.extract(nombre, nombres, limit=limite_resultados)
    coincidencias = []
    for resultado, puntaje in resultados:
        if puntaje >= umbral:
            fila = df_principal[df_principal['Nombre'] == resultado].iloc[0]
            datos = {
                "Nombre": resultado,
                "Total": fila["Total"],
                "Puntaje de similitud": puntaje
            }
            for columna in columnas_adicionales:
                datos[columna] = fila[columna] if not pd.isna(fila[columna]) else "N/A"  # Manejar celdas vacías
            coincidencias.append(datos)
    if coincidencias:
        return coincidencias
    else:
        return "No se encontraron coincidencias lo suficientemente cercanas."

# Función para obtener detalles del desglose de productos
def obtener_detalles_desglose(nombre):
    df_filtrado = df_desglose[df_desglose['Nombre'] == nombre]
    if df_filtrado.empty:
        return "No se encontraron detalles para este nombre."
    
    # Agrupar por Producto, sumar la columna "Valor", y concatenar Años y Fuentes
    detalles = df_filtrado.groupby('Producto', as_index=False).agg({
        'Valor': 'sum',
        'Año': lambda x: ', '.join(map(str, x.unique())),
        'Fuente': lambda x: ', '.join(map(str, x.dropna().unique()))  # Ignorar valores vacíos en "Fuente"
    })
    
    # Formatear la columna "Valor"
    detalles['Valor'] = detalles['Valor'].apply(formatear_valor)
    
    return detalles

# Función que se ejecuta al presionar el botón de búsqueda
def buscar():
    global opciones_seleccion
    nombre_a_buscar = entrada_nombre.get()
    if nombre_a_buscar.strip() == "":
        messagebox.showwarning("Advertencia", "Por favor, ingresa un nombre para buscar.")
        return

    coincidencias = buscar_coincidencias(nombre_a_buscar)
    area_texto.delete(1.0, "end")

    if isinstance(coincidencias, list):
        opciones_seleccion = [coincidencia["Nombre"] for coincidencia in coincidencias]
        combo_opciones.configure(values=opciones_seleccion)  # Actualizar el Combobox con las opciones
        combo_opciones.set(opciones_seleccion[0])  # Seleccionar la primera opción por defecto
        area_texto.insert("end", f"Coincidencias para '{nombre_a_buscar}':\n\n")
        for i, datos in enumerate(coincidencias, start=1):
            area_texto.insert("end", f"Coincidencia {i}:\n")
            for clave, valor in datos.items():
                area_texto.insert("end", f"{clave}: {valor}\n")
            area_texto.insert("end", "\n")
    else:
        area_texto.insert("end", coincidencias)

# Función que se ejecuta al seleccionar una opción del Combobox
def seleccionar_opcion(event):
    nombre_seleccionado = combo_opciones.get()
    if nombre_seleccionado:
        # Mostrar los valores asociados en el área de texto
        area_texto.delete(1.0, "end")
        coincidencias = buscar_coincidencias(nombre_seleccionado)
        if isinstance(coincidencias, list):
            for coincidencia in coincidencias:
                if coincidencia["Nombre"] == nombre_seleccionado:
                    area_texto.insert("end", f"Detalles para '{nombre_seleccionado}':\n\n")
                    for clave, valor in coincidencia.items():
                        area_texto.insert("end", f"{clave}: {valor}\n")
                    break

# Función que se ejecuta al presionar el botón de detalles
def mostrar_detalles():
    global detalles_globales
    nombre_seleccionado = combo_opciones.get()
    if not nombre_seleccionado:
        messagebox.showwarning("Advertencia", "Por favor, selecciona una opción primero.")
        return

    detalles_globales = obtener_detalles_desglose(nombre_seleccionado)

    if isinstance(detalles_globales, pd.DataFrame):
        # Limpiar la tabla antes de cargar nuevos datos
        for row in tabla_detalles.get_children():
            tabla_detalles.delete(row)
        
        # Configurar las columnas del Treeview
        columnas = detalles_globales.columns.tolist()
        tabla_detalles["columns"] = columnas
        tabla_detalles["show"] = "headings"
        for col in columnas:
            tabla_detalles.heading(col, text=col)
            tabla_detalles.column(col, width=100, stretch=True)  # Permitir expansión horizontal
        
        # Insertar los datos en la tabla
        for _, fila in detalles_globales.iterrows():
            valores = [fila[col] if not pd.isna(fila[col]) else "N/A" for col in columnas]  # Manejar celdas vacías
            tabla_detalles.insert("", "end", values=valores)
        
        # Mostrar el frame de la tabla
        frame_tabla.pack(pady=10, fill="both", expand=True)
    else:
        messagebox.showwarning("Advertencia", detalles_globales)

# Función para cambiar el modo de apariencia
def cambiar_modo(modo):
    ctk.set_appearance_mode(modo)

# Crear la ventana principal con customtkinter
ctk.set_appearance_mode("dark")  # Modo oscuro por defecto
ctk.set_default_color_theme("blue")  # Tema azul por defecto

ventana = ctk.CTk()
ventana.title("Silverlinks 1.0")  # Nombre del programa
ventana.geometry("1000x700")

# Configurar la fuente Arial para todo el programa
fuente = ("Arial", 12)  # Tamaño de fuente predeterminado

# Crear un frame para los botones de modo claro/oscuro en la esquina inferior derecha
frame_botones_modo = ctk.CTkFrame(ventana, corner_radius=10)
frame_botones_modo.place(relx=1.0, rely=1.0, anchor="se", x=-10, y=-10)

# Botones para cambiar el modo de apariencia
boton_modo_claro = ctk.CTkButton(frame_botones_modo, text="☀️", width=30, height=30, command=lambda: cambiar_modo("light"))
boton_modo_claro.pack(side="left", padx=5, pady=5)

boton_modo_oscuro = ctk.CTkButton(frame_botones_modo, text="🌙", width=30, height=30, command=lambda: cambiar_modo("dark"))
boton_modo_oscuro.pack(side="left", padx=5, pady=5)

# Crear una etiqueta con fuente Arial
etiqueta = ctk.CTkLabel(ventana, text="Ingresa el nombre a buscar:", font=("Arial", 14))
etiqueta.pack(pady=10)

# Crear un campo de entrada para la búsqueda con fuente Arial
entrada_nombre = ctk.CTkEntry(ventana, width=400, font=fuente)
entrada_nombre.pack(pady=10)

# Crear un botón de búsqueda con fuente Arial
boton_buscar = ctk.CTkButton(ventana, text="Buscar", font=fuente, command=buscar)
boton_buscar.pack(pady=10)

# Crear un Combobox para seleccionar una opción con fuente Arial
combo_opciones = ttk.Combobox(ventana, width=50, state="readonly", font=fuente)
combo_opciones.pack(pady=10)
combo_opciones.bind("<<ComboboxSelected>>", seleccionar_opcion)  # Vincular evento de selección

# Crear un área de texto con barra de desplazamiento para mostrar coincidencias
area_texto = scrolledtext.ScrolledText(ventana, width=120, height=10, wrap="word", font=("Arial", 12))
area_texto.pack(pady=10)

# Crear un botón para mostrar detalles con fuente Arial
boton_detalles = ctk.CTkButton(ventana, text="Mostrar Detalles", font=fuente, command=mostrar_detalles)
boton_detalles.pack(pady=10)

# Crear un frame para la tabla de detalles
frame_tabla = ctk.CTkFrame(ventana)
frame_tabla.pack(pady=10, fill="both", expand=True)
frame_tabla.pack_forget()  # Ocultar el frame inicialmente

# Crear un Treeview para mostrar los detalles en forma de tabla
tabla_detalles = ttk.Treeview(frame_tabla)
tabla_detalles.pack(side="left", fill="both", expand=True)

# Configurar la fuente Arial para el Treeview
style = ttk.Style()
style.configure("Treeview", font=("Arial", 12))
style.configure("Treeview.Heading", font=("Arial", 12, "bold"))

# Añadir barras de desplazamiento vertical y horizontal
scrollbar_vertical = ttk.Scrollbar(frame_tabla, orient="vertical", command=tabla_detalles.yview)
scrollbar_vertical.pack(side="right", fill="y")
scrollbar_horizontal = ttk.Scrollbar(frame_tabla, orient="horizontal", command=tabla_detalles.xview)
scrollbar_horizontal.pack(side="bottom", fill="x")
tabla_detalles.configure(yscrollcommand=scrollbar_vertical.set, xscrollcommand=scrollbar_horizontal.set)

# Leer el contenido del archivo hoja.txt
try:
    with open("DatosPrograma.txt", "r", encoding="utf-8") as archivo:
        disclaimer_texto = archivo.read()
except FileNotFoundError:
    disclaimer_texto = "Disclaimer no disponible."

# Crear un label para mostrar el disclaimer en la parte inferior de la ventana
disclaimer_label = ctk.CTkLabel(ventana, text=disclaimer_texto, font=("Arial", 10), wraplength=980, justify="left")
disclaimer_label.pack(side="bottom", pady=10, padx=10, fill="x")

# Ejecutar la ventana
ventana.mainloop()