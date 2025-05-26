import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkcalendar import DateEntry
from openpyxl import load_workbook
from docx import Document
from docx.shared import Inches
import shutil, os

# Cargar datos desde Excel
excel_path = r"C:\Users\javie\Desktop\Generador de reporte\NODOS.xlsx"
wb = load_workbook(excel_path)
ws = wb.active

clientes_data = []
for row in ws.iter_rows(min_row=2, values_only=True):
    clientes_data.append({
        "nombre_nodo": row[3],
        "municipio": row[13],
        "direccion": row[17],
        "codigo_postal": row[18],
        "nombre_sitio": row[2],
        "tipo_espacio": row[4],
        "entidad": row[6],
        "latitud": row[15],
        "longitud": row[14],
        "id_nodo_prbd": row[1] if len(row) > 9 else ""
    })

nombres_nodo = [c["nombre_nodo"] for c in clientes_data]
imagenes = []

# Función para subir fotos
def subir_fotos():
    archivos = filedialog.askopenfilenames(filetypes=[("Imágenes", "*.jpg *.jpeg *.png")])
    if archivos:
        imagenes.clear()
        imagenes.extend(archivos)
        lbl_fotos.config(text=f"{len(imagenes)} foto(s) seleccionada(s)")

def reemplazar_texto(doc, buscar, reemplazo):
    for p in doc.paragraphs:
        if buscar in p.text:
            for run in p.runs:
                if buscar in run.text:
                    run.text = run.text.replace(buscar, reemplazo)

    for tabla in doc.tables:
        for fila in tabla.rows:
            for celda in fila.cells:
                for p in celda.paragraphs:
                    if buscar in p.text:
                        for run in p.runs:
                            if buscar in run.text:
                                run.text = run.text.replace(buscar, reemplazo)

def generar_reporte():
    seleccionado = combo.get()
    datos = next((c for c in clientes_data if c["nombre_nodo"] == seleccionado), None)

    if not datos:
        messagebox.showerror("Error", "Nodo no encontrado.")
        return

    try:
        shutil.copy(r"C:\Users\javie\Desktop\Generador de reporte\REPORTE CINERIA.docx", "temp_reporte.docx")
    except FileNotFoundError:
        messagebox.showerror("Error", "No se encontró 'REPORTE CINERIA.docx'")
        return

    doc = Document("temp_reporte.docx")

    # Reemplazo de textos
    reemplazar_texto(doc, "cliente", "RED GUANAJUATO")
    reemplazar_texto(doc, "municipio", f"{datos['municipio']}")
    reemplazar_texto(doc, "direccion", f"{datos['direccion']}")
    reemplazar_texto(doc, "codigo", f"{datos['codigo_postal']}")
    reemplazar_texto(doc, "nombre del sitio", f"{datos['nombre_sitio']}")
    reemplazar_texto(doc, "tipo de espacio", f"{datos['tipo_espacio']}")
    reemplazar_texto(doc, "entidad", f"{datos['entidad']}")
    reemplazar_texto(doc, "latitud", f"{str(datos['latitud'])}")
    reemplazar_texto(doc, "longitud", f"{str(datos['longitud'])}")
    reemplazar_texto(doc, "id", f"{datos['id_nodo_prbd']}")
    reemplazar_texto(doc, "emision", f"{fecha_emision.get()}")
    reemplazar_texto(doc, "fecha de apertura", f"{fecha_apertura.get()}")
    reemplazar_texto(doc, "llegada", f"{fecha_llegada.get()}")
    reemplazar_texto(doc, "cierre", f"{fecha_cierre.get()}")
    reemplazar_texto(doc, "Trabajador", trabajador_var.get())
    reemplazar_texto(doc, "Hora", hora_var.get())
    reemplazar_texto(doc, "Tecnico", tecnico_var.get())
    reemplazar_texto(doc, "Servicio", servicio_var.get())
    reemplazar_texto(doc, "actividades", actividades_var.get())
    reemplazar_texto(doc, "Mantenimiento", mantenimiento_var.get())

    if imagenes:
        doc.add_page_break()
        doc.add_heading("EVIDENCIA FOTOGRÁFICA", level=1)
        for img in imagenes:
            doc.add_picture(img, width=Inches(4))
            doc.add_paragraph("")

    guardar_como = filedialog.asksaveasfilename(defaultextension=".docx", filetypes=[("Word Document", "*.docx")])
    if guardar_como:
        doc.save(guardar_como)
        os.remove("temp_reporte.docx")
        messagebox.showinfo("Éxito", f"Reporte guardado:\n{guardar_como}")

# Interfaz gráfica
root = tk.Tk()
root.title("Generador de Reportes CINERGIA")

# Nodo
tk.Label(root, text="Selecciona un nodo:", font=("Arial", 12)).pack(pady=5)
combo = ttk.Combobox(root, values=nombres_nodo, width=60)
combo.pack(pady=5)

# Fechas
tk.Label(root, text="Fecha de emisión del reporte:").pack()
fecha_emision = DateEntry(root, width=20)
fecha_emision.pack()

tk.Label(root, text="Fecha de apertura:").pack()
fecha_apertura = DateEntry(root, width=20)
fecha_apertura.pack()

tk.Label(root, text="Fecha de llegada al sitio:").pack()
fecha_llegada = DateEntry(root, width=20)
fecha_llegada.pack()

tk.Label(root, text="Fecha de cierre:").pack()
fecha_cierre = DateEntry(root, width=20)
fecha_cierre.pack()

# Trabajador
tk.Label(root, text="Selecciona al trabajador:").pack()
trabajador_var = tk.StringVar()
trabajadores = ["Carlos Pérez", "María López", "Luis Sánchez", "Otro"]
combo_trabajador = ttk.Combobox(root, textvariable=trabajador_var, values=trabajadores, width=40)
combo_trabajador.pack()

# Hora
tk.Label(root, text="Hora (ej. 03:30 PM):").pack()
hora_var = tk.StringVar()
hora_entry = tk.Entry(root, textvariable=hora_var, width=20)
hora_entry.pack()

# Técnico
tk.Label(root, text="Técnico:").pack()
tecnico_var = tk.StringVar()
tecnicos = ["Daniel Vázquez", "Leticia Hernández", "Javier Rojas", "Otro"]
combo_tecnico = ttk.Combobox(root, textvariable=tecnico_var, values=tecnicos, width=40)
combo_tecnico.pack()

# Servicio y 
tk.Label(root, text="Servicio :").pack()
servicio_var = tk.StringVar()
servicios = ["sevicios", "sevicios de equipo", "sevicios Revisión de red", "Otro"]
combo_servicio = ttk.Combobox(root, textvariable=servicio_var, values=servicios, width=50)
combo_servicio.pack()

# Mantenimiento
tk.Label(root, text="mantenimiento:").pack()
mantenimiento_var = tk.StringVar()
mantenimiento = ["Mantenimiento preventivo", "Instalación de equipo", "Revisión de red", "Otro"]
combo_mantenimiento = ttk.Combobox(root, textvariable=mantenimiento_var, values=mantenimiento, width=50)
combo_mantenimiento.pack()

# Actividades realizadas
tk.Label(root, text="Actividades realizadas:").pack()
actividades_var = tk.StringVar()
tk.Entry(root, textvariable=actividades_var, width=60).pack()

# Subir fotos
tk.Button(root, text="Subir fotos", command=subir_fotos).pack(pady=5)
lbl_fotos = tk.Label(root, text="Ninguna foto seleccionada")
lbl_fotos.pack()

# Botón generar
tk.Button(root, text="Generar reporte", command=generar_reporte, bg="green", fg="white", font=("Arial", 11)).pack(pady=15)

root.mainloop()