# Generador de resporte 
# Funciona unicamente con "Reportes CINERGIA" y "NODOS"
# Es un programa que hace una copia del Word original y edita los campos seleccionado,
# tambien permite agregar actividades mas de las que ya estan asignadas por defecto,
# la cules deben de estar en un json llamado "actividades.json", el programa permite 
# que el usuario pueda seleccionar fotos guardadas en su equipo y pueda ver cuales son
# mostrando una vista previa y la eliminsacion de las imagenes que no se necesiten
# @Version 1.0 
# 25/06/2025
# By: Javier Yepez Ramirez

import os
import sys
import shutil
import datetime
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkcalendar import DateEntry
from openpyxl import load_workbook
from docx import Document
from docx.shared import Inches
import json
from PIL import Image, ImageTk

root = tk.Tk()

# Ruta al icono
icon_path = os.path.join(os.path.dirname(__file__), 'icono.ico')
print(f"Ruta del icono: {icon_path}")

if os.path.exists(icon_path):
    try:
        root.iconbitmap(icon_path)
    except Exception as e:
        print(f"No se pudo cargar el icono: {e}")
else:
    print("El archivo icono.ico no fue encontrado.")

# Función para recursos empaquetados con PyInstaller
def get_path(relative_path):
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath("."), relative_path)

# Archivos empaquetados
reporte_origen = get_path("REPORTE CINERIA.docx")
actividades_path = get_path("actividades.json")

# Archivos externos (NO empaquetados)
excel_path = os.path.join(os.getcwd(), "NODOS.xlsx")
reporte_destino = os.path.join(os.getcwd(), "temp_reporte.docx")

# Abrir Excel
wb = load_workbook(excel_path)
ws = wb.active


clientes_data = []
for row in ws.iter_rows(min_row=2, values_only=True):
    clientes_data.append({
        "nombre_nodo": row[3],
        "municipio": row[13],
        "direccion": row[17],
        "codigo_postal": row[18],
        "nombre_sitio": row[3],
        "tipo_espacio": row[4],
        "entidad": row[6],
        "latitud": row[14],
        "longitud": row[15],
        "clave_nodo": row[16],
        "nomenclatura_redgto": row[2]
    })

nombres_nodo = [c["nombre_nodo"] for c in clientes_data]
imagenes = []

# --- Funciones --- 
def reemplazar_texto(doc, buscar, reemplazo):
    reemplazo = str(reemplazo)
    
    # Reemplazo en párrafos normales
    for p in doc.paragraphs:
        if buscar in p.text:
            for run in p.runs:
                if run.text and buscar in run.text:
                    run.text = run.text.replace(buscar, reemplazo)
    
    # Reemplazo en tablas
    for tabla in doc.tables:
        for fila in tabla.rows:
            for celda in fila.cells:
                for p in celda.paragraphs:
                    if buscar in p.text:
                        for run in p.runs:
                            if run.text and buscar in run.text:
                                run.text = run.text.replace(buscar, reemplazo)

    # Reemplazo en encabezado (header)
    for section in doc.sections:
        header = section.header
        for p in header.paragraphs:
            if buscar in p.text:
                for run in p.runs:
                    if run.text and buscar in run.text:
                        run.text = run.text.replace(buscar, reemplazo)

def generar_reporte():
    seleccionado = nodo_var.get()
    datos = next((c for c in clientes_data if c["nombre_nodo"] == seleccionado), None)
    if not datos:
        messagebox.showerror("Error", "Nodo no encontrado.")
        return
    try:
        shutil.copy(reporte_origen, reporte_destino)
    except FileNotFoundError:
        messagebox.showerror("Error", "No se encontró 'REPORTE CINERIA.docx'")
        return

    doc = Document(reporte_destino)
    reemplazos = {
        "cliente": "RED GUANAJUATO",
        "municipio": datos['municipio'],
        "direccion": datos['direccion'],
        "codigo": datos['codigo_postal'],
        "nombre": f"{datos['clave_nodo']} - {datos['nombre_sitio']}",
        "tipo de espacio": datos['tipo_espacio'],
        "entidad": entidad_var.get(),
        "latitud": str(datos['latitud']),
        "longitud": str(datos['longitud']),
        "id": datos['nomenclatura_redgto'],
        "emision": fecha_emision.get(),
        "fecha de apertura": fecha_apertura.get(),
        "llegada": fecha_llegada.get(),
        "cierre": fecha_cierre.get(),
        "Trabajador": trabajador_var.get(),
        "Hora": hora_var.get(),
        "Tecnico": tecnico_var.get(),
        "Servicio1": servicio_var.get(),
        "act": '\n'.join(
            listbox_actividades.get(i) for i in orden_seleccion_actividades
        ) or "",
        "Mantenimiento1": mantenimiento_var.get(),
    }

    for buscar, reemplazo in reemplazos.items():
        reemplazar_texto(doc, buscar, reemplazo)

    if imagenes:
        for tabla in doc.tables:
            for fila in tabla.rows:
                for celda in fila.cells:
                    if "Subir fotos" in celda.text:
                        for img_path in imagenes:
                            try:
                                p = celda.add_paragraph()
                                r = p.add_run()
                                r.add_picture(img_path, width=Inches(3))
                            except Exception as e:
                                messagebox.showerror("Error al insertar imagen", f"{img_path}\n{e}")
                        break

    guardar_como = filedialog.asksaveasfilename(defaultextension=".docx", filetypes=[("Word Document", "*.docx")])
    if guardar_como:
        doc.save(guardar_como)
        if os.path.exists(reporte_destino):
            os.remove(reporte_destino)
        messagebox.showinfo("Éxito", f"Reporte guardado:\n{guardar_como}")
        limpiar_campos()

def limpiar_campos():
    hoy = datetime.date.today()
    fecha_emision.set_date(hoy)
    fecha_apertura.set_date(hoy)
    fecha_llegada.set_date(hoy)
    fecha_cierre.set_date(hoy)

    nodo_var.set('')
    entidad_var.set('')
    trabajador_var.set('')
    hora_var.set('')
    tecnico_var.set('')
    servicio_var.set('')
    nueva_actividad_var.set("")
    mantenimiento_var.set('')
    imagenes.clear()
    lbl_fotos.config(text="0 foto(s) seleccionada(s)")
    orden_seleccion_actividades.clear()


def mostrar_creditos():
    ventana_creditos = tk.Toplevel(root)
    ventana_creditos.title("Créditos")
    ventana_creditos.geometry("400x200")
    ventana_creditos.configure(bg="#f2f2f2")
    
    texto_creditos = (
        "Generador de Reportes CINERGIA\n"
        "Desarrollado por: Javier Yepez Ramírez\n"
        "Universidad de Guanajuato - DICIS\n"
        "2025\n"
        "\n"
        "¡Gracias por usar esta aplicación!"
    )
    
    lbl_creditos = ttk.Label(ventana_creditos, text=texto_creditos, justify="center", background="#f2f2f2", font=("Segoe UI", 11))
    lbl_creditos.pack(expand=True, padx=20, pady=20)
    
    btn_cerrar = ttk.Button(ventana_creditos, text="Cerrar", command=ventana_creditos.destroy)
    btn_cerrar.pack(pady=(0, 2))

def agregar_actividad():
    valor_entrada = entry_nueva_actividad.get().strip()
    if valor_entrada and valor_entrada not in listbox_actividades.get(0, "end"):
        listbox_actividades.insert("end", valor_entrada)
        actividades_lista.append(valor_entrada)
        with open(actividades_path, "w", encoding="utf-8") as f:
            json.dump(actividades_lista, f, ensure_ascii=False, indent=2)
        nueva_actividad_var.set("")
    elif not valor_entrada:
        messagebox.showwarning("Campo vacío", "Escribe una actividad para agregarla.")
    else:
        messagebox.showinfo("Duplicado", "La actividad ya está en la lista.")

class AutocompleteCombobox(ttk.Combobox):
    def set_completion_list(self, completion_list):
        self._completion_list = sorted(completion_list, key=str.lower)
        self['values'] = self._completion_list
        self.bind('<KeyRelease>', self.handle_keyrelease)

    def handle_keyrelease(self, event):
        if event.keysym in ("Up", "Down", "Return", "Escape"):
            # Solo abrir desplegable si usuario presiona flecha abajo
            if event.keysym == "Down":
                self.event_generate('<Down>')
            return

        typed = self.get().lower()
        if typed == "":
            filtered = self._completion_list
        else:
            filtered = [item for item in self._completion_list if typed in item.lower()]

        self['values'] = filtered

        # Mantener el texto que el usuario escribió
        current_text = self.get()
        self.delete(0, tk.END)
        self.insert(0, current_text)


def actualizar_orden_seleccion(event):
    seleccion_actual = listbox_actividades.curselection()
    for i in seleccion_actual:
        if i not in orden_seleccion_actividades:
            orden_seleccion_actividades.append(i)

    # Elimina los índices que ya no están seleccionados
    for i in orden_seleccion_actividades[:]:
        if i not in seleccion_actual:
            orden_seleccion_actividades.remove(i)
            
def mostrar_miniaturas():
    # Limpiar el frame
    for widget in frame_miniaturas.winfo_children():
        widget.destroy()

    thumbnails.clear()
    for idx, ruta in enumerate(imagenes):
        try:
            img = Image.open(ruta)
            img.thumbnail((80, 80))
            img_tk = ImageTk.PhotoImage(img)
            thumbnails.append(img_tk)  # guardar referencia para evitar recolección

            marco = tk.Frame(frame_miniaturas, bd=1, relief="raised")
            marco.grid(row=0, column=idx, padx=4)

            lbl_img = tk.Label(marco, image=img_tk)
            lbl_img.pack()

            btn_x = tk.Button(marco, text="X", command=lambda i=idx: eliminar_imagen(i), bg="#e74c3c", fg="white")
            btn_x.pack(fill="x")

        except Exception as e:
            print(f"Error al cargar imagen: {ruta}", e)

def eliminar_imagen(indice):
    if 0 <= indice < len(imagenes):
        del imagenes[indice]
        lbl_fotos.config(text=f"{len(imagenes)} foto(s) seleccionada(s)")
        mostrar_miniaturas()

def subir_fotos():
    archivos = filedialog.askopenfilenames(filetypes=[("Imágenes", "*.jpg *.jpeg *.png")])
    if archivos:
        imagenes.extend(archivos)
        lbl_fotos.config(text=f"{len(imagenes)} foto(s) seleccionada(s)")
        mostrar_miniaturas()
        
#Variables globales
orden_seleccion_actividades = []
imagenes = []  
thumbnails = [] 

# --- INTERFAZ ---
root.title("Generador de Reportes CINERGIA")
root.configure(bg="#ffffff")
root.geometry("950x700")

style = ttk.Style(root)
style.theme_use('clam')
style.configure("TLabel", background="#ffffff", font=("Segoe UI", 10))
style.configure("TCombobox", font=("Segoe UI", 10))
style.configure("TEntry", font=("Segoe UI", 10))
style.configure("TButton", font=("Segoe UI", 11, "bold"), padding=6)
style.map("TButton",
          foreground=[('active', 'white')],
          background=[('active', '#2ecc71')])

# Scrollable main_frame dentro de un canvas
frame_canvas = ttk.Frame(root)
frame_canvas.pack(fill="both", expand=True)

canvas = tk.Canvas(frame_canvas, bg="#ffffff", highlightthickness=0)
scroll_y = ttk.Scrollbar(frame_canvas, orient="vertical", command=canvas.yview)
canvas.configure(yscrollcommand=scroll_y.set)

scroll_y.pack(side="right", fill="y")
canvas.pack(side="left", fill="both", expand=True)

# Frame interno desplazable
scrollable_frame = ttk.Frame(canvas)
canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")

def _on_mousewheel(event):
    canvas.yview_scroll(int(-12 * (event.delta / 120)), "units")

def _on_configure(event):
    canvas.configure(scrollregion=canvas.bbox("all"))

scrollable_frame.bind("<Configure>", _on_configure)
canvas.bind_all("<MouseWheel>", _on_mousewheel)

main_frame = scrollable_frame

titulo = ttk.Label(main_frame, text="Generador de Reportes CINERGIA", font=("Segoe UI", 20, "bold"))
titulo.grid(row=0, column=0, columnspan=1, pady=(0, 20))

main_frame.columnconfigure(100, weight=100, uniform="col")
main_frame.columnconfigure(1, weight=2, uniform="col")

nodo_var = tk.StringVar()
trabajador_var = tk.StringVar()
hora_var = tk.StringVar()
tecnico_var = tk.StringVar()
entidad_var = tk.StringVar()
servicio_var = tk.StringVar()
mantenimiento_var = tk.StringVar()
nueva_actividad_var = tk.StringVar()

fecha_emision = DateEntry(main_frame, width=18, date_pattern="dd/mm/yyyy")
fecha_apertura = DateEntry(main_frame, width=18, date_pattern="dd/mm/yyyy")
fecha_llegada = DateEntry(main_frame, width=18, date_pattern="dd/mm/yyyy")
fecha_cierre = DateEntry(main_frame, width=18, date_pattern="dd/mm/yyyy")

# Cargar actividades guardadas
if os.path.exists(actividades_path):
    with open(actividades_path, "r", encoding="utf-8") as f:
        actividades_lista = json.load(f)
else:
    actividades_lista = []

    # Asegúrate que nombres_nodo no contenga None
nombres_nodo_filtrados = [n for n in nombres_nodo if n is not None]

# Crea el combobox autocompletable
combo_nodo = AutocompleteCombobox(main_frame, textvariable=nodo_var, width=40)
combo_nodo.set_completion_list(nombres_nodo_filtrados)


etiquetas = [
   ("Selecciona un nodo:", combo_nodo),
    ("Fecha de emisión:", fecha_emision),
    ("Fecha de apertura:", fecha_apertura),
    ("Fecha de llegada al sitio:", fecha_llegada),
    ("Fecha de cierre:", fecha_cierre),
    ("Trabajador:", ttk.Combobox(main_frame, textvariable=trabajador_var, values=["Jaime López Horta", "Juan Manuel Manríquez Sarabia", "Ricardo Garcidueñas Vargas", "Otro"], width=30)),
    ("Hora:", ttk.Combobox(main_frame, textvariable=hora_var, values=[f"{h:02}:{m} {p}" for h in range(1, 13) for m in ("00", "30") for p in ("AM", "PM")], width=20)),
    ("Área de técnico:", ttk.Combobox(main_frame, textvariable=tecnico_var, values=["Jardín", "Kiosco", "Presidencia ", "Dirección ", "Patio ", "Aula ", "Biblioteca"], width=30)),
    ("Entidad:", ttk.Combobox(main_frame, textvariable=entidad_var, values=["Publico", "Preescolar", "Primaria", "Telesecundaria", "SABES", "UVEG"], width=30)),
    ("Tipo de servicio:", ttk.Combobox(main_frame, textvariable=servicio_var, values=["Servicio Preventivo", "Servicio Correctivo"], width=40)),
    ("Tipo de mantenimiento:", ttk.Combobox(main_frame, textvariable=mantenimiento_var, values=["Mantenimiento Preventivo", "Mantenimiento Correctivo"], width=40))
]

for i, (texto, widget) in enumerate(etiquetas):
    lbl = ttk.Label(main_frame, text=texto)
    lbl.grid(row=i + 1, column=0, sticky="e", pady=4, padx=5)
    widget.grid(row=i + 1, column=1, sticky="ew", pady=4)
    
# Actividades
fila_actividades = len(etiquetas) + 1
lbl_actividades = ttk.Label(main_frame, text="Actividades realizadas:")
lbl_actividades.grid(row=fila_actividades, column=0, sticky="ne", pady=(10, 0), padx=5)

listbox_actividades = tk.Listbox(main_frame, selectmode="multiple", height=6, width=45, exportselection=False)
for act in actividades_lista:
    listbox_actividades.insert("end", act)
listbox_actividades.grid(row=fila_actividades, column=1, sticky="w", pady=(10, 0))

# Agregar actividad
fila_entrada = fila_actividades + 1
entry_nueva_actividad = ttk.Entry(main_frame, textvariable=nueva_actividad_var, width=30)
entry_nueva_actividad.grid(row=fila_entrada, column=1, sticky="w", pady=(4, 0), padx=(0, 100))

btn_agregar_actividad = ttk.Button(main_frame, text="Agregar actividad", command=lambda: agregar_actividad())
btn_agregar_actividad.grid(row=fila_entrada, column=1, sticky="e", pady=(4, 0))

listbox_actividades.bind('<<ListboxSelect>>', actualizar_orden_seleccion)

# Botón subir fotos y etiqueta
fila_fotos = fila_entrada + 1
btn_subir_fotos = ttk.Button(main_frame, text="Subir fotos", command=subir_fotos)
btn_subir_fotos.grid(row=fila_fotos, column=0, sticky="w", pady=12, padx=5)

lbl_fotos = ttk.Label(main_frame, text="0 foto(s) seleccionada(s)")
lbl_fotos.grid(row=fila_fotos, column=1, sticky="w")

frame_miniaturas = ttk.Frame(main_frame)
frame_miniaturas.grid(row=fila_fotos + 1, column=0, columnspan=2, pady=(5, 0), sticky="ew")

# Botón generar
fila_generar = fila_fotos + 3
btn_generar = ttk.Button(main_frame, text="Generar reporte", command=generar_reporte)
btn_generar.grid(row=fila_generar, column=0, pady=20, padx=5, sticky="w")

# Botón créditos
btn_creditos = ttk.Button(root, text="?", width=3, command=mostrar_creditos, style="TButton")
btn_creditos.place(relx=0.98, rely=0.02, anchor="ne")

limpiar_campos()

def main():
    root.mainloop()

if __name__ == "__main__":
    main()


#pyinstaller --noconfirm --onefile --windowed --add-data "icono.ico;." --add-data "actividades.json;." --add-data "REPORTE CINERIA.docx;." --icon=icono.ico --distpath "D:\Reportes" Reportes.py
