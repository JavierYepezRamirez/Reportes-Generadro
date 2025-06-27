# Sistema de licencias
# Permite que una como putadora tenga acceso por medio del ID de la maquina 
# si no tiene hacesso no la dejara, pero si si pasara al Reportes.py
# @Version 1.0 
# 25/06/2025
# By: Javier Yepez Ramirez

import os
import sys
import firebase_admin
from firebase_admin import credentials, firestore
import tkinter as tk
from tkinter import messagebox
from PIL import Image, ImageTk  # pip install pillow
from utils import obtener_id_maquina
import Reportes

def iniciar_firebase():
    try:
        ruta_clave = os.path.join(os.path.dirname(__file__), "firebase_key.json")
        cred = credentials.Certificate(ruta_clave)
        firebase_admin.initialize_app(cred)
        return firestore.client()
    except Exception as e:
        messagebox.showerror("Error", f"Error iniciando Firebase:\n{e}")
        sys.exit(1)

def licencia_valida(db, id_maquina):
    licencias = db.collection("licencias").where("id_maquina", "==", id_maquina).get()
    for doc in licencias:
        data = doc.to_dict()
        if data.get("activa", False):
            return True
    return False

def registrar_nueva_licencia(db, id_maquina, usuario):
    licencias = db.collection("licencias").where("id_maquina", "==", id_maquina).get()
    if not licencias:
        db.collection("licencias").add({
            "id_maquina": id_maquina,
            "usuario": usuario,
            "activa": False
        })
        messagebox.showinfo("Licencia", "⚠ ID no registrado. Se ha enviado solicitud de activación.")
    else:
        messagebox.showwarning("Licencia", "⚠ Licencia existente pero inactiva.")

def pedir_usuario_con_logo(root, icon_path, logo_path):
    ventana = tk.Toplevel(root)
    ventana.title("Ingreso de Usuario")
    ventana.iconbitmap(icon_path)
    ventana.geometry("400x250")
    ventana.resizable(False, False)

    # Cargar y mostrar logo
    try:
        img_orig = Image.open(logo_path)
        img_resized = img_orig.resize((100, 100), Image.Resampling.LANCZOS)
        img_tk = ImageTk.PhotoImage(img_resized)
        lbl_logo = tk.Label(ventana, image=img_tk)
        lbl_logo.image = img_tk  # Mantener referencia
        lbl_logo.pack(pady=10)
    except Exception as e:
        print(f"No se pudo cargar el logo: {e}")

    lbl_texto = tk.Label(ventana, text="Por favor, ingrese su nombre o usuario:", font=("Segoe UI", 12))
    lbl_texto.pack()

    entrada = tk.Entry(ventana, font=("Segoe UI", 12))
    entrada.pack(pady=5)
    entrada.focus_set()

    usuario = []

    def aceptar():
        val = entrada.get().strip()
        if not val:
            messagebox.showwarning("Advertencia", "Debe ingresar un usuario.", parent=ventana)
            return
        usuario.append(val)
        ventana.destroy()

    btn_aceptar = tk.Button(ventana, text="Aceptar", command=aceptar, font=("Segoe UI", 12))
    btn_aceptar.pack(pady=15)

    ventana.grab_set()
    root.wait_window(ventana)

    return usuario[0] if usuario else None

def main():
    root = tk.Tk()
    root.withdraw()

    # Rutas de icono y logo, ajusta estas rutas a tus archivos
    icono_path = os.path.join(os.path.dirname(__file__), "icono.ico")
    logo_path = os.path.join(os.path.dirname(__file__), "logo.png")

    # Asignar icono a la ventana oculta (opcional)
    try:
        root.iconbitmap(icono_path)
    except Exception as e:
        print(f"No se pudo asignar icono a root: {e}")

    db = iniciar_firebase()
    id_maquina = obtener_id_maquina()
    print("ID de máquina:", id_maquina)

    if licencia_valida(db, id_maquina):
        root.destroy()
        print("✅ Licencia válida. Ejecutando Reportes.py...")
        import Reportes
        Reportes.main()
    else:
        usuario = pedir_usuario_con_logo(root, icono_path, logo_path)
        if not usuario:
            messagebox.showwarning("Usuario", "No se ingresó usuario. El programa se cerrará.", parent=root)
            root.destroy()
            return

        registrar_nueva_licencia(db, id_maquina, usuario)
        messagebox.showerror("Licencia", "Licencia no válida. Contacte al administrador.", parent=root)
        root.destroy()

if __name__ == "__main__":
    main()

# pyinstaller --noconfirm --onefile --windowed --add-data "icono.ico;." --add-data "logo.png;." --add-data "actividades.json;." --add-data "REPORTE CINERIA.docx;." --add-data "firebase_key.json;." --icon=icono.ico --distpath "C:\Users\javie\Desktop\Reportes" index.py
