# Sistema de licencias
# Permite que una como putadora tenga acceso por medio del ID de la maquina 
# si no tiene hacesso no la dejara, pero si si pasara al Reportes.py
# @Version 1.0 
# 25/06/2025
# By: Javier Yepez Ramirez

import os
import firebase_admin
from firebase_admin import credentials, firestore
import tkinter as tk
from tkinter import simpledialog, messagebox
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
        exit(1)

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

def pedir_usuario(root):
    usuario = simpledialog.askstring("Ingreso de Usuario", "Por favor, ingrese su nombre o usuario:", parent=root)
    return usuario

def main():
    root = tk.Tk()
    root.withdraw()

    db = iniciar_firebase()
    id_maquina = obtener_id_maquina()
    print("ID de máquina:", id_maquina)

    if licencia_valida(db, id_maquina):
        root.destroy()  # Cierra la ventana oculta ANTES de importar Reportes
        print("✅ Licencia válida. Ejecutando Reportes.py...")
        import Reportes  # Importa Reportes aquí, luego de cerrar root
        Reportes.main()
    else:
        usuario = pedir_usuario(root)
        if not usuario:
            messagebox.showwarning("Usuario", "No se ingresó usuario. El programa se cerrará.", parent=root)
            root.destroy()
            return

        registrar_nueva_licencia(db, id_maquina, usuario)
        messagebox.showerror("Licencia", "Licencia no válida. Contacte al administrador.", parent=root)
        root.destroy()

if __name__ == "__main__":
    main()


# pyinstaller --noconfirm --onefile --windowed --add-data "icono.ico;." --add-data "actividades.json;." --add-data "REPORTE CINERIA.docx;." --add-data "firebase_key.json;." --icon=icono.ico --distpath "C:\Users\javie\Desktop\Reportes" index.py
