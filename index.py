# Login
# Logeo para acceder a Reportes.py
# @Version 1.0 
# 25/06/2025
# By: Javier Yepez Ramirez

import json
import tkinter as tk
from tkinter import messagebox
import os
import sys

# Importamos Reportes solo cuando sea necesario para evitar que se abra ventana al importar
Reportes = None

def get_path(relative_path):
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath("."), relative_path)

def cargar_credenciales():
    ruta_json = os.path.join(os.path.dirname(__file__), 'credenciales.json')
    if not os.path.exists(ruta_json):
        messagebox.showerror("Error", "Archivo credenciales.json no encontrado.")
        return {}
    with open(ruta_json, 'r') as f:
        return json.load(f)

def validar_credenciales(usuario, password):
    datos = cargar_credenciales()
    return usuario in datos and datos[usuario]["password"] == password

login_exitoso = False

def iniciar_sesion():
    global login_exitoso, Reportes
    user = entry_usuario.get()
    pwd = entry_password.get()

    if validar_credenciales(user, pwd):
        messagebox.showinfo("Acceso concedido", f"Bienvenido, {user}")
        login_exitoso = True
        root.destroy()  # Cerrar ventana login

        # Importar Reportes aquí para que no se ejecute al inicio
        import Reportes
        Reportes.main()
    else:
        messagebox.showerror("Acceso denegado", "Credenciales incorrectas.")

root = tk.Tk()
root.title("Login - Reportes CINERGIA")
root.geometry("350x250")
root.resizable(False, False)

icon_path = get_path("icono.ico")
if os.path.exists(icon_path):
    try:
        root.iconbitmap(icon_path)
    except Exception as e:
        print(f"No se pudo cargar el icono: {e}")
else:
    print("El archivo icono.ico no fue encontrado.")

tk.Label(root, text="Usuario:").pack(pady=(20, 5))
entry_usuario = tk.Entry(root)
entry_usuario.pack()

tk.Label(root, text="Contraseña:").pack(pady=5)
entry_password = tk.Entry(root, show="*")
entry_password.pack()

btn_login = tk.Button(root, text="Iniciar sesión", command=iniciar_sesion)
btn_login.pack(pady=20)

root.mainloop()

# pyinstaller --noconfirm --onefile --windowed --add-data "icono.ico;." --add-data "credenciales.json;." --add-data "actividades.json;." --add-data "REPORTE CINERIA.docx;." --icon=icono.ico --distpath "C:\Users\javie\Desktop\Reportes" --name "Reportes CINERGIA" index.py
