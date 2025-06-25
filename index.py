import json
import tkinter as tk
from tkinter import messagebox
import subprocess
import os
import sys

# === Funciones de autenticación ===
def cargar_credenciales():
    if not os.path.exists('credenciales.json'):
        messagebox.showerror("Error", "Archivo credenciales.json no encontrado.")
        return {}
    with open('credenciales.json', 'r') as f:
        return json.load(f)

def validar_credenciales(usuario, password):
    datos = cargar_credenciales()
    return usuario in datos and datos[usuario]["password"] == password

# === Acciones del botón de login ===
def iniciar_sesion():
    user = entry_usuario.get()
    pwd = entry_password.get()

    if validar_credenciales(user, pwd):
        messagebox.showinfo("Acceso concedido", f"Bienvenido, {user}")
        root.destroy()

        # Ejecuta Reportes.py como proceso aparte
        python_exe = sys.executable
        ruta_script = os.path.join(os.path.dirname(__file__), "Reportes.py")
        subprocess.Popen([python_exe, ruta_script])
    else:
        messagebox.showerror("Acceso denegado", "Credenciales incorrectas.")

# === Interfaz de login ===
root = tk.Tk()
root.title("Login - Reportes CINERGIA")
root.geometry("300x200")
root.resizable(False, False)

tk.Label(root, text="Usuario:").pack(pady=(20, 5))
entry_usuario = tk.Entry(root)
entry_usuario.pack()

tk.Label(root, text="Contraseña:").pack(pady=5)
entry_password = tk.Entry(root, show="*")
entry_password.pack()

btn_login = tk.Button(root, text="Iniciar sesión", command=iniciar_sesion)
btn_login.pack(pady=20)

root.mainloop()
