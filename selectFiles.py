from tkinter import filedialog, messagebox, simpledialog
import pandas as pd
import tkinter as tk

def seleccionar_archivos():
    root = tk.Tk()
    root.withdraw()  # Oculta la ventana principal
    archivos = filedialog.askopenfilenames(title="Seleccionar archivos CSV", filetypes=[("CSV files", "*.csv")])
    if not archivos:  # Si no se selecciona archivo
        raise ValueError("No se seleccionó ningún archivo.")
    elif len(archivos) > 1:  # Si se seleccionan múltiples archivos
        raise ValueError("Por favor, selecciona solo un archivo.")
    ruta_archivo = archivos[0]  # Tomar el primer archivo seleccionado
    return pd.read_csv(ruta_archivo, encoding='utf-8', skiprows=9)

def nombre_archivo():
    nombre_salida = simpledialog.askstring("Guardar como", "Ingrese el nombre del archivo (sin extensión):")
    if nombre_salida:
        nombre_salida += ".xlsx"
        return nombre_salida
    else: 
        messagebox.showerror("Error", "Por favor ingresa un nombre.")
        return ''
    