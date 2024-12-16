from tkinter import simpledialog
import pandas as pd
import os
import tkinter as tk
from tkinter import filedialog

# Función para seleccionar archivos CSV
def seleccionar_archivos():
    root = tk.Tk()
    root.withdraw()  # Oculta la ventana principal
    archivos = filedialog.askopenfilenames(title="Seleccionar archivos CSV", filetypes=[("CSV files", "*.csv")])
    return archivos

# Función para combinar archivos CSV
def combinar_csv(archivos):
    dfs = []
    for i, archivo in enumerate(archivos):
        try:
            # Leer todas las filas del archivo, pero ignorar las primeras 6
            df = pd.read_csv(archivo, encoding='utf-8', skiprows=6)  
            if i == 0:
                df = df.iloc[6:]  # Eliminar 6 filas del primer archivo
            else:
                df = df.iloc[7:]  # Eliminar 7 filas de otros archivos
            dfs.append(df)
        except pd.errors.ParserError as e:
            print(f"Error leyendo el archivo {archivo}: {e}")
    if dfs:  # Chequear si hay DataFrames válidos antes de concatenar
        combined_df = pd.concat(dfs, ignore_index=True)
        return combined_df
    else:
        return pd.DataFrame()  # Retornar un DataFrame vacío si no hay objetos para concatenar

# Main Execution
def leer_archivos_csv():
    archivos_csv = seleccionar_archivos()
    if archivos_csv:
        data_combined = combinar_csv(archivos_csv)
        if not data_combined.empty:  # Verificar si el DataFrame combinado no está vacío
            nombre_salida = simpledialog.askstring("Guardar como", "Ingrese el nombre del archivo (sin extensión):")
            if nombre_salida:
                nombre_salida += ".xlsx"
                return data_combined, nombre_salida
        else:
            print("No se pudieron combinar los archivos correctamente.")
    else:
        print("No se seleccionaron archivos.")
