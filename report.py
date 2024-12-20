from openpyxl import load_workbook
import tkinter as tk
from tkinter import BooleanVar, Checkbutton, Button

# Cargar el archivo Excel existente
archivo_excel = "metricas.xlsx" 

hoja_nombre = "Hoja1"  #

def anotar_datos_excel(datos, columna_inicio, fila_inicio):
    
    try:
        # Intentar cargar el archivo Excel existente
        wb = load_workbook(archivo_excel)
    except FileNotFoundError:
        print(f"El archivo '{archivo_excel}' no existe.")
        return

    # Seleccionar o crear la hoja donde escribirás los datos
    if hoja_nombre in wb.sheetnames:
        ws = wb[hoja_nombre]
    else:
        print(f"La hoja '{hoja_nombre}' no existe en el archivo.")
        return

    # Escribir los datos en las celdas
    for i, valor in enumerate(datos, start=fila_inicio):
        ws.cell(row=i, column=columna_inicio, value=valor)

    # Guardar los cambios en el archivo (nuevo o existente)
    wb.save(archivo_excel)

def seleccionar_tipo_de_reporte():
    """
    Muestra una ventana para que el usuario seleccione el tipo de reporte a generar.

    :return: tuple (funnels_report, database_report) - Estado de los checkboxes seleccionados por el usuario.
    """
    # Variables para almacenar el estado de los checkboxes
    root = tk.Tk()
    root.title("Seleccionar Tipo de Reporte")

    funnels_var = BooleanVar(value=True)
    database_var = BooleanVar(value=True)

    # Etiqueta principal
    label = tk.Label(root, text="Seleccione el tipo de reporte")
    label.pack(pady=10)

    # Checkbox para Funnels
    funnel_checkbox = Checkbutton(root, text="Funnels", variable=funnels_var)
    funnel_checkbox.pack(anchor=tk.W, padx=20)

    # Checkbox para Base de Datos
    database_checkbox = Checkbutton(root, text="Base de Datos", variable=database_var)
    database_checkbox.pack(anchor=tk.W, padx=20)

    # Función para cerrar la ventana y devolver los valores seleccionados
    def continuar():
        root.quit()
        root.destroy()

    # Botón para continuar
    continuar_button = Button(root, text="Continuar", command=continuar)
    continuar_button.pack(pady=20)

    root.mainloop()

    return funnels_var.get(), database_var.get()


