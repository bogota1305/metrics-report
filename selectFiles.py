from tkinter import filedialog, messagebox, simpledialog, Label, Button, StringVar
import pandas as pd
import tkinter as tk

def seleccionar_archivos_para_casos():
    # Diccionario para almacenar las rutas de los archivos seleccionados
    archivos_seleccionados = {
        "Customized Kit - Funnel": None,
        "All In One - Funnel": None,
        "Shop - Funnel": None,
        "My Account - Funnel": None,
        "Buy Again - Funnel": None,
        "My Subscriptions - Funnel": None,
    }

    def seleccionar_archivo(caso, label):
        archivo = filedialog.askopenfilename(title=f"Seleccionar: {caso}",
                                             filetypes=[("CSV files", "*.csv")])
        if archivo:
            archivos_seleccionados[caso] = archivo
            label.config(text=f"Seleccionado: {archivo}")

    # Crear ventana principal
    root = tk.Tk()
    root.title("Seleccionar archivos para cada caso")
    
    # Variable para almacenar la selección del mes
    mes_seleccionado = StringVar(value="Primer mes") 

    # Crear botones y etiquetas para cada caso
    for caso in archivos_seleccionados.keys():
        frame = tk.Frame(root)
        frame.pack(pady=5, padx=10, anchor="w")
        
        label = Label(frame, text=f"Selecciona: {caso}", wraplength=500, justify="left")
        label.pack(side="left", padx=10)
        
        boton = Button(frame, text=f"Seleccionar archivo para {caso}",
                       command=lambda c=caso, l=label: seleccionar_archivo(c, l))
        boton.pack(side="right")

    # Espaciado entre secciones
    tk.Label(root, text="").pack()

    # Sección para seleccionar el mes
    tk.Label(root, text="Seleccione el mes:").pack(pady=5)

    # Botones de radio para la selección del mes
    tk.Radiobutton(root, text="Primer mes", variable=mes_seleccionado, value="Primer mes").pack(anchor="w")
    tk.Radiobutton(root, text="Segundo mes", variable=mes_seleccionado, value="Segundo mes").pack(anchor="w")

    # Botón para confirmar selección
    confirmar = Button(root, text="Confirmar selección", command=root.quit)
    confirmar.pack(pady=20)

    # Mostrar ventana
    root.mainloop()

    # Cerrar ventana
    root.destroy()

    mes = 2
    if(mes_seleccionado.get() == 'Primer mes'):
        mes = 1

    # Retornar los archivos seleccionados y el mes elegido
    return archivos_seleccionados, mes