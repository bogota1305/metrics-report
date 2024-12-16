import pandas as pd
import matplotlib.pyplot as plt

from readCSV import leer_archivos_csv

# Carga los datos del archivo Excel
ruta_archivo = "novFunnel.xlsx"
data, nombre_archivo = leer_archivos_csv()

# Eliminar columnas innecesarias
data_filtered = data.drop(columns=['Completion rate', 'Abandonments', 'Abandonment rate'])

# Separar las filas donde 'Day' es "Total" para manejarlo aparte
data_total = data_filtered[data_filtered['Day'] == "Total"]
data_days = data_filtered[data_filtered['Day'] != "Total"]

# Convertir los valores numéricos en 'Day' a enteros para asegurar un orden correcto
data_days['Day'] = pd.to_numeric(data_days['Day'], errors='coerce')

# Crear una tabla pivotada para los días numéricos
pivot_days = data_days.pivot_table(
    index='Step',
    columns='Day',
    values='Active users',
    aggfunc='sum'
)

# Ordenar las columnas por los días en orden ascendente
pivot_days = pivot_days.sort_index(axis=1)

# Agregar una columna "Total" que sume los valores de todos los días
pivot_days['Total'] = pivot_days.sum(axis=1)

# Eliminar las filas específicas del DataFrame pivot_days
steps_to_remove = [
    "1. Ingreso"
]

pivot_days_filtered = pivot_days.drop(index=steps_to_remove, errors='ignore')

# Limpiar los nombres de los pasos
pivot_days_filtered.index = pivot_days_filtered.index.str.replace(r'^\d+\.\s*', '', regex=True)

# Identificar la primera fila como referencia para calcular los porcentajes
reference_row = pivot_days_filtered.iloc[0]  # Seleccionar la primera fila
reference_row_name = pivot_days_filtered.index[0]  # Guardar el nombre de la fila de referencia

# Calcular la tabla de porcentajes basada en la primera fila
percentages_table = pivot_days_filtered.div(reference_row) * 100

# Renombrar las filas para diferenciar las tablas
percentages_table.index = [f"{step} (%)" for step in percentages_table.index]

# Concatenar la tabla original con la tabla de porcentajes
final_table_with_percentages = pd.concat([pivot_days_filtered, percentages_table])

# Separar las tablas de números y porcentajes
counts_table = final_table_with_percentages[~final_table_with_percentages.index.str.contains(r'\(%\)', regex=True)]

percentages_table = final_table_with_percentages[final_table_with_percentages.index.str.contains(r'\(%\)', regex=True)]

# Excluir la primera fila de la tabla de porcentajes (que representa el 100%)
percentages_table = percentages_table[percentages_table.index != f"{reference_row_name} (%)"]

# Formatear la tabla de porcentajes como valores porcentuales
percentages_table = percentages_table.applymap(lambda x: f"{x:.2f}%" if pd.notnull(x) else "")

# Concatenar las tablas con una fila de espacio en blanco entre ambas
empty_row = pd.DataFrame([["" for _ in counts_table.columns]], columns=counts_table.columns)
final_table_spaced = pd.concat([counts_table, empty_row, percentages_table])

# Calcular los porcentajes relativos al paso anterior
percentages_previous_step = pivot_days_filtered.copy()

for column in percentages_previous_step.columns:
    percentages_previous_step[column] = percentages_previous_step[column] / percentages_previous_step[column].shift(1) 
    percentages_previous_step[column] *= 100

# Reemplazar valores indefinidos (NaN) y negativos con 0%
percentages_previous_step = percentages_previous_step.fillna(0).clip(lower=0)

# Formatear la tabla de porcentajes como valores porcentuales
percentages_previous_step = percentages_previous_step.applymap(
    lambda x: f"{x:.2f}%" if pd.notnull(x) else "")
    
# Renombrar las filas para indicar el paso anterior
steps_with_previous = percentages_previous_step.index.tolist()
steps_with_previous = [
    f"{steps_with_previous[i]} (vs '{steps_with_previous[i - 1]}')"
    if i > 0 else steps_with_previous[i] for i in range(len(steps_with_previous))
]
percentages_previous_step.index = steps_with_previous

# Excluir la primera fila ya que no tiene paso anterior
percentages_previous_step = percentages_previous_step.iloc[1:]

# Concatenar la nueva tabla con un espacio en blanco al final de las demás tablas
empty_row_previous = pd.DataFrame(
    [["" for _ in pivot_days_filtered.columns]], columns=pivot_days_filtered.columns)
final_table_spaced_with_previous = pd.concat(
    [final_table_spaced, empty_row, percentages_previous_step]
)

# Crear una gráfica a partir de la tabla de porcentajes, excluyendo la columna "Total"
plt.figure(figsize=(10, 6))
for step in percentages_table.index:
    numeric_values = percentages_table.loc[step].drop('Total', errors='ignore').replace('%', '', regex=True).astype(float)
    plt.plot(numeric_values, label=step.replace(' (%)', ''))

plt.title('Percentage Transition by Step')
plt.xlabel('Days')
plt.ylabel('Percentage')
plt.legend(loc='upper left', bbox_to_anchor=(1, 1))
plt.grid(True)
plt.tight_layout()

# Guardar la gráfica en un archivo temporal
chart_path = 'firstStep.png'
plt.savefig(chart_path)
plt.close()

# Crear una gráfica a partir de la tabla de porcentajes, excluyendo la columna "Total"
plt.figure(figsize=(10, 6))
for step in percentages_previous_step.index:
    numeric_values = percentages_previous_step.loc[step].drop('Total', errors='ignore').replace('%', '', regex=True).astype(float)
    plt.plot(numeric_values, label=step.replace(' (%)', ''))

plt.title('Percentage Step vs Previous')
plt.xlabel('Days')
plt.ylabel('Percentage')
plt.legend(loc='upper left', bbox_to_anchor=(1, 1))
plt.grid(True)
plt.tight_layout()

# Guardar la gráfica en un archivo temporal
chart_path_previous = 'stepPrevious.png'
plt.savefig(chart_path_previous)
plt.close()

# Guardar la tabla final y la gráfica en un archivo Excel
with pd.ExcelWriter(nombre_archivo, engine='xlsxwriter') as writer:
    final_table_spaced_with_previous.to_excel(writer, sheet_name='Data', index=True, startrow=0, startcol=0)
    workbook = writer.book
    worksheet = writer.sheets['Data']

    # Ajustar el ancho de la columna de los pasos
    max_step_length = max(len(str(step)) for step in final_table_spaced_with_previous.index)
    worksheet.set_column(0, 0, max_step_length + 2)

    # Insertar la gráfica en el archivo Excel
    worksheet.insert_image('B18', chart_path)
    # Insertar la gráfica en el archivo Excel
    worksheet.insert_image('R18', chart_path_previous)
