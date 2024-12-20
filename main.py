from ga4Funnels import get_funnel
from modules.date_selector import open_date_selector
from orders import get_orders
from payments import get_payments
from renewalsAndNoRecurrents import get_sales
from report import anotar_datos_excel, seleccionar_tipo_de_reporte
from selectFiles import seleccionar_archivos_para_casos

funnels_report = True
database_report = True
folder_name = 'funnels'
columna = 4

funnels_report, database_report = seleccionar_tipo_de_reporte()

if funnels_report:
    archivos, mes = seleccionar_archivos_para_casos()
    columna = mes + 3


if database_report:
    start_date, end_date, folder_name, all_var, orders_var, unique_orders_var, sales_var, payment_errors_var = open_date_selector()

    values, items = get_orders(start_date, end_date, folder_name, unique_orders_var) 

    anotar_datos_excel(values, columna, 35)
    anotar_datos_excel(items, columna, 45)

    if(sales_var == 1):
        total_sales = get_sales(start_date, end_date, folder_name)
        anotar_datos_excel(total_sales, columna, 57)

    if(payment_errors_var == 1):
        total_payments = get_payments(start_date, end_date, folder_name) 
        anotar_datos_excel(total_payments, columna, 62)

if funnels_report:
    if archivos['Customized Kit - Funnel'] != None:
        get_funnel(archivos['Customized Kit - Funnel'], 'Customized Kit - Funnel.xlsx', columna, 6, folder_name)

    if archivos['All In One - Funnel'] != None:
        get_funnel(archivos['All In One - Funnel'], 'All In One - Funnel.xlsx', columna, 12, folder_name)

    if archivos['Shop - Funnel'] != None:
        get_funnel(archivos['Shop - Funnel'], 'Shop - Funnel.xlsx', columna, 17, folder_name)

    if archivos['My Account - Funnel'] != None:
        get_funnel(archivos['My Account - Funnel'], 'My Account - Funnel.xlsx', columna, 22, folder_name)

    if archivos['Buy Again - Funnel'] != None:
        get_funnel(archivos['Buy Again - Funnel'], 'Buy Again - Funnel.xlsx', columna, 26, folder_name)

    if archivos['My Subscriptions - Funnel'] != None:
        get_funnel(archivos['My Subscriptions - Funnel'], 'My Subscriptions - Funnel.xlsx', columna, 31, folder_name)


