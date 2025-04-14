from block_payments import get_blocked_payments
from exceptedRenewals import get_expected_renewals
from ga4Funnels import get_funnel
from modules.date_selector import open_date_selector
from orders import get_orders
from payments import get_payments
from realRenewalFrecuency import realRenewalFrequency
from renewalsAndNoRecurrents import get_sales
from report import anotar_datos_excel, seleccionar_donde_almacenar, seleccionar_tipo_de_reporte
from selectFiles import seleccionar_archivos_para_casos, seleccionar_archivos_stripe
from uploadCloud import upload_to_dropbox

funnels_report = True
database_report = True
stripe_block_payments = True
folder_name = 'funnels'
columna = 4
dropbox_var = False

funnels_report, database_report, stripe_block_payments = seleccionar_tipo_de_reporte()

if funnels_report:
    archivos, mes = seleccionar_archivos_para_casos()
    columna = mes + 3
    
    if(database_report == False):
        dropbox_var, drive_var = seleccionar_donde_almacenar()

if database_report:
    start_date, end_date, folder_name, all_var, orders_var, unique_orders_var, sales_var, payment_errors_var, expected_renewals_var, frequency_var = open_date_selector()

    dropbox_var, drive_var = seleccionar_donde_almacenar()

    values, items, urls = get_orders(start_date, end_date, folder_name, unique_orders_var, dropbox_var) 

    anotar_datos_excel(values, columna, 36)
    anotar_datos_excel(items, columna, 46)

    if(dropbox_var):
        anotar_datos_excel(urls, columna, 36, True)
        anotar_datos_excel(urls, columna, 46, True)

    if(sales_var == 1):
        total_sales, urls = get_sales(start_date, end_date, folder_name, dropbox_var)
        anotar_datos_excel(total_sales, columna, 58) 

        if(dropbox_var):
            anotar_datos_excel(urls, columna, 58, True)

    if(payment_errors_var == 1):
        total_payments, urls = get_payments(start_date, end_date, folder_name, dropbox_var) 
        anotar_datos_excel(total_payments, columna, 63)

        if(dropbox_var):
            anotar_datos_excel(urls, columna, 63, True)

    if(expected_renewals_var == 1):
        total_expected_renewals = get_expected_renewals(start_date, end_date, folder_name) 
    
    if(frequency_var == 1):
        realRenewalFrequency(start_date, end_date, folder_name) 

if funnels_report:
    if archivos['Customized Kit - Funnel'] != None:
        get_funnel(archivos['Customized Kit - Funnel'], 'Customized Kit - Funnel.xlsx', columna, 5, folder_name, dropbox_var)

    if archivos['All In One - Funnel'] != None:
        get_funnel(archivos['All In One - Funnel'], 'All In One - Funnel.xlsx', columna, 12, folder_name, dropbox_var)

    if archivos['Shop - Funnel'] != None:
        get_funnel(archivos['Shop - Funnel'], 'Shop - Funnel.xlsx', columna, 17, folder_name, dropbox_var)

    if archivos['My Account - Funnel'] != None:
        get_funnel(archivos['My Account - Funnel'], 'My Account - Funnel.xlsx', columna, 22, folder_name, dropbox_var)

    if archivos['Buy Again - Funnel'] != None:
        get_funnel(archivos['Buy Again - Funnel'], 'Buy Again - Funnel.xlsx', columna, 26, folder_name, dropbox_var)

    if archivos['My Subscriptions - Funnel'] != None:
        get_funnel(archivos['My Subscriptions - Funnel'], 'My Subscriptions - Funnel.xlsx', columna, 31, folder_name, dropbox_var)

if stripe_block_payments:
    archivos = seleccionar_archivos_stripe() 
    if archivos['Blocked Payments'] != None and archivos['All Payments'] != None:
        get_blocked_payments(archivos['Blocked Payments'], archivos['All Payments'], 'Blocked payments', 'Stripe data')

if(dropbox_var):
    upload_to_dropbox("metricas.xlsx", dropbox_path=f"/MyReports/{folder_name}/metricas.xlsx") 