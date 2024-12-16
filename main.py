from modules.date_selector import open_date_selector
from orders import get_orders
from payments import get_payments
from renewalsAndNoRecurrents import get_sales

start_date, end_date, folder_name, all_var, orders_var, unique_orders_var, sales_var, payment_errors_var = open_date_selector()

if(all_var == 1 | orders_var == 1):
    unique_orders_var = [1, 1, 1, 1, 1, 1, 1, 1, 1]

get_orders(start_date, end_date, folder_name, unique_orders_var)

if(sales_var == 1):
    get_sales(start_date, end_date, folder_name)

if(payment_errors_var == 1):
    get_payments(start_date, end_date, folder_name) 