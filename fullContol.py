import pandas as pd
import mysql.connector
import matplotlib.pyplot as plt
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.drawing.image import Image
from openpyxl.styles import PatternFill
from tkinter import Tk, Label, Button, Entry
from tkinter import messagebox
from tkcalendar import Calendar
from modules.colors import lighten_color
from modules.database_queries import execute_query
from modules.date_selector import open_date_selector
from modules.excel_creator import save_dataframe_to_excel, save_dataframe_to_excel_orders
from report import anotar_datos_excel

def fullControl(start_date, end_date):

    query_total = f"""
    SELECT 
        v.subscription_id, v.first_date_sms_renewal_true AS full_control_starting_date, COUNT(r.id) AS sucess_renewal_count_after_first, GROUP_CONCAT(r.createdAt ORDER BY r.createdAt ASC SEPARATOR ', ') AS success_renewal_dates_after_first
    FROM 
        sales_and_subscriptions.first_sms_renewal_versions v
    LEFT JOIN 
        sales_and_subscriptions.renewals r
        ON v.subscription_id = r.subscriptionId
        AND r.createdAt > v.first_date_sms_renewal_true
        AND r.status = 'SUCCESS'  -- Filtro para incluir solo renovaciones con status 'SUCCESS'
        AND r.createdAt < "{end_date} 08:00:00" AND r.createdAt >= "{start_date} 08:00:00"
        
    -- WHERE v.first_date_sms_renewal_true > "2025-04-01 08:00:00" AND v.first_date_sms_renewal_true < "2025-05-01 08:00:00"
    GROUP BY 
        v.subscription_id, v.first_date_sms_renewal_true;
    """

    query_new = f"""
    SELECT 
        v.subscription_id, v.first_date_sms_renewal_true AS full_control_starting_date, COUNT(r.id) AS sucess_renewal_count_after_first, GROUP_CONCAT(r.createdAt ORDER BY r.createdAt ASC SEPARATOR ', ') AS success_renewal_dates_after_first
    FROM 
        sales_and_subscriptions.first_sms_renewal_versions v
    LEFT JOIN 
        sales_and_subscriptions.renewals r
        ON v.subscription_id = r.subscriptionId
        AND r.createdAt > v.first_date_sms_renewal_true
        AND r.status = 'SUCCESS'  -- Filtro para incluir solo renovaciones con status 'SUCCESS'
        AND r.createdAt < "{end_date} 08:00:00" AND r.createdAt >= "{start_date} 08:00:00"
        
    WHERE v.first_date_sms_renewal_true > "{start_date} 08:00:00" AND v.first_date_sms_renewal_true < "{end_date} 08:00:00"
    GROUP BY 
        v.subscription_id, v.first_date_sms_renewal_true;
    """

    fct = execute_query(query_total)
    fcn = execute_query(query_new)

    total_subs = fct['subscription_id'].count()
    new_subs = fcn['subscription_id'].count()

    percentage_new_subs = round(new_subs/total_subs*100, 2)

    percentage_array = [f"{percentage_new_subs}%"]

    return percentage_array












