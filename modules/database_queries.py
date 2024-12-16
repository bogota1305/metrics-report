import pandas as pd
import mysql.connector
from tkinter import Tk, Label, Entry, Button

# Variables globales para la conexión
host = ''
user = ''
password = ''

def execute_query(query):
    """Ejecuta una consulta SQL y devuelve un DataFrame."""
    global host, user, password
    db_config = {
        'host': host,
        'user': user,
        'password': password
    }
    connection = mysql.connector.connect(**db_config)
    data = pd.read_sql(query, connection)
    connection.close()
    return data


