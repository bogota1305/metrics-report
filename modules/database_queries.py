import pandas as pd
import mysql.connector
from tkinter import Tk, Label, Entry, Button

# Variables globales para la conexión
host = 'prod0.db.becleverman.com'
user = 'prodproductuser'
password = 'YVE1RbUidDBk3qMHZUlLDSVPEqHN9y3YCrmu5b1GgwGcXy9V8A'

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


