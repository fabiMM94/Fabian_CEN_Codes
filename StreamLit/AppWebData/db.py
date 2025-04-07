# db.py
import sqlite3
import pandas as pd

DB_FILE = "database.db"

class BaseDeDatos:
    def __init__(self):
        self.db_file = DB_FILE

    # Función para obtener datos de la base de datos
    def obtener_datos(self, query="SELECT * FROM empleados"):
        conn = sqlite3.connect(self.db_file)
        df = pd.read_sql(query, conn)
        conn.close()
        return df