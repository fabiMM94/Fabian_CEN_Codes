# import pyodbc
 
# DB_NAME = "test_access"
# SERVER_NAME = "10.5.54.108\\NB-CD-DEE2"  # O solo "192.168.1.100" si es instancia por defecto
# usuario = "Fabi.fabi"
# contrasena = "Fabino123_#"
 
# try:
#     conn = pyodbc.connect(
#         f"Driver={{ODBC Driver 17 for SQL Server}};"
#         f"Server={SERVER_NAME};"
#         f"Database={DB_NAME};"
#         f"UID={usuario};"
#         f"PWD={contrasena};"
#     )
#     print("✅ Conectado correctamente.")
# except Exception as e:
#     print(f"❌ Error al conectar: {e}")
 


import urllib

import pandas as pd

from sqlalchemy import create_engine
 
# Configuración de conexión

DB_NAME = "test_access"

#SERVER_NAME = "10.5.54.108\\NB-CD-DEE2"
#SERVER_NAME = "10.5.54.108\\SQLEXPRESS"
#SERVER_NAME = "10.5.54.108"
SERVER_NAME = "10.5.54.108\\SQLEXPRESS"
usuario = "Fabi.fabi"

contrasena = "Fabino123_#"
 
params = urllib.parse.quote_plus(
    f"Driver={{ODBC Driver 17 for SQL Server}};"
    f"Server={SERVER_NAME};"
    f"Database={DB_NAME};"
    f"UID={usuario};"
    f"PWD={contrasena};"

)

connection_string = f"mssql+pyodbc:///?odbc_connect={params}"
engine = create_engine(connection_string, echo=True)
# Ejecutar query
try:
    print("Comienza la query")
    query = "SELECT * FROM CompanyT"
    df = pd.read_sql(query, engine)
    print("✅ Datos extraídos con éxito:")
    print(df.head())

except Exception as e:
    print(f"❌ Error al ejecutar la query: {e}")
 