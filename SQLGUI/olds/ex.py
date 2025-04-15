
# import urllib
# from sqlalchemy import create_engine
# who_is = "Cris2"
# users = {"Hayter": ["NEWCIM", "localhost"], "Cris1": ["NEWCIM", "CBARRERA29N7\SQLEXPRESS"], "Cris2": ["SCADA_2", "NB-CD-DEE2"]}
# user = users[who_is]

# usuario = "Fabi.fabi"
# contrasena = "Fabifabiño123_-@sjk,chba<siclawr83746rq3efdqwl.kjsh!!!!"
 
# DB_NAME = user[0]
# SERVER_NAME = user[1]
# params = urllib.parse.quote_plus(
#     f"Driver={{ODBC Driver 17 for SQL Server}};"
#     f"Server={SERVER_NAME};"
#     f"Database={DB_NAME};"
#     f"UID={usuario};"
#     f"PWD={contrasena};"
# )
# connection_string = f"mssql+pyodbc:///?odbc_connect={params}"
# engine = create_engine(connection_string, echo=True)



import streamlit as st
import pandas as pd
import pyodbc 
from sqlalchemy import create_engine, text

who_is = "Cris2"
users = {
    "Cris1": ["test_access", "CBARRERA29N7\\SQLEXPRESS"],
    "Cris2": ["test_access", "NB-CD-DEE2"]
}
user = users[who_is]
 
DB_NAME = user[0]
SERVER_NAME = user[1]
usuario = "Fabi.fabi"
contrasena = "Fabifabiño123_-@sjk,chba<siclawr83746rq3efdqwl.kjsh!!!!"
try:
   conn = pyodbc.connect(
    f"Driver={{ODBC Driver 17 for SQL Server}};"
    f"Server={SERVER_NAME};"
    f"Database={DB_NAME};"
    f"UID={usuario};"
    #f"PWD={contrasena};"
    )
except Exception as e:
        print(f"Error al ejecutar query: {e}")
        # st.error(f"❌ Error al ejecutar query: {e}")
        #return pd.DataFrame()   