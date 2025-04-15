import streamlit as st
import pandas as pd
import pyodbc

# Configurar el diseño de la página (debe ser el primer comando de Streamlit)
st.set_page_config(
    page_title="Mi App",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Título
st.title("Consulta a SQL Server")

# Variable global para la conexión
conn = None

# Función para conectarse a la base de datos
def get_connection():
    global conn
    if conn is None:  # Si no hay una conexión activa, crear una nueva
        try:
            conn = pyodbc.connect(
                'DRIVER={ODBC Driver 17 for SQL Server};'
                'SERVER=NB-CD-DEE;'
                'DATABASE=red_prueba (1);'
                'Trusted_Connection=yes;'
            )
        except Exception as e:
            st.error(f"Error de conexión: {e}")
            return None
    return conn

# Función para cerrar la conexión
def close_connection():
    global conn
    if conn:
        conn.close()
        conn = None

# Consulta de ejemplo
def consultar_datos():
    conn = get_connection()
    if conn:
        query = "SELECT * FROM dbo.PowerTransformerEnd"
        df = pd.read_sql(query, conn)
        return df
    return pd.DataFrame()  # Retorna un DataFrame vacío si no hay conexión

def consultar_nombre_y_resistencia():
    conn = get_connection()
    if conn:
        query = """
            SELECT 
                PowerTransformerEnd.PowerTransformerEndID,
                Resistance.value
            FROM 
                dbo.PowerTransformerEnd
            JOIN 
                dbo.Resistance
            ON 
                PowerTransformerEnd.rID = Resistance.ResistanceID
        """
        df = pd.read_sql(query, conn)
        return df
    return pd.DataFrame()  # Retorna un DataFrame vacío si no hay conexión

def consultar_nombre_resistencia_y_reactancia():
    conn = get_connection()
    if conn:
        query = """
            SELECT 
                PowerTransformerEnd.PowerTransformerEndID,
                Resistance.value AS resistance_value,
                Reactance.value AS reactance_value
            FROM 
                dbo.PowerTransformerEnd
            JOIN 
                dbo.Resistance ON PowerTransformerEnd.rID = Resistance.ResistanceID
            JOIN 
                dbo.Reactance ON PowerTransformerEnd.xID = Reactance.ReactanceID
            JOIN     
                dbo.Reactance ON PowerTransformerEnd.ratedSID = ApparentPower.ApparentPowerID 
        """
        df = pd.read_sql(query, conn)
        return df
    return pd.DataFrame()  # Retorna un DataFrame vacío si no hay conexión
def consultar_power_transformer_end():
    conn = get_connection()
    if conn:
        query = """
            SELECT 
                PowerTransformerEnd.PowerTransformerEndID,
                Resistance.value AS resistance_value,
                Reactance.value AS reactance_value,
                ApparentPower.value AS apparent_power_value
            FROM 
                dbo.PowerTransformerEnd
            JOIN 
                dbo.Resistance ON PowerTransformerEnd.rID = Resistance.ResistanceID
            JOIN 
                dbo.Reactance ON PowerTransformerEnd.xID = Reactance.ReactanceID
            JOIN 
                dbo.ApparentPower ON PowerTransformerEnd.ratedSID = ApparentPower.ApparentPowerID
        """
        df = pd.read_sql(query, conn)
        return df
    return pd.DataFrame() 
#------------------------------------------------ Estilo global para todas las tablas



#------------------------------------------------ Botones     
# Interfaz de Streamlit
if st.button("Consultar datos"):
    datos = consultar_datos()

    # Mostrar la tabla
    st.dataframe(datos, height=800, width=None)

if st.button("Ver PowerTransformerEndID y valor de resistencia"):
    datos = consultar_nombre_y_resistencia()
    st.dataframe(datos, height=600, width=None)

if st.button("Ver PowerTransformerEndID, resistencia y reactancia"):
    datos = consultar_nombre_resistencia_y_reactancia()
    st.dataframe(datos, height=600, width=None)


if st.button("Ver PowerTransformerEndID, resistencia, reactancia y potencia aparente"):
    datos = consultar_power_transformer_end()
    st.dataframe(datos, height=600, width=None)
# Al final, cerramos la conexión para liberar recursos
close_connection()