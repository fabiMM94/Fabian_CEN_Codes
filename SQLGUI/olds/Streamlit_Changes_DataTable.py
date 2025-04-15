import streamlit as st
import pandas as pd
from sqlalchemy import create_engine, MetaData, Table, select, update
import pyodbc

# Configurar conexión a SQL Server
server = 'NB-CD-DEE'
database = 'red_prueba (1)'  # <-- Reemplaza con el nombre real
connection_string = (
    f"mssql+pyodbc://@{server}/{database}?driver=ODBC+Driver+17+for+SQL+Server&trusted_connection=yes"
)
engine = create_engine(connection_string)
metadata = MetaData(schema="dbo")

# Cargar la tabla
nueva_tabla = Table('NuevaTabla2', metadata, autoload_with=engine)

# Leer datos como DataFrame
def load_data():
    with engine.connect() as conn:
        result = conn.execute(select(nueva_tabla))
        df = pd.DataFrame(result.mappings().all())
    return df

# Actualizar datos modificados
def update_database(original_df, edited_df):
    # Asegúrate de que ambos DataFrames tengan el mismo índice y columnas
    original_df = original_df.sort_index(axis=1).reset_index(drop=True)
    edited_df = edited_df.sort_index(axis=1).reset_index(drop=True)
    
    # Comparar los DataFrames
    changes = edited_df.compare(original_df)
    
    if not changes.empty:
        with engine.begin() as conn:
            for idx in changes.index.get_level_values(0).unique():
                row = edited_df.loc[idx]
                
                # Convertir el id y otros parámetros a tipos compatibles
                id_value = int(row["id"])  # Convertir el id a un tipo nativo de Python
                nombre_value = str(row["nombre"])  # Convertir el nombre a string
                
                stmt = (
                    update(nueva_tabla)
                    .where(nueva_tabla.c.id == id_value)
                    .values(nombre=nombre_value)
                )
                conn.execute(stmt)
        st.success("✅ Cambios guardados en la base de datos.")
    else:
        st.info("ℹ️ No hay cambios para guardar.")

# Interfaz Streamlit
st.title("🛠️ Editor de Tabla SQL - NuevaTabla2")

# Cargar datos originales
original_df = load_data()

# Asegúrate de que las columnas que se editan sean las mismas
original_df = original_df[["id", "nombre"]]

# Mostrar y permitir editar
edited_df = st.data_editor(
    original_df,
    use_container_width=True,
    num_rows="dynamic",
    key="editable_table"
)

# Botón para guardar
if st.button("💾 Guardar cambios en SQL Server"):
    update_database(original_df, edited_df)
