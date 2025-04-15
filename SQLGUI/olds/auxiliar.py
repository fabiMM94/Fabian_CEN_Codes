import streamlit as st
import pandas as pd
from sqlalchemy import create_engine, text

# Configurar la página
st.set_page_config(page_title="Editor SQL", layout="wide")
st.title("Editor de datos PowerTransformerEnd")

# Crear motor de conexión SQLAlchemy
@st.cache_resource
def get_engine():
    return create_engine(
        'mssql+pyodbc://NB-CD-DEE/red_prueba (1)?driver=ODBC+Driver+17+for+SQL+Server',
        fast_executemany=True
    )

engine = get_engine()

# Consultar tabla final (PowerTransformerEnd + joins)
def consultar_power_transformer_end():
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
    with engine.connect() as conn:
        df = pd.read_sql(text(query), conn)
    return df

# Función para actualizar las 3 tablas pequeñas
def actualizar_tablas(df_editada):
    with engine.begin() as conn:  # inicia transacción
        for _, row in df_editada.iterrows():
            conn.execute(text("""
                UPDATE dbo.Resistance
                SET value = :res_val
                WHERE ResistanceID = (
                    SELECT rID FROM dbo.PowerTransformerEnd WHERE PowerTransformerEndID = :id
                )
            """), {"res_val": row["resistance_value"], "id": row["PowerTransformerEndID"]})

            conn.execute(text("""
                UPDATE dbo.Reactance
                SET value = :react_val
                WHERE ReactanceID = (
                    SELECT xID FROM dbo.PowerTransformerEnd WHERE PowerTransformerEndID = :id
                )
            """), {"react_val": row["reactance_value"], "id": row["PowerTransformerEndID"]})

            conn.execute(text("""
                UPDATE dbo.ApparentPower
                SET value = :s_val
                WHERE ApparentPowerID = (
                    SELECT ratedSID FROM dbo.PowerTransformerEnd WHERE PowerTransformerEndID = :id
                )
            """), {"s_val": row["apparent_power_value"], "id": row["PowerTransformerEndID"]})

# 🧠 Carga los datos solo una vez y almacénalos en session_state
if "df_editada" not in st.session_state:
    st.session_state["df_editada"] = consultar_power_transformer_end()

# Mostrar editor de tabla
st.subheader("Editar valores:")
df_editada = st.data_editor(st.session_state["df_editada"], use_container_width=True, num_rows="dynamic")

# Botón para guardar cambios
if st.button("Guardar cambios en SQL Server"):
    try:
        actualizar_tablas(df_editada)
        st.success("✅ Cambios guardados correctamente en la base de datos.")

        # Volver a consultar los datos y actualizar el session_state
        st.session_state["df_editada"] = consultar_power_transformer_end()
    except Exception as e:
        st.error(f"❌ Error al guardar los cambios: {e}")

# Mostrar tabla actual
st.subheader("Vista actual de la tabla")
st.dataframe(st.session_state["df_editada"], use_container_width=True)


