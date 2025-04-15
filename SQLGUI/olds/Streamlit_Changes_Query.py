import streamlit as st
import pandas as pd
from sqlalchemy import create_engine, MetaData, Table, select, update
import pyodbc

# Configurar conexión a SQL Server
server = 'NB-CD-DEE'
database = 'red_prueba (1)'  # Nombre real de la base de datos
connection_string = (
    f"mssql+pyodbc://@{server}/{database}?driver=ODBC+Driver+17+for+SQL+Server&trusted_connection=yes"
)
engine = create_engine(connection_string)
metadata = MetaData(schema="dbo")

# Cargar las tablas
power_transformer_end = Table('PowerTransformerEnd', metadata, autoload_with=engine)
resistance = Table('Resistance', metadata, autoload_with=engine)
reactance = Table('Reactance', metadata, autoload_with=engine)
apparent_power = Table('ApparentPower', metadata, autoload_with=engine)

# Nueva tabla creada a partir de la consulta
def load_new_table():
    with engine.connect() as conn:
        query = select([
            power_transformer_end.c.PowerTransformerEndID,
            resistance.c.value.label('resistance_value'),
            reactance.c.value.label('reactance_value'),
            apparent_power.c.value.label('apparent_power_value')
        ]).join(resistance, power_transformer_end.c.riD == resistance.c.ResistanceID) \
          .join(reactance, power_transformer_end.c.xID == reactance.c.ReactanceID) \
          .join(apparent_power, power_transformer_end.c.ratedSID == apparent_power.c.ApparentPowerID)

        result = conn.execute(query)
        rows = result.fetchall()  # Obtener las filas
        # Crear el DataFrame
        df = pd.DataFrame(rows, columns=[
            'PowerTransformerEndID', 'resistance_value', 'reactance_value', 'apparent_power_value'
        ])
    return df

# Actualizar las tablas de origen
def update_database(original_df, edited_df):
    changes = edited_df.compare(original_df)
    if not changes.empty:
        with engine.begin() as conn:
            for idx in changes.index.get_level_values(0).unique():
                row = edited_df.loc[idx]
                stmt_resistance = (
                    update(resistance)
                    .where(resistance.c.ResistanceID == row["riD"])
                    .values(value=row["resistance_value"])
                )
                stmt_reactance = (
                    update(reactance)
                    .where(reactance.c.ReactanceID == row["xID"])
                    .values(value=row["reactance_value"])
                )
                stmt_apparent_power = (
                    update(apparent_power)
                    .where(apparent_power.c.ApparentPowerID == row["ratedSID"])
                    .values(value=row["apparent_power_value"])
                )

                conn.execute(stmt_resistance)
                conn.execute(stmt_reactance)
                conn.execute(stmt_apparent_power)

        st.success("✅ Cambios guardados en las tablas de origen.")
    else:
        st.info("ℹ️ No hay cambios para guardar.")

# Interfaz Streamlit
st.title("🛠️ Editor de Datos de Transformador de Potencia")

# Cargar la nueva tabla combinada
original_df = load_new_table()

# Mostrar y permitir editar
edited_df = st.data_editor(
    original_df[["PowerTransformerEndID", "resistance_value", "reactance_value", "apparent_power_value"]],
    use_container_width=True,
    num_rows="dynamic",
    key="editable_table"
)

# Botón para guardar los cambios
if st.button("💾 Guardar cambios en las tablas de origen"):
    update_database(original_df, edited_df)
