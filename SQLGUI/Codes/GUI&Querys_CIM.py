import streamlit as st
import pandas as pd
import sqlalchemy as sa
import pyodbc
from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy.orm import sessionmaker, relationship
from sqlalchemy import create_engine, Column, Integer, Float, ForeignKey, text, Table, MetaData


# Configuración de página de Streamlit
st.set_page_config(page_title="Editor de Datos SQL Server", layout="wide")
st.title("Editor de Datos SQL Server")

# Configuración de conexión a SQL Server
SERVER  = 'NB-CD-DEE'
DATABASE = 'red_prueba (1)' 
#SERVER = 'FMEDINA98JL\\SQLEXPRESS'
#DATABASE = 'red_prueba'
DRIVER = 'ODBC Driver 17 for SQL Server'  # Ajusta según el driver que tengas instalado

# Función para crear la conexión a la base de datos
def get_connection_string():
    # Para autenticación de Windows
    connection_string = f"mssql+pyodbc://{SERVER}/{DATABASE}?driver={DRIVER}&trusted_connection=yes"
    # Si necesitas autenticación con usuario y contraseña, usa:
    # connection_string = f"mssql+pyodbc://{username}:{password}@{SERVER}/{DATABASE}?driver={DRIVER}"
    return connection_string

# Crear engine de SQLAlchemy
try:
    engine = create_engine(get_connection_string())
    st.sidebar.success("Conexión exitosa a SQL Server")
except Exception as e:
    st.sidebar.error(f"Error de conexión: {e}")
    st.stop()

# Función para obtener datos combinados
def get_combined_data():
    query = """
    SELECT 
        pte.PowerTransformerEndID, 
        r.value AS resistance_value, 
        x.value AS reactance_value, 
        ap.value AS apparent_power_value,
        r.ResistanceID,
        x.ReactanceID,
        ap.ApparentPowerID
    FROM dbo.PowerTransformerEnd pte
    LEFT JOIN dbo.Resistance r ON pte.riD = r.ResistanceID
    LEFT JOIN dbo.Reactance x ON pte.xID = x.ReactanceID
    LEFT JOIN dbo.ApparentPower ap ON pte.ratedSID = ap.ApparentPowerID
    """
    
    try:
        return pd.read_sql(query, engine)
    except Exception as e:
        st.error(f"Error al consultar datos: {e}")
        st.error(f"Query: {query}")
        return pd.DataFrame()

# Función para insertar nuevos valores y actualizar referencias
def update_with_new_records(transformer_id, resistance_value, reactance_value, apparent_power_value, r_id, x_id, rated_s_id):
    try:
        with engine.begin() as conn:
            changes_made = False
            new_values = {}
            
            # Para cada valor editado, crear un nuevo registro
            
            # 1. Resistencia
            if r_id is not None:
                # Obtener el máximo ID actual
                max_r_id = conn.execute(text("SELECT MAX(ResistanceID) FROM dbo.Resistance")).scalar()
                new_r_id = max_r_id + 1 if max_r_id else 1
                
                # Activar IDENTITY_INSERT antes de la inserción
                conn.execute(text("SET IDENTITY_INSERT dbo.Resistance ON"))
                try:
                    # Insertar nuevo registro siempre
                    conn.execute(
                        text("INSERT INTO dbo.Resistance (ResistanceID, value) VALUES (:id, :value)"),
                        {"id": new_r_id, "value": resistance_value}
                    )
                    new_values['r_id'] = new_r_id
                    changes_made = True
                finally:
                    # Desactivar IDENTITY_INSERT después de la inserción
                    conn.execute(text("SET IDENTITY_INSERT dbo.Resistance OFF"))
            
            # 2. Reactancia
            if x_id is not None:
                # Obtener el máximo ID actual
                max_x_id = conn.execute(text("SELECT MAX(ReactanceID) FROM dbo.Reactance")).scalar()
                new_x_id = max_x_id + 1 if max_x_id else 1
                
                # Activar IDENTITY_INSERT antes de la inserción
                conn.execute(text("SET IDENTITY_INSERT dbo.Reactance ON"))
                try:
                    # Insertar nuevo registro siempre
                    conn.execute(
                        text("INSERT INTO dbo.Reactance (ReactanceID, value) VALUES (:id, :value)"),
                        {"id": new_x_id, "value": reactance_value}
                    )
                    new_values['x_id'] = new_x_id
                    changes_made = True
                finally:
                    # Desactivar IDENTITY_INSERT después de la inserción
                    conn.execute(text("SET IDENTITY_INSERT dbo.Reactance OFF"))
            
            # 3. Potencia Aparente
            if rated_s_id is not None:
                # Obtener el máximo ID actual
                max_s_id = conn.execute(text("SELECT MAX(ApparentPowerID) FROM dbo.ApparentPower")).scalar()
                new_s_id = max_s_id + 1 if max_s_id else 1
                
                # Activar IDENTITY_INSERT antes de la inserción
                conn.execute(text("SET IDENTITY_INSERT dbo.ApparentPower ON"))
                try:
                    # Insertar nuevo registro siempre
                    conn.execute(
                        text("INSERT INTO dbo.ApparentPower (ApparentPowerID, value) VALUES (:id, :value)"),
                        {"id": new_s_id, "value": apparent_power_value}
                    )
                    new_values['s_id'] = new_s_id
                    changes_made = True
                finally:
                    # Desactivar IDENTITY_INSERT después de la inserción
                    conn.execute(text("SET IDENTITY_INSERT dbo.ApparentPower OFF"))
                    
            # Actualizar referencias en PowerTransformerEnd
            if changes_made:
                update_fields = []
                update_values = {"id": transformer_id}
                
                if 'r_id' in new_values:
                    update_fields.append("riD = :r_id")
                    update_values["r_id"] = new_values['r_id']
                    
                if 'x_id' in new_values:
                    update_fields.append("xID = :x_id")
                    update_values["x_id"] = new_values['x_id']
                    
                if 's_id' in new_values:
                    update_fields.append("ratedSID = :s_id")
                    update_values["s_id"] = new_values['s_id']
                    
                if update_fields:
                    update_query = f"UPDATE dbo.PowerTransformerEnd SET {', '.join(update_fields)} WHERE PowerTransformerEndID = :id"
                    conn.execute(text(update_query), update_values)
                    
            return True, changes_made, new_values
    except Exception as e:
        st.error(f"Error al actualizar datos: {e}")
        return False, False, {}

# Función para verificar si las tablas existen
def check_tables_exist():
    try:
        # Verificar existencia de tablas
        tables_query = """
        SELECT TABLE_NAME 
        FROM INFORMATION_SCHEMA.TABLES 
        WHERE TABLE_SCHEMA = 'dbo' 
        AND TABLE_NAME IN ('PowerTransformerEnd', 'Resistance', 'Reactance', 'ApparentPower')
        """
        
        tables_df = pd.read_sql(tables_query, engine)
        existing_tables = tables_df['TABLE_NAME'].tolist()
        
        if len(existing_tables) < 4:
            missing_tables = set(['PowerTransformerEnd', 'Resistance', 'Reactance', 'ApparentPower']) - set(existing_tables)
            st.warning(f"Algunas tablas no existen en la base de datos: {', '.join(missing_tables)}")
            return False
        return True
    except Exception as e:
        st.error(f"Error al verificar tablas: {e}")
        return False

# Sidebar para opciones
st.sidebar.title("Opciones")

# Verificar tablas
if st.sidebar.button("Verificar Tablas"):
    if check_tables_exist():
        st.sidebar.success("Todas las tablas necesarias existen")
    else:
        st.sidebar.error("Faltan algunas tablas en la base de datos")

# Mostrar datos
if st.sidebar.button("Cargar Datos"):
    df = get_combined_data()
    
    if not df.empty:
        st.session_state.data = df
        st.success("Datos cargados correctamente")
    else:
        st.warning("No se encontraron datos o hubo un error en la consulta")

# Editor de datos
if 'data' in st.session_state:
    st.subheader("Datos Combinados")
    
    # Preparar DataFrame para edición
    df_display = st.session_state.data.copy()
    editable_df = df_display[['PowerTransformerEndID', 'resistance_value', 'reactance_value', 'apparent_power_value']].copy()
    
    # Guardar el DataFrame original en session_state para comparar después
    if 'original_df' not in st.session_state:
        st.session_state.original_df = editable_df.copy()
    
    # Usar el editor de datos de Streamlit
    edited_df = st.data_editor(
        editable_df, 
        num_rows="fixed", 
        key="data_editor",
        column_config={
            "PowerTransformerEndID": st.column_config.NumberColumn("Transformer ID", disabled=True),
            "resistance_value": st.column_config.NumberColumn("Resistencia", min_value=0.0, format="%.5f"),
            "reactance_value": st.column_config.NumberColumn("Reactancia", min_value=0.0, format="%.5f"),
            "apparent_power_value": st.column_config.NumberColumn("Potencia Aparente", min_value=0.0, format="%.2f")
        }
    )
    
    # Columna para mostrar los cambios
    col1, col2 = st.columns([3, 1])
    
    # Verificar cambios
    if col1.button("Guardar Cambios"):
        success_count = 0
        error_count = 0
        changes_count = 0
        new_ids_summary = []
        
        # Progreso visual
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        # Comparar con datos originales para detectar cambios
        original_df = st.session_state.data
        total_rows = len(edited_df)
        
        for i, (index, row) in enumerate(edited_df.iterrows()):
            if index >= len(original_df):
                st.warning(f"Índice {index} fuera de rango. Saltando.")
                continue
                
            original_row = original_df.iloc[index]
            
            # Detectar si hay cambios (cualquier valor diferente)
            has_changes = (
                abs(row['resistance_value'] - original_row['resistance_value']) > 1e-6 or
                abs(row['reactance_value'] - original_row['reactance_value']) > 1e-6 or
                abs(row['apparent_power_value'] - original_row['apparent_power_value']) > 1e-6
            )
            
            # Si hay cambios, actualizar y crear nuevos registros
            if has_changes:
                result, changes_made, new_ids = update_with_new_records(
                    original_row['PowerTransformerEndID'],
                    row['resistance_value'],
                    row['reactance_value'],
                    row['apparent_power_value'],
                    original_row['ResistanceID'],
                    original_row['ReactanceID'],
                    original_row['ApparentPowerID']
                )
                
                if result:
                    success_count += 1
                    if changes_made:
                        changes_count += 1
                        # Guardar resumen de cambios para mostrar
                        summary = f"Transformador #{original_row['PowerTransformerEndID']}:"
                        if 'r_id' in new_ids:
                            summary += f" Nuevo ID Resistencia: {new_ids['r_id']}"
                        if 'x_id' in new_ids:
                            summary += f" Nuevo ID Reactancia: {new_ids['x_id']}"
                        if 's_id' in new_ids:
                            summary += f" Nuevo ID Potencia: {new_ids['s_id']}"
                        new_ids_summary.append(summary)
                else:
                    error_count += 1
            
            # Actualizar barra de progreso
            progress_percent = (i + 1) / total_rows
            progress_bar.progress(progress_percent)
            status_text.text(f"Procesando {i+1}/{total_rows} registros...")
        
        # Actualizar datos en sesión si hubo actualizaciones exitosas
        if success_count > 0:
            updated_df = get_combined_data()
            if not updated_df.empty:
                st.session_state.data = updated_df
                st.session_state.original_df = updated_df[['PowerTransformerEndID', 'resistance_value', 'reactance_value', 'apparent_power_value']].copy()
            
            if changes_count > 0:
                st.success(f"Se actualizaron {changes_count} registros con nuevos valores.")
                # Mostrar resumen de cambios
                with st.expander("Ver detalles de actualizaciones"):
                    for summary in new_ids_summary:
                        st.text(summary)
            else:
                st.info("No se detectaron cambios para guardar.")
                
            if error_count > 0:
                st.warning(f"No se pudieron actualizar {error_count} registros.")
        elif error_count > 0:
            st.error(f"No se pudieron actualizar {error_count} registros.")
        else:
            st.info("No se detectaron cambios para guardar.")
        
        # Eliminar barra de progreso y texto de estado
        progress_bar.empty()
        status_text.empty()
    
    # Botón para descartar cambios
    if col2.button("Descartar Cambios"):
        if 'original_df' in st.session_state:
            st.experimental_rerun()  # Esto recargará la página
        else:
            st.warning("No hay cambios para descartar")

# Sección para consulta SQL personalizada
st.sidebar.subheader("Consulta SQL Personalizada")
custom_query = st.sidebar.text_area("Ingresa tu consulta SQL:")

if st.sidebar.button("Ejecutar Consulta"):
    if custom_query:
        try:
            custom_df = pd.read_sql(custom_query, engine)
            st.subheader("Resultados de la Consulta Personalizada")
            st.dataframe(custom_df)
            
            # Opción para descargar resultados
            csv = custom_df.to_csv(index=False).encode('utf-8')
            st.download_button(
                "Descargar resultados como CSV",
                csv,
                "query_results.csv",
                "text/csv",
                key='download-csv'
            )
        except Exception as e:
            st.error(f"Error al ejecutar consulta: {e}")
    else:
        st.warning("Ingresa una consulta SQL válida")

# Información adicional
st.sidebar.markdown("---")
st.sidebar.info("""
**Instrucciones:**
1. Haz clic en "Verificar Tablas" para comprobar que existen las tablas necesarias
2. Haz clic en "Cargar Datos" para obtener los datos
3. Edita los valores en la tabla
4. Haz clic en "Guardar Cambios" para actualizar
   
Nota: Cuando editas valores, se crean SIEMPRE nuevos registros en las tablas correspondientes con nuevos IDs,
y se actualizan las referencias en la tabla PowerTransformerEnd.
""")