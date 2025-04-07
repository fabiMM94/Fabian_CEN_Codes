import streamlit as st
import file_manager
from Results import ResultsTable

# Configurar el diseño de la página
st.set_page_config(page_title="EMTP Results Reporter", layout="wide")

# Cargar el archivo desde la PC del usuario
uploaded_file = st.file_uploader("Selecciona un archivo ECF:", type=["html"])

if uploaded_file:
    # Guardamos en la sesión para mantener el estado
    st.session_state.html_full_file_name = uploaded_file.name  

    st.success(f"Archivo seleccionado: {uploaded_file.name}")

# Estilos personalizados en CSS
st.markdown("""
    <style>
        /* Fondo de la sección del título */
        .title-section {
            background-color:  #4a4a4a;
            padding: 20px;
            border-radius: 10px;
            text-align: center;
            color: white;
        }
        .main-title {
            font-size: 36px;
            font-weight: bold;
        }
        .subtitle {
            font-size: 24px;
        }
        /* Estilo de los botones */
        .stButton>button {
            background-color: #1f4e79;
            color: white;
            border-radius: 8px;
            font-size: 16px;
            padding: 10px 20px;
            width: 100%;
        }
        .stButton>button:hover {
            background-color: #4c709d;
        }
    </style>
""", unsafe_allow_html=True)

# Instancia de la clase Gui
class Gui:
    def __init__(self):
        self.fm = file_manager.FileManager()
    
    def write_file(self):
        selected_file = self.fm.promt_select_ecf_file()
        if selected_file:
            st.session_state.html_full_file_name = selected_file
            st.success(f"Archivo seleccionado: {selected_file}")
        else:
            st.warning("No se seleccionó ningún archivo.")
        return selected_file

    def show_html_results(self):
        if st.session_state.html_full_file_name:
            results = ResultsTable(st.session_state.html_full_file_name, None)
            return results.get_report()
        else:
            st.error("El archivo HTML no ha sido seleccionado.")
            return None
    
    def show_BusData_results(self):
        if st.session_state.html_full_file_name:
            results = ResultsTable(st.session_state.html_full_file_name, None)
            return results.get_BusData()
        else:
            st.error("El archivo HTML no ha sido seleccionado.")
            return None
        
    def show_GenData_results(self):
        if st.session_state.html_full_file_name:
            results = ResultsTable(st.session_state.html_full_file_name, None)
            return results.get_GenData()
        else:
            st.error("El archivo HTML no ha sido seleccionado.")
            return None

    def show_LoadData_results(self):
        if st.session_state.html_full_file_name:
            results = ResultsTable(st.session_state.html_full_file_name, None)
            return results.get_LoadData()
        else:
            st.error("El archivo HTML no ha sido seleccionado.")
            return None

# Crear la instancia de GUI
gui_instance = Gui()

# Sección de título con fondo
st.markdown("""
    <div class="title-section">
        <h1 class="main-title">EMTP Results Reporter</h1>
        <h2 class="subtitle">CEN - DSLTR (LAB)</h2>
    </div>
""", unsafe_allow_html=True)

st.divider()

# Diseño en dos columnas
col1, col2 = st.columns([2, 1])

# Sección izquierda (Descripción)
with col1:
    st.markdown("""
        ## About this app:
        - **HTML output file analysis**
        - **ECF file exploration**
        - **TBD** (busbars, tlines, switches, synch. gens, etc.)
    """)
    
# Sección derecha (Botones de acción)
with col2:
    st.write("### Opciones:")

    if st.button("Select ECF file"):
        gui_instance.write_file()
    
    if st.button("Show Report"):
        df = gui_instance.show_html_results()
        if df is not None:
            st.dataframe(df, width=1000, height=600)

# Línea divisoria
st.divider()

# Sección de datos organizados en tres columnas
col3, col4, col5 = st.columns(3)

with col3:
    if st.button("Get Bus Voltages"):
        df = gui_instance.show_BusData_results()
        if df is not None:
            st.dataframe(df, width=1200, height=800)

with col4:
    if st.button("Get Generation Data"):
        df = gui_instance.show_GenData_results()
        if df is not None:
            st.dataframe(df, width=1200, height=800)

with col5:
    if st.button("Get Load Data"):
        df = gui_instance.show_LoadData_results()
        if df is not None:
            st.dataframe(df, width=1200, height=800)

# Línea divisoria
st.divider()

# Otras opciones
if st.button("Más opciones"):
    st.info("Funciones en desarrollo... 🚀")