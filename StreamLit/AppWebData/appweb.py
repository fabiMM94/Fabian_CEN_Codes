import streamlit as st
import json
import os
import sqlite3
import pandas as pd

# 📌 Archivo donde se guardarán los usuarios
USUARIOS_FILE = "usuarios.json"
DB_FILE = "database.db"

# 📌 Función para cargar usuarios desde JSON
def cargar_usuarios():
    if os.path.exists(USUARIOS_FILE):
        with open(USUARIOS_FILE, "r") as f:
            return json.load(f)
    return {"admin": "1234", "usuario": "abcd"}  # Usuarios predeterminados

# 📌 Función para guardar nuevos usuarios en JSON
def guardar_usuarios(usuarios):
    with open(USUARIOS_FILE, "w") as f:
        json.dump(usuarios, f)

# 📌 Función para obtener datos de la base de datos
def obtener_datos(query="SELECT * FROM empleados"):
    conn = sqlite3.connect(DB_FILE)
    df = pd.read_sql(query, conn)
    conn.close()
    return df

# 📌 Inicializar usuarios en la sesión si no existen
if "usuarios" not in st.session_state:
    st.session_state.usuarios = cargar_usuarios()

# 📌 Inicializar sesión
if "autenticado" not in st.session_state:
    st.session_state.autenticado = False

if "registro" not in st.session_state:
    st.session_state.registro = False

st.title("Aplicacion Web de base de datos")

# 🟢 Si el usuario NO está autenticado, mostrar login o registro
if not st.session_state.autenticado:
    if not st.session_state.registro:
        usuario = st.text_input("Nombre de Usuario")
        contrasena = st.text_input("Contraseña", type="password")

        if st.button("Iniciar Sesión"):
            if st.session_state.usuarios.get(usuario) == contrasena:
                st.session_state.autenticado = True
                st.success("¡Inicio de sesión exitoso!")
                st.rerun()
            else:
                st.error("Usuario o contraseña incorrectos")

        if st.button("Crear Usuario"):
            st.session_state.registro = True
            st.rerun()
    
    else:
        st.subheader("Registro de Nuevo Usuario")
        nuevo_usuario = st.text_input("Nuevo Nombre de Usuario")
        nueva_contrasena = st.text_input("Nueva Contraseña", type="password")

        if st.button("Registrar"):
            if nuevo_usuario in st.session_state.usuarios:
                st.error("El usuario ya existe. Elige otro nombre.")
            elif nuevo_usuario and nueva_contrasena:
                st.session_state.usuarios[nuevo_usuario] = nueva_contrasena
                guardar_usuarios(st.session_state.usuarios)  # Guardar en el JSON
                st.success(f"Usuario '{nuevo_usuario}' creado con éxito. ¡Inicia sesión!")
                st.session_state.registro = False
                st.rerun()
            else:
                st.error("Por favor, completa todos los campos.")

        if st.button("Cancelar"):
            st.session_state.registro = False
            st.rerun()

# 🔵 Si el usuario está autenticado, mostrar la base de datos
if st.session_state.autenticado:
    st.subheader("Bienvenido, consulta la base de datos")

    # Botón para mostrar todos los empleados
    if st.button("Mostrar todos los empleados"):
        df = obtener_datos()  # Consultar todos los empleados
        st.dataframe(df)  # Mostrar los datos en un DataFrame en la interfaz

    # Búsqueda por nombre
    nombre_busqueda = st.text_input("Buscar empleado por nombre:")
    if st.button("Buscar"):
        query = f"SELECT * FROM empleados WHERE nombre LIKE '%{nombre_busqueda}%'"
        df = obtener_datos(query)  # Llamada a la función con la consulta SQL dinámica
        st.dataframe(df)  # Mostrar los resultados filtrados

    # Filtrar empleados por puesto
    puesto_seleccionado = st.selectbox("Filtrar por puesto:", ["Todos", "Ingeniera", "Analista", "Gerente"])
    if puesto_seleccionado != "Todos":
        query = f"SELECT * FROM empleados WHERE puesto = '{puesto_seleccionado}'"
        df = obtener_datos(query)  # Llamada a la función con la consulta filtrada por puesto
        st.dataframe(df)  # Mostrar los resultados filtrados por puesto

    # Botón para cerrar sesión
    if st.button("Cerrar Sesión"):
        st.session_state.autenticado = False
        st.rerun()
