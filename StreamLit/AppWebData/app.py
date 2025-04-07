# app.py
import streamlit as st
from sesion import Sesion
from db import BaseDeDatos

# 📌 Inicializar las clases
sesion = Sesion()
db = BaseDeDatos()

# 📌 Título de la aplicación
st.title("Aplicación Web de Base de Datos")

# 🟢 Si el usuario NO está autenticado, mostrar login o registro
if not st.session_state.autenticado:
    if not st.session_state.registro:
        usuario_input = st.text_input("Nombre de Usuario")
        contrasena_input = st.text_input("Contraseña", type="password")

        if st.button("Iniciar Sesión"):
            sesion.iniciar_sesion(usuario_input, contrasena_input)

        if st.button("Crear Usuario"):
            st.session_state.registro = True
            st.rerun()
    
    else:
        st.subheader("Registro de Nuevo Usuario")
        nuevo_usuario = st.text_input("Nuevo Nombre de Usuario")
        nueva_contrasena = st.text_input("Nueva Contraseña", type="password")

        if st.button("Registrar"):
            if nuevo_usuario in sesion.usuario.usuarios:
                st.error("El usuario ya existe. Elige otro nombre.")
            elif nuevo_usuario and nueva_contrasena:
                sesion.registrar_usuario(nuevo_usuario, nueva_contrasena)
            else:
                st.error("Por favor, completa todos los campos.")

        if st.button("Cancelar"):
            sesion.cancelar_registro()

# 🔵 Si el usuario está autenticado, mostrar la base de datos
if st.session_state.autenticado:
    st.subheader("Bienvenido, consulta la base de datos")

    # Botón para mostrar todos los empleados
    if st.button("Mostrar todos los empleados"):
        df = db.obtener_datos()  # Consultar todos los empleados
        st.dataframe(df)  # Mostrar los datos en un DataFrame en la interfaz

    # Búsqueda por nombre
    nombre_busqueda = st.text_input("Buscar empleado por nombre:")
    if st.button("Buscar"):
        query = f"SELECT * FROM empleados WHERE nombre LIKE '%{nombre_busqueda}%'"
        df = db.obtener_datos(query)  # Llamada a la función con la consulta SQL dinámica
        st.dataframe(df)  # Mostrar los resultados filtrados

    # Filtrar empleados por puesto
    puesto_seleccionado = st.selectbox("Filtrar por puesto:", ["Todos", "Ingeniera", "Analista", "Gerente"])
    if puesto_seleccionado != "Todos":
        query = f"SELECT * FROM empleados WHERE puesto = '{puesto_seleccionado}'"
        df = db.obtener_datos(query)  # Llamada a la función con la consulta filtrada por puesto
        st.dataframe(df)  # Mostrar los resultados filtrados por puesto

    # Botón para cerrar sesión
    if st.button("Cerrar Sesión"):
        st.session_state.autenticado = False
        st.rerun()