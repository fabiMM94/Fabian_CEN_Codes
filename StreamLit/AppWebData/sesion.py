# sesion.py
import streamlit as st
from usuarios import Usuario

class Sesion:
    def __init__(self):
        self.usuario = Usuario()
        # Inicializa el estado de la sesión en Streamlit
        if "autenticado" not in st.session_state:
            st.session_state.autenticado = False

        if "registro" not in st.session_state:
            st.session_state.registro = False

    # Función para iniciar sesión
    def iniciar_sesion(self, usuario_input, contrasena_input):
        if self.usuario.verificar_usuario(usuario_input, contrasena_input):
            st.session_state.autenticado = True
            st.success("¡Inicio de sesión exitoso!")
            st.rerun()
        else:
            st.error("Usuario o contraseña incorrectos")

    # Función para registrar un nuevo usuario
    def registrar_usuario(self, nuevo_usuario, nueva_contrasena):
        self.usuario.registrar_usuario(nuevo_usuario, nueva_contrasena)
        st.success(f"Usuario '{nuevo_usuario}' creado con éxito. ¡Inicia sesión!")
        st.session_state.registro = False
        st.rerun()

    # Función para cancelar el registro
    def cancelar_registro(self):
        st.session_state.registro = False
        st.rerun()