# usuarios.py
import json
import os

USUARIOS_FILE = "usuarios.json"

class Usuario:
    def __init__(self):
        self.usuarios = self.cargar_usuarios()

    # Función para cargar usuarios desde JSON
    def cargar_usuarios(self):
        if os.path.exists(USUARIOS_FILE):
            with open(USUARIOS_FILE, "r") as f:
                return json.load(f)
        return {"admin": "1234", "usuario": "abcd"}  # Usuarios predeterminados

    # Función para guardar nuevos usuarios en JSON
    def guardar_usuarios(self):
        with open(USUARIOS_FILE, "w") as f:
            json.dump(self.usuarios, f)

    # Función para verificar las credenciales del usuario
    def verificar_usuario(self, usuario, contrasena):
        return self.usuarios.get(usuario) == contrasena

    # Función para registrar un nuevo usuario
    def registrar_usuario(self, nuevo_usuario, nueva_contrasena):
        self.usuarios[nuevo_usuario] = nueva_contrasena
        self.guardar_usuarios()