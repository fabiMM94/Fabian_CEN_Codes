import sqlite3
import os

# Obtener la ruta del directorio donde se está ejecutando el script
directorio_actual = os.path.dirname(os.path.abspath(__file__))

# Ruta completa para el archivo .db en la misma carpeta
db_path = os.path.join(directorio_actual, 'database.db')

# Conectar a la base de datos (se crea si no existe)
conn = sqlite3.connect(db_path)
cursor = conn.cursor()

# Crear la tabla 'empleados' si no existe
cursor.execute("""
CREATE TABLE IF NOT EXISTS empleados (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    nombre TEXT,
    puesto TEXT,
    salario REAL
)
""")
print("Tabla 'empleados' creada o ya existe.")

# Verificar si la tabla tiene datos
cursor.execute("SELECT COUNT(*) FROM empleados")
num_empleados = cursor.fetchone()[0]
print(f"Empleados en la tabla: {num_empleados}")

if num_empleados == 0:
    # Insertar datos de prueba solo si la tabla está vacía
    empleados_prueba = [
        ("Ana Pérez", "Ingeniera", 50000),
        ("Carlos Gómez", "Analista", 45000),
        ("Marta López", "Gerente", 60000)
    ]
    cursor.executemany("INSERT INTO empleados (nombre, puesto, salario) VALUES (?, ?, ?)", empleados_prueba)
    print("Datos de prueba insertados.")
else:
    print("La tabla 'empleados' ya tiene datos.")

# Guardar los cambios y cerrar la conexión
conn.commit()
conn.close()

print(f"Base de datos guardada en: {db_path}")