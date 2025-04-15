#----------------- intento 1  con pyodbc : Si funciona
# import pyodbc

# # Configuración de conexión
# server = 'NB-CD-DEE'
# database = 'red_prueba (1)'  # <- reemplaza con el nombre real de tu base
# conn_str = (
#     f'DRIVER={{ODBC Driver 17 for SQL Server}};'
#     f'SERVER={server};'
#     f'DATABASE={database};'
#     f'Trusted_Connection=yes;'
# )

# # Conectarse y crear la tabla
# try:
#     conn = pyodbc.connect(conn_str)
#     cursor = conn.cursor()

#     # Crear tabla red_prueba
#     cursor.execute("""
#         CREATE TABLE dbo.NuevaTabla (
#             id INT PRIMARY KEY,
#             nombre VARCHAR(100),
#             fecha_creacion DATETIME DEFAULT GETDATE()
#         )
#     """)
#     conn.commit()

#     print("Tabla 'red_prueba' creada exitosamente.")
# except Exception as e:
#     print("Error al conectarse o crear la tabla:", e)
# finally:
#     cursor.close()
#     conn.close()


# --------------------------------------------------------------
#        Resumen 
#        Este codigo crea una nueva tabla en la base de datos y agrega elementos a a esa tabla 
#        Si la tabla esta creada sobre escribe filas a esa tabla
#------------------------------------------------------------------------------------------------


#------------------------------ Intento 2 con Alchemy 
from sqlalchemy import create_engine, MetaData, Table, Column, Integer, String, DateTime, insert, select
from sqlalchemy.sql import func

# Configuración de conexión
server = 'NB-CD-DEE'
database = 'red_prueba (1)'  # <-- Reemplaza con tu base de datos real
connection_string = (
    f"mssql+pyodbc://@{server}/{database}?driver=ODBC+Driver+17+for+SQL+Server&trusted_connection=yes"
)

# Crear engine
engine = create_engine(connection_string)

# Definir metadata y tabla con ID autoincremental
metadata = MetaData(schema="dbo")

nueva_tabla = Table(
    'NuevaTabla2', metadata,
    Column('id', Integer, primary_key=True, autoincrement=True),
    Column('nombre', String(100)),
    Column('fecha_creacion', DateTime, server_default=func.getdate())
)

# Crear la tabla si no existe
metadata.create_all(engine)
print("✅ Tabla 'dbo.NuevaTabla2' creada o ya existente.")

# Nuevos datos SIN ID (se generará automáticamente)
values_list = [
    {"nombre": "Cuarto registro"},
    {"nombre": "Quinto registro"},
    {"nombre": "Sexto registro"},
]

# Insertar nuevos registros
with engine.connect() as conn:
    stmt = insert(nueva_tabla)
    conn.execute(stmt, values_list)
    conn.commit()
    print("✅ Nuevos registros insertados.")

# Mostrar todos los registros
# with engine.connect() as conn:
#     result = conn.execute(select(nueva_tabla))
#     print("\n📋 Registros actuales en 'dbo.NuevaTabla2':")
#     for row in result.mappings():
#         print(dict(row))
