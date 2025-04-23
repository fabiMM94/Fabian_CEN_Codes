import pandas as pd

# Definir nombres de archivo
nombre_archivo_xlsx = 'PVData.xlsx'
nombre_archivo_csv = 'PVData.csv'

nombre_archivo_xlsx_2 = 'WPData.xlsx'
nombre_archivo_csv_2 = 'WPData.csv'

nombre_archivo_xlsx_3 = 'LFData.xlsx'
nombre_archivo_csv_3 = 'LFData.csv'

nombre_archivo_xlsx_4 = 'LoadData.xlsx'
nombre_archivo_csv_4 = 'LoadData.csv'

nombre_hoja = 0  # También puedes poner el nombre de la hoja, por ejemplo: 'Hoja1'

# Leer el archivo xlsx
df = pd.read_excel(nombre_archivo_xlsx, sheet_name=nombre_hoja)

# Guardar como CSV
df.to_csv(nombre_archivo_csv, index=False)

print(f"Archivo convertido exitosamente: {nombre_archivo_csv}")


# Leer el archivo xlsx
df2 = pd.read_excel(nombre_archivo_xlsx_2 , sheet_name=nombre_hoja)

# Guardar como CSV
df2.to_csv(nombre_archivo_csv_2, index=False)

print(f"Archivo convertido exitosamente: {nombre_archivo_csv_2}")

# Leer el archivo xlsx
df3 = pd.read_excel(nombre_archivo_xlsx_3 , sheet_name=nombre_hoja)

# Guardar como CSV
df3.to_csv(nombre_archivo_csv_3, index=False)

print(f"Archivo convertido exitosamente: {nombre_archivo_csv_3}")

# Leer el archivo xlsx
df4 = pd.read_excel(nombre_archivo_xlsx_4 , sheet_name=nombre_hoja)

# Guardar como CSV
df4.to_csv(nombre_archivo_csv_4, index=False)

print(f"Archivo convertido exitosamente: {nombre_archivo_csv_4}")