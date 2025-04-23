
import pandas as pd
import glob
import os

def csv_to_xlsx(csv_file, xlsx_file):
    try:
        df = pd.read_csv(csv_file)
        df.to_excel(xlsx_file, index=False, engine='openpyxl')
        print(f"Conversión exitosa: {xlsx_file}")
    except Exception as e:
        print(f"Error: {e}")

def convert_all_csv_to_xlsx():
    # Obtener el directorio de trabajo actual
    directory = os.getcwd()  # Detecta la ruta en la que estamos ejecutando el script
    print(f"Buscando archivos CSV en: {directory}")  # Mostrar el directorio actual
    
    # Verificar si el directorio existe
    if not os.path.exists(directory):
        print(f"El directorio {directory} no existe.")
        return
    
    # Buscar todos los archivos CSV en el directorio
    csv_files = glob.glob(os.path.join(directory, "*.csv"))
    print(f"Archivos encontrados: {csv_files}")  # Mostrar los archivos encontrados
    
    if not csv_files:
        print("No se encontraron archivos CSV en el directorio.")
        return
    
    # Convertir cada archivo CSV a XLSX
    for csv_file in csv_files:
        xlsx_file = os.path.splitext(csv_file)[0] + ".xlsx"
        csv_to_xlsx(csv_file, xlsx_file)

# Uso
def main():
    # Llamar a la función que convierte los CSV a XLSX en el directorio actual
    convert_all_csv_to_xlsx()

if __name__ == "__main__":
    main()