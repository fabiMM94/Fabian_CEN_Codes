import pandas as pd
import os
import zipfile
import rarfile
from tkinter import filedialog, messagebox
import tkinter as tk
from datetime import datetime
import re

class PostProcesadorComtrade:
    def __init__(self):
        """
        Post-procesador para analizar anexos y determinar contenido Comtrade
        """
        self.carpeta_anexos = ""
        self.archivo_excel = ""
        self.df_original = None
        self.df_final = None
        
        # Configurar tkinter para dialogs
        self.root = tk.Tk()
        self.root.withdraw()  # Ocultar ventana principal
        
        print("🔍 POST-PROCESADOR DE ANÁLISIS COMTRADE")
        print("="*60)
        print("Este código:")
        print("1. 📁 Selecciona carpeta de Anexos")
        print("2. 📊 Selecciona Excel de respuestas")
        print("3. 🔍 Analiza contenido de archivos .rar/.zip")
        print("4. ✅ Verifica presencia de archivos .dat y .cfg")
        print("5. 📋 Crea Excel final con columna 'Contiene Comtrade'")
    
    def paso1_seleccionar_carpeta_anexos(self):
        """
        Paso 1: Seleccionar carpeta de Anexos
        """
        print(f"\n📁 PASO 1: SELECCIONAR CARPETA DE ANEXOS")
        print("="*50)
        
        messagebox.showinfo(
            "Seleccionar Carpeta", 
            "A continuación selecciona la carpeta 'Anexos' que contiene los archivos descargados"
        )
        
        self.carpeta_anexos = filedialog.askdirectory(title="Seleccionar carpeta 'Anexos'")
        
        if not self.carpeta_anexos:
            print("❌ No se seleccionó carpeta. Cancelando proceso.")
            return False
        
        print(f"✅ Carpeta seleccionada: {self.carpeta_anexos}")
        
        # Verificar que la carpeta contiene archivos
        archivos = [f for f in os.listdir(self.carpeta_anexos) if os.path.isfile(os.path.join(self.carpeta_anexos, f))]
        archivos_rar_zip = [f for f in archivos if f.lower().endswith(('.rar', '.zip'))]
        
        print(f"📊 Total archivos en carpeta: {len(archivos)}")
        print(f"📦 Archivos .rar/.zip encontrados: {len(archivos_rar_zip)}")
        
        if len(archivos_rar_zip) == 0:
            print("⚠️ No se encontraron archivos .rar o .zip en la carpeta")
            return False
        
        # Mostrar preview de archivos
        print(f"\n📋 Preview de archivos .rar/.zip (primeros 5):")
        for archivo in archivos_rar_zip[:5]:
            print(f"   • {archivo}")
        
        if len(archivos_rar_zip) > 5:
            print(f"   ... y {len(archivos_rar_zip) - 5} archivos más")
        
        return True
    
    def paso2_seleccionar_excel(self):
        """
        Paso 2: Seleccionar archivo Excel con respuestas
        """
        print(f"\n📊 PASO 2: SELECCIONAR EXCEL DE RESPUESTAS")
        print("="*50)
        
        messagebox.showinfo(
            "Seleccionar Excel", 
            "A continuación selecciona el archivo Excel que contiene las respuestas con columna 'Anexos Descargados'"
        )
        
        self.archivo_excel = filedialog.askopenfilename(
            title="Seleccionar archivo Excel de respuestas",
            filetypes=[("Archivos Excel", "*.xlsx *.xls"), ("Todos los archivos", "*.*")]
        )
        
        if not self.archivo_excel:
            print("❌ No se seleccionó archivo Excel. Cancelando proceso.")
            return False
        
        print(f"✅ Excel seleccionado: {os.path.basename(self.archivo_excel)}")
        
        # Cargar y verificar Excel
        try:
            self.df_original = pd.read_excel(self.archivo_excel, engine='openpyxl')
            print(f"📊 Filas cargadas: {len(self.df_original)}")
            print(f"📋 Columnas: {len(self.df_original.columns)}")
            
            # Verificar columnas necesarias
            columnas_necesarias = [
                'Número Respuesta', 'Correlativo Respuesta', 'Fecha de Envío', 
                'Empresa', 'Responde a', 'Anexos Descargados'
            ]
            
            columnas_faltantes = [col for col in columnas_necesarias if col not in self.df_original.columns]
            
            if columnas_faltantes:
                print(f"❌ Faltan columnas necesarias: {columnas_faltantes}")
                print(f"📋 Columnas disponibles: {list(self.df_original.columns)}")
                return False
            
            print("✅ Todas las columnas necesarias están presentes")
            
            # Mostrar preview
            print(f"\n📋 Preview del Excel (primeras 3 filas):")
            for index, row in self.df_original.head(3).iterrows():
                correlativo = row.get('Correlativo Respuesta', 'N/A')
                empresa = row.get('Empresa', 'N/A')
                anexos = row.get('Anexos Descargados', 'N/A')
                print(f"   {index+1}. {correlativo} - {empresa}")
                print(f"      Anexos: {anexos}")
            
            return True
            
        except Exception as e:
            print(f"❌ Error cargando Excel: {e}")
            return False
    
    def analizar_archivo_comprimido(self, ruta_archivo):
        """
        Analiza un archivo .rar o .zip para verificar si contiene .dat y .cfg
        
        Args:
            ruta_archivo (str): Ruta completa al archivo comprimido
        
        Returns:
            dict: Resultado del análisis
        """
        resultado = {
            'tiene_dat': False,
            'tiene_cfg': False,
            'archivos_encontrados': [],
            'error': None
        }
        
        try:
            extension = os.path.splitext(ruta_archivo)[1].lower()
            
            if extension == '.zip':
                # Analizar archivo ZIP
                with zipfile.ZipFile(ruta_archivo, 'r') as zip_file:
                    archivos_internos = zip_file.namelist()
                    
            elif extension == '.rar':
                # Analizar archivo RAR
                with rarfile.RarFile(ruta_archivo, 'r') as rar_file:
                    archivos_internos = rar_file.namelist()
            else:
                resultado['error'] = f"Extensión no soportada: {extension}"
                return resultado
            
            # Analizar archivos internos
            for archivo_interno in archivos_internos:
                nombre_archivo = os.path.basename(archivo_interno).lower()
                resultado['archivos_encontrados'].append(archivo_interno)
                
                if nombre_archivo.endswith('.dat'):
                    resultado['tiene_dat'] = True
                elif nombre_archivo.endswith('.cfg'):
                    resultado['tiene_cfg'] = True
            
        except Exception as e:
            resultado['error'] = str(e)
        
        return resultado
    
    def buscar_archivos_anexos(self, nombres_anexos):
        """
        Busca archivos de anexos en la carpeta basándose en los nombres
        
        Args:
            nombres_anexos (str): Nombres de anexos separados por ' | '
        
        Returns:
            list: Lista de rutas de archivos encontrados
        """
        if not nombres_anexos or nombres_anexos in ['Sin anexos descargados', 'Sin anexos', 'Error']:
            return []
        
        archivos_encontrados = []
        
        # Separar nombres de anexos
        nombres_lista = nombres_anexos.split(' | ')
        
        for nombre_anexo in nombres_lista:
            nombre_anexo = nombre_anexo.strip()
            
            # Buscar archivo exacto
            ruta_exacta = os.path.join(self.carpeta_anexos, nombre_anexo)
            if os.path.exists(ruta_exacta):
                archivos_encontrados.append(ruta_exacta)
                continue
            
            # Buscar por coincidencia parcial (en caso de que el nombre esté ligeramente diferente)
            archivos_carpeta = os.listdir(self.carpeta_anexos)
            for archivo_carpeta in archivos_carpeta:
                if nombre_anexo in archivo_carpeta or archivo_carpeta in nombre_anexo:
                    ruta_completa = os.path.join(self.carpeta_anexos, archivo_carpeta)
                    if ruta_completa not in archivos_encontrados:
                        archivos_encontrados.append(ruta_completa)
        
        return archivos_encontrados
    
    def paso3_analizar_y_crear_excel(self):
        """
        Paso 3: Analizar anexos y crear Excel final
        """
        print(f"\n🔍 PASO 3: ANALIZANDO ANEXOS Y CREANDO EXCEL FINAL")
        print("="*60)
        
        # Crear DataFrame final con columnas específicas
        columnas_finales = [
            'Número Respuesta', 'Correlativo Respuesta', 'Fecha de Envío', 
            'Empresa', 'Responde a', 'Contiene Comtrade'
        ]
        
        self.df_final = pd.DataFrame(columns=columnas_finales)
        
        total_filas = len(self.df_original)
        print(f"📊 Procesando {total_filas} respuestas...")
        print("="*60)
        
        # Contadores para estadísticas
        stats = {
            'sin_anexos': 0,
            'con_comtrade': 0,
            'sin_comtrade': 0,
            'errores': 0
        }
        
        for index, row in self.df_original.iterrows():
            correlativo = row.get('Correlativo Respuesta', '')
            anexos_descargados = row.get('Anexos Descargados', '')
            
            print(f"\n[{index+1}/{total_filas}] Analizando: {correlativo}")
            
            # Crear fila para el DataFrame final
            fila_final = {
                'Número Respuesta': row.get('Número Respuesta', ''),
                'Correlativo Respuesta': correlativo,
                'Fecha de Envío': row.get('Fecha de Envío', ''),
                'Empresa': row.get('Empresa', ''),
                'Responde a': row.get('Responde a', ''),
                'Contiene Comtrade': 'No'  # Por defecto
            }
            
            # Paso 5: Verificar anexos
            if anexos_descargados in ['Sin anexos descargados', 'Sin anexos', 'Error', '']:
                print(f"   ⚠️ Sin anexos descargados")
                fila_final['Contiene Comtrade'] = 'No'
                stats['sin_anexos'] += 1
            else:
                print(f"   📎 Anexos: {anexos_descargados}")
                
                # Buscar archivos en carpeta
                archivos_encontrados = self.buscar_archivos_anexos(anexos_descargados)
                
                if not archivos_encontrados:
                    print(f"   ❌ No se encontraron archivos en carpeta")
                    fila_final['Contiene Comtrade'] = 'No'
                    stats['errores'] += 1
                else:
                    print(f"   📁 Encontrados {len(archivos_encontrados)} archivos")
                    
                    # Analizar cada archivo comprimido
                    tiene_comtrade = False
                    
                    for archivo_path in archivos_encontrados:
                        if archivo_path.lower().endswith(('.rar', '.zip')):
                            nombre_archivo = os.path.basename(archivo_path)
                            print(f"      🔍 Analizando: {nombre_archivo}")
                            
                            resultado = self.analizar_archivo_comprimido(archivo_path)
                            
                            if resultado['error']:
                                print(f"         ❌ Error: {resultado['error']}")
                            else:
                                archivos_internos = resultado['archivos_encontrados']
                                tiene_dat = resultado['tiene_dat']
                                tiene_cfg = resultado['tiene_cfg']
                                
                                print(f"         📋 Archivos internos: {len(archivos_internos)}")
                                print(f"         📄 Tiene .dat: {'✅' if tiene_dat else '❌'}")
                                print(f"         📄 Tiene .cfg: {'✅' if tiene_cfg else '❌'}")
                                
                                if tiene_dat and tiene_cfg:
                                    print(f"         🎉 ¡CONTIENE COMTRADE VÁLIDO!")
                                    tiene_comtrade = True
                                    break  # Si uno tiene Comtrade, ya es suficiente
                        else:
                            print(f"      ⚠️ Archivo no comprimido: {os.path.basename(archivo_path)}")
                    
                    if tiene_comtrade:
                        fila_final['Contiene Comtrade'] = 'Si'
                        stats['con_comtrade'] += 1
                        print(f"   ✅ RESULTADO: Contiene Comtrade")
                    else:
                        fila_final['Contiene Comtrade'] = 'No'
                        stats['sin_comtrade'] += 1
                        print(f"   ❌ RESULTADO: No contiene Comtrade válido")
            
            # Agregar fila al DataFrame final
            self.df_final = pd.concat([self.df_final, pd.DataFrame([fila_final])], ignore_index=True)
        
        # Mostrar estadísticas finales
        print(f"\n📈 ESTADÍSTICAS FINALES:")
        print("="*40)
        print(f"📊 Total respuestas procesadas: {total_filas}")
        print(f"✅ Con Comtrade válido: {stats['con_comtrade']}")
        print(f"❌ Sin Comtrade válido: {stats['sin_comtrade']}")
        print(f"⚠️ Sin anexos: {stats['sin_anexos']}")
        print(f"🔧 Errores: {stats['errores']}")
        
        return True
    
    def paso4_exportar_excel_final(self):
        """
        Paso 4: Exportar Excel final
        """
        print(f"\n📋 PASO 4: EXPORTANDO EXCEL FINAL")
        print("="*50)
        
        # Generar nombre de archivo
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        nombre_archivo = f"analisis_comtrade_final_{timestamp}.xlsx"
        
        try:
            # Exportar a Excel
            self.df_final.to_excel(nombre_archivo, index=False, engine='openpyxl')
            
            print(f"✅ Excel final exportado: {nombre_archivo}")
            print(f"📊 Total filas: {len(self.df_final)}")
            print(f"📋 Columnas: {len(self.df_final.columns)}")
            
            # Mostrar preview
            print(f"\n📋 PREVIEW DEL EXCEL FINAL (primeras 3 filas):")
            print("="*60)
            for index, row in self.df_final.head(3).iterrows():
                correlativo = row.get('Correlativo Respuesta', 'N/A')
                empresa = row.get('Empresa', 'N/A')
                contiene_comtrade = row.get('Contiene Comtrade', 'N/A')
                
                print(f"{index+1}. {correlativo} - {empresa}")
                print(f"   📊 Contiene Comtrade: {contiene_comtrade}")
            
            if len(self.df_final) > 3:
                print(f"... y {len(self.df_final) - 3} filas más")
            
            # Estadísticas de la columna Contiene Comtrade
            if 'Contiene Comtrade' in self.df_final.columns:
                comtrade_stats = self.df_final['Contiene Comtrade'].value_counts()
                print(f"\n📈 RESUMEN 'CONTIENE COMTRADE':")
                print("="*30)
                for valor, count in comtrade_stats.items():
                    emoji = "✅" if valor == "Si" else "❌"
                    print(f"{emoji} {valor}: {count} respuestas")
            
            print(f"\n🎯 ¡PROCESO COMPLETADO EXITOSAMENTE!")
            print(f"📁 Archivo final: {os.path.abspath(nombre_archivo)}")
            
            return True
            
        except Exception as e:
            print(f"❌ Error exportando Excel: {e}")
            return False
    
    def ejecutar_proceso_completo(self):
        """
        Ejecuta todo el proceso de post-procesamiento
        """
        try:
            # Paso 1: Seleccionar carpeta Anexos
            if not self.paso1_seleccionar_carpeta_anexos():
                return False
            
            # Paso 2: Seleccionar Excel
            if not self.paso2_seleccionar_excel():
                return False
            
            # Paso 3: Analizar y crear Excel
            if not self.paso3_analizar_y_crear_excel():
                return False
            
            # Paso 4: Exportar Excel final
            if not self.paso4_exportar_excel_final():
                return False
            
            print(f"\n🎉 ¡PROCESO DE POST-PROCESAMIENTO COMPLETADO!")
            return True
            
        except Exception as e:
            print(f"❌ Error en proceso: {e}")
            import traceback
            traceback.print_exc()
            return False
        
        finally:
            # Cerrar tkinter
            if hasattr(self, 'root'):
                self.root.destroy()

# Función principal
if __name__ == "__main__":
    print("🚀 INICIANDO POST-PROCESADOR DE ANÁLISIS COMTRADE")
    print("="*60)
    
    procesador = PostProcesadorComtrade()
    
    try:
        exito = procesador.ejecutar_proceso_completo()
        
        if exito:
            print("\n✅ PROCESO EXITOSO")
            print("📋 Se ha creado el Excel final con la columna 'Contiene Comtrade'")
        else:
            print("\n❌ PROCESO FALLIDO")
            print("Revisa los mensajes de error anteriores")
    
    except KeyboardInterrupt:
        print("\n🛑 Proceso interrumpido por el usuario")
    
    except Exception as e:
        print(f"\n❌ Error inesperado: {e}")
        import traceback
        traceback.print_exc()
    
    input("\n⏳ Presiona ENTER para cerrar...")