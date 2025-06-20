import pandas as pd
import os
import zipfile
import rarfile
import py7zr  # Para archivos .7z
from tkinter import filedialog, messagebox
import tkinter as tk
from datetime import datetime
import re

class PostProcesadorComtradeSubCarpetas:
    def __init__(self):
        """
        Post-procesador para analizar anexos en subcarpetas específicas
        Genera Excel con dos hojas: Cartas y Anexos
        """
        self.carpeta_anexos = ""
        self.archivo_excel = ""
        self.df_original = None
        self.df_cartas = None
        self.df_anexos = None
        
        # Configurar tkinter para dialogs
        self.root = tk.Tk()
        self.root.withdraw()  # Ocultar ventana principal
        
        print("🔍 POST-PROCESADOR COMTRADE CON SUBCARPETAS")
        print("="*60)
        print("Este código:")
        print("1. 📁 Selecciona carpeta de Anexos principal")
        print("2. 📊 Selecciona Excel de respuestas")
        print("3. 🔍 Analiza subcarpetas específicas por 'SubCarpeta Anexos'")
        print("4. 📄 Detecta archivos .dat y .cfg en subcarpetas")
        print("5. 📋 Genera Excel con 2 hojas: Cartas y Anexos")
        print("6. 🔗 Relaciona CartaID con AnexosID")
    
    def paso1_seleccionar_carpeta_anexos(self):
        """
        Paso 1: Seleccionar carpeta principal de Anexos
        """
        print(f"\n📁 PASO 1: SELECCIONAR CARPETA PRINCIPAL DE ANEXOS")
        print("="*50)
        
        messagebox.showinfo(
            "Seleccionar Carpeta Principal", 
            "Selecciona la carpeta principal 'Anexos' que contiene las subcarpetas"
        )
        
        self.carpeta_anexos = filedialog.askdirectory(title="Seleccionar carpeta principal 'Anexos'")
        
        if not self.carpeta_anexos:
            print("❌ No se seleccionó carpeta. Cancelando proceso.")
            return False
        
        print(f"✅ Carpeta principal seleccionada: {self.carpeta_anexos}")
        
        # Verificar subcarpetas
        try:
            contenido = os.listdir(self.carpeta_anexos)
            subcarpetas = [item for item in contenido if os.path.isdir(os.path.join(self.carpeta_anexos, item))]
            archivos = [item for item in contenido if os.path.isfile(os.path.join(self.carpeta_anexos, item))]
            
            print(f"📊 Total subcarpetas: {len(subcarpetas)}")
            print(f"📄 Total archivos sueltos: {len(archivos)}")
            
            # Mostrar preview de subcarpetas
            if subcarpetas:
                print(f"\n📋 Preview de subcarpetas (primeras 5):")
                for subcarpeta in subcarpetas[:5]:
                    print(f"   📁 {subcarpeta}")
                if len(subcarpetas) > 5:
                    print(f"   ... y {len(subcarpetas) - 5} subcarpetas más")
            
            return True
            
        except Exception as e:
            print(f"❌ Error accediendo a la carpeta: {e}")
            return False
    
    def paso2_seleccionar_excel(self):
        """
        Paso 2: Seleccionar archivo Excel con respuestas
        """
        print(f"\n📊 PASO 2: SELECCIONAR EXCEL DE RESPUESTAS")
        print("="*50)
        
        messagebox.showinfo(
            "Seleccionar Excel", 
            "Selecciona el archivo Excel que contiene la columna 'SubCarpeta Anexos'"
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
                'Correlativo Respuesta', 'Fecha de Envío', 'Empresa', 
                'Responde a', 'Documento Descargado', 'SubCarpeta Anexos',
                'Envía anexos', 'Anexos Descargados', 'Envío Principal'
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
                subcarpeta = row.get('SubCarpeta Anexos', 'N/A')
                empresa = row.get('Empresa', 'N/A')
                print(f"   {index+1}. {correlativo} - {empresa}")
                print(f"      SubCarpeta: {subcarpeta}")
            
            return True
            
        except Exception as e:
            print(f"❌ Error cargando Excel: {e}")
            return False
    
    def analizar_archivo_comprimido(self, ruta_archivo):
        """
        Analiza un archivo .rar, .zip o .7z para verificar si contiene .dat y .cfg
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
                with zipfile.ZipFile(ruta_archivo, 'r') as zip_file:
                    archivos_internos = zip_file.namelist()
                    
            elif extension == '.rar':
                with rarfile.RarFile(ruta_archivo, 'r') as rar_file:
                    archivos_internos = rar_file.namelist()
                    
            elif extension == '.7z':
                with py7zr.SevenZipFile(ruta_archivo, mode='r') as sz_file:
                    archivos_internos = sz_file.getnames()
                    
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
    
    def analizar_subcarpeta_recursivamente(self, ruta_subcarpeta):
        """
        Analiza recursivamente una subcarpeta buscando archivos .dat y .cfg
        Incluye archivos sueltos, comprimidos y subcarpetas anidadas
        """
        resultado = {
            'tiene_dat': False,
            'tiene_cfg': False,
            'detalles': [],
            'error': None
        }
        
        try:
            if not os.path.exists(ruta_subcarpeta):
                resultado['error'] = f"Subcarpeta no encontrada: {ruta_subcarpeta}"
                return resultado
            
            # Función recursiva para explorar directorios
            def explorar_directorio(directorio, nivel=0):
                indent = "  " * nivel
                
                try:
                    contenido = os.listdir(directorio)
                    archivos = []
                    subdirectorios = []
                    
                    for item in contenido:
                        ruta_item = os.path.join(directorio, item)
                        if os.path.isfile(ruta_item):
                            archivos.append((item, ruta_item))
                        elif os.path.isdir(ruta_item):
                            subdirectorios.append((item, ruta_item))
                    
                    # Analizar archivos en este nivel
                    for nombre_archivo, ruta_archivo in archivos:
                        nombre_lower = nombre_archivo.lower()
                        
                        if nombre_lower.endswith('.dat'):
                            resultado['tiene_dat'] = True
                            resultado['detalles'].append(f"{indent}📄 .dat encontrado: {nombre_archivo}")
                            
                        elif nombre_lower.endswith('.cfg'):
                            resultado['tiene_cfg'] = True
                            resultado['detalles'].append(f"{indent}📄 .cfg encontrado: {nombre_archivo}")
                            
                        elif nombre_lower.endswith(('.zip', '.rar', '.7z')):
                            # Analizar archivo comprimido
                            analisis_comprimido = self.analizar_archivo_comprimido(ruta_archivo)
                            
                            if analisis_comprimido['error']:
                                resultado['detalles'].append(f"{indent}❌ Error en {nombre_archivo}: {analisis_comprimido['error']}")
                            else:
                                if analisis_comprimido['tiene_dat']:
                                    resultado['tiene_dat'] = True
                                    resultado['detalles'].append(f"{indent}📦 {nombre_archivo}: contiene .dat")
                                
                                if analisis_comprimido['tiene_cfg']:
                                    resultado['tiene_cfg'] = True
                                    resultado['detalles'].append(f"{indent}📦 {nombre_archivo}: contiene .cfg")
                                
                                if not analisis_comprimido['tiene_dat'] and not analisis_comprimido['tiene_cfg']:
                                    archivos_internos = len(analisis_comprimido['archivos_encontrados'])
                                    resultado['detalles'].append(f"{indent}📦 {nombre_archivo}: {archivos_internos} archivos, sin .dat/.cfg")
                    
                    # Explorar subdirectorios recursivamente
                    for nombre_subdir, ruta_subdir in subdirectorios:
                        resultado['detalles'].append(f"{indent}📁 Explorando: {nombre_subdir}")
                        explorar_directorio(ruta_subdir, nivel + 1)
                
                except PermissionError:
                    resultado['detalles'].append(f"{indent}❌ Sin permisos para acceder: {os.path.basename(directorio)}")
                except Exception as e:
                    resultado['detalles'].append(f"{indent}❌ Error explorando {os.path.basename(directorio)}: {str(e)}")
            
            # Iniciar exploración recursiva
            resultado['detalles'].append(f"🔍 Analizando: {os.path.basename(ruta_subcarpeta)}")
            explorar_directorio(ruta_subcarpeta)
            
        except Exception as e:
            resultado['error'] = str(e)
        
        return resultado
    
    def convertir_si_no_a_verdadero_falso(self, valor):
        """
        Convierte valores Si/No a VERDADERO/FALSO
        """
        if isinstance(valor, str):
            valor_lower = valor.lower().strip()
            if valor_lower in ['si', 'sí', 'yes', 'true', '1']:
                return 'VERDADERO'
            elif valor_lower in ['no', 'false', '0']:
                return 'FALSO'
        return valor  # Mantener valor original si no coincide
    
    def procesar_anexos_descargados(self, anexos_descargados_str):
        """
        Procesa la cadena de anexos descargados y devuelve una lista
        """
        if not anexos_descargados_str or anexos_descargados_str in ['Sin anexos descargados', 'Sin anexos', 'Error', '']:
            return []
        
        # Separar por ' | ' y limpiar
        anexos_lista = [anexo.strip() for anexo in str(anexos_descargados_str).split(' | ') if anexo.strip()]
        return anexos_lista
    
    def paso3_analizar_y_crear_dataframes(self):
        """
        Paso 3: Analizar subcarpetas y crear DataFrames para las hojas Cartas y Anexos
        """
        print(f"\n🔍 PASO 3: ANALIZANDO SUBCARPETAS Y CREANDO DATAFRAMES")
        print("="*60)
        
        # Inicializar DataFrames
        self.df_cartas = pd.DataFrame(columns=[
            'CartaID', 'Correlativo', 'Fecha de Envío', 'Empresa', 
            'Responde a', 'PDF carta', 'SubCarpeta Anexos', 
            'Envía anexos', 'Contiene Comtrade', 'URL correlativo'
        ])
        
        self.df_anexos = pd.DataFrame(columns=[
            'AnexoID', 'CartaID', 'Anexos Descargados', 'SubCarpeta Anexos'
        ])
        
        total_filas = len(self.df_original)
        print(f"📊 Procesando {total_filas} respuestas...")
        print("="*60)
        
        # Contadores
        carta_id = 1
        anexo_id = 1
        
        # Estadísticas
        stats = {
            'sin_subcarpeta': 0,
            'subcarpeta_no_encontrada': 0,
            'con_comtrade': 0,
            'sin_comtrade': 0,
            'total_anexos': 0
        }
        
        for index, row in self.df_original.iterrows():
            correlativo = row.get('Correlativo Respuesta', '')
            subcarpeta_anexos = row.get('SubCarpeta Anexos', '')
            anexos_descargados = row.get('Anexos Descargados', '')
            
            print(f"\n[{index+1}/{total_filas}] Procesando: {correlativo}")
            print(f"   📁 SubCarpeta: {subcarpeta_anexos}")
            
            # Crear fila para Cartas
            fila_carta = {
                'CartaID': carta_id,
                'Correlativo': correlativo,
                'Fecha de Envío': row.get('Fecha de Envío', ''),
                'Empresa': row.get('Empresa', ''),
                'Responde a': row.get('Responde a', ''),
                'PDF carta': row.get('Documento Descargado', ''),
                'SubCarpeta Anexos': subcarpeta_anexos,
                'Envía anexos': self.convertir_si_no_a_verdadero_falso(row.get('Envía anexos', '')),
                'Contiene Comtrade': 'FALSO',  # Por defecto
                'URL correlativo': row.get('Envío Principal', '')
            }
            
            # Procesar anexos descargados para la hoja Anexos
            anexos_lista = self.procesar_anexos_descargados(anexos_descargados)
            
            if anexos_lista:
                print(f"   📎 Anexos encontrados: {len(anexos_lista)}")
                for anexo in anexos_lista:
                    fila_anexo = {
                        'AnexoID': anexo_id,
                        'CartaID': carta_id,
                        'Anexos Descargados': anexo,
                        'SubCarpeta Anexos': subcarpeta_anexos
                    }
                    self.df_anexos = pd.concat([self.df_anexos, pd.DataFrame([fila_anexo])], ignore_index=True)
                    anexo_id += 1
                    stats['total_anexos'] += 1
                    print(f"      📄 {anexo}")
            else:
                print(f"   ⚠️ Sin anexos descargados")
            
            # Analizar subcarpeta para determinar si contiene Comtrade
            if not subcarpeta_anexos or subcarpeta_anexos in ['', 'N/A', 'Sin subcarpeta']:
                print(f"   ❌ Sin subcarpeta especificada")
                fila_carta['Contiene Comtrade'] = 'FALSO'
                stats['sin_subcarpeta'] += 1
            else:
                ruta_subcarpeta = os.path.join(self.carpeta_anexos, subcarpeta_anexos)
                
                if not os.path.exists(ruta_subcarpeta):
                    print(f"   ❌ Subcarpeta no encontrada: {subcarpeta_anexos}")
                    fila_carta['Contiene Comtrade'] = 'FALSO'
                    stats['subcarpeta_no_encontrada'] += 1
                else:
                    print(f"   🔍 Analizando subcarpeta: {subcarpeta_anexos}")
                    
                    # Analizar recursivamente la subcarpeta
                    resultado_analisis = self.analizar_subcarpeta_recursivamente(ruta_subcarpeta)
                    
                    if resultado_analisis['error']:
                        print(f"   ❌ Error analizando subcarpeta: {resultado_analisis['error']}")
                        fila_carta['Contiene Comtrade'] = 'FALSO'
                    else:
                        # Mostrar detalles del análisis
                        for detalle in resultado_analisis['detalles']:
                            print(f"      {detalle}")
                        
                        # Determinar si contiene Comtrade (necesita .dat Y .cfg)
                        if resultado_analisis['tiene_dat'] and resultado_analisis['tiene_cfg']:
                            fila_carta['Contiene Comtrade'] = 'VERDADERO'
                            stats['con_comtrade'] += 1
                            print(f"   ✅ RESULTADO: Contiene Comtrade (.dat Y .cfg)")
                        else:
                            fila_carta['Contiene Comtrade'] = 'FALSO'
                            stats['sin_comtrade'] += 1
                            missing = []
                            if not resultado_analisis['tiene_dat']:
                                missing.append('.dat')
                            if not resultado_analisis['tiene_cfg']:
                                missing.append('.cfg')
                            print(f"   ❌ RESULTADO: Falta {' y '.join(missing)}")
            
            # Agregar fila a Cartas
            self.df_cartas = pd.concat([self.df_cartas, pd.DataFrame([fila_carta])], ignore_index=True)
            carta_id += 1
        
        # Mostrar estadísticas finales
        print(f"\n📈 ESTADÍSTICAS FINALES:")
        print("="*50)
        print(f"📊 Total cartas procesadas: {total_filas}")
        print(f"📎 Total anexos procesados: {stats['total_anexos']}")
        print(f"✅ Cartas con Comtrade válido: {stats['con_comtrade']}")
        print(f"❌ Cartas sin Comtrade válido: {stats['sin_comtrade']}")
        print(f"⚠️ Cartas sin subcarpeta: {stats['sin_subcarpeta']}")
        print(f"📁 Subcarpetas no encontradas: {stats['subcarpeta_no_encontrada']}")
        
        return True
    
    def paso4_exportar_excel_final(self):
        """
        Paso 4: Exportar Excel final con dos hojas
        """
        print(f"\n📋 PASO 4: EXPORTANDO EXCEL CON DOS HOJAS")
        print("="*50)
        
        # Generar nombre de archivo
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        nombre_archivo = f"analisis_comtrade_subcarpetas_{timestamp}.xlsx"
        
        try:
            # Crear ExcelWriter para múltiples hojas
            with pd.ExcelWriter(nombre_archivo, engine='openpyxl') as writer:
                # Exportar hoja Cartas
                self.df_cartas.to_excel(writer, sheet_name='Cartas', index=False)
                
                # Exportar hoja Anexos
                self.df_anexos.to_excel(writer, sheet_name='Anexos', index=False)
            
            print(f"✅ Excel final exportado: {nombre_archivo}")
            print(f"\n📊 HOJA 'CARTAS':")
            print(f"   📋 Total filas: {len(self.df_cartas)}")
            print(f"   📋 Columnas: {len(self.df_cartas.columns)}")
            
            print(f"\n📊 HOJA 'ANEXOS':")
            print(f"   📋 Total filas: {len(self.df_anexos)}")
            print(f"   📋 Columnas: {len(self.df_anexos.columns)}")
            
            # Preview de la hoja Cartas
            print(f"\n📋 PREVIEW HOJA 'CARTAS' (primeras 3 filas):")
            print("="*60)
            for index, row in self.df_cartas.head(3).iterrows():
                carta_id = row.get('CartaID', 'N/A')
                correlativo = row.get('Correlativo', 'N/A')
                contiene_comtrade = row.get('Contiene Comtrade', 'N/A')
                
                print(f"{carta_id}. {correlativo}")
                print(f"   📊 Contiene Comtrade: {contiene_comtrade}")
            
            # Preview de la hoja Anexos
            print(f"\n📋 PREVIEW HOJA 'ANEXOS' (primeras 5 filas):")
            print("="*60)
            for index, row in self.df_anexos.head(5).iterrows():
                anexo_id = row.get('AnexoID', 'N/A')
                carta_id = row.get('CartaID', 'N/A')
                anexo = row.get('Anexos Descargados', 'N/A')
                subcarpeta = row.get('SubCarpeta Anexos', 'N/A')
                
                print(f"AnexoID {anexo_id} -> CartaID {carta_id}: {anexo}")
                print(f"   📁 SubCarpeta: {subcarpeta}")
            
            # Estadísticas de Contiene Comtrade
            if 'Contiene Comtrade' in self.df_cartas.columns:
                comtrade_stats = self.df_cartas['Contiene Comtrade'].value_counts()
                print(f"\n📈 RESUMEN 'CONTIENE COMTRADE':")
                print("="*30)
                for valor, count in comtrade_stats.items():
                    emoji = "✅" if valor == "VERDADERO" else "❌"
                    print(f"{emoji} {valor}: {count} cartas")
            
            print(f"\n🎯 ¡PROCESO COMPLETADO EXITOSAMENTE!")
            print(f"📁 Archivo final: {os.path.abspath(nombre_archivo)}")
            
            return True
            
        except Exception as e:
            print(f"❌ Error exportando Excel: {e}")
            return False
    
    def ejecutar_proceso_completo(self):
        """
        Ejecuta todo el proceso de post-procesamiento con subcarpetas
        """
        try:
            # Paso 1: Seleccionar carpeta Anexos principal
            if not self.paso1_seleccionar_carpeta_anexos():
                return False
            
            # Paso 2: Seleccionar Excel
            if not self.paso2_seleccionar_excel():
                return False
            
            # Paso 3: Analizar subcarpetas y crear DataFrames
            if not self.paso3_analizar_y_crear_dataframes():
                return False
            
            # Paso 4: Exportar Excel final con dos hojas
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
    print("🚀 INICIANDO POST-PROCESADOR CON SUBCARPETAS")
    print("="*60)
    
    procesador = PostProcesadorComtradeSubCarpetas()
    
    try:
        exito = procesador.ejecutar_proceso_completo()
        
        if exito:
            print("\n✅ PROCESO EXITOSO")
            print("📋 Se ha creado el Excel con dos hojas:")
            print("   📄 Hoja 'Cartas': Información de cartas con CartaID")
            print("   📄 Hoja 'Anexos': Relación AnexoID-CartaID-Anexos")
            print("🔍 Análisis incluye exploración recursiva de subcarpetas")
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