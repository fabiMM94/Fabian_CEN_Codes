import pandas as pd
import os
import zipfile
import rarfile
import py7zr  # Para archivos .7z
from tkinter import filedialog, messagebox
import tkinter as tk
from datetime import datetime
import re

class PostProcesadorComtradeMejorado:
    def __init__(self):
        """
        Post-procesador mejorado para analizar anexos y determinar contenido Comtrade
        Soporta múltiples archivos por correlativo, archivos sueltos y .7z
        """
        self.carpeta_anexos = ""
        self.archivo_excel = ""
        self.df_original = None
        self.df_final = None
        
        # Configurar tkinter para dialogs
        self.root = tk.Tk()
        self.root.withdraw()  # Ocultar ventana principal
        
        print("🔍 POST-PROCESADOR DE ANÁLISIS COMTRADE MEJORADO")
        print("="*60)
        print("Este código mejorado:")
        print("1. 📁 Selecciona carpeta de Anexos")
        print("2. 📊 Selecciona Excel de respuestas")
        print("3. 🔍 Analiza contenido de archivos .rar/.zip/.7z")
        print("4. 📄 Detecta archivos .dat y .cfg sueltos")
        print("5. 📋 Maneja múltiples archivos por correlativo")
        print("6. ✅ Verifica presencia de archivos .dat Y .cfg")
        print("7. 📋 Crea Excel final con columna 'Contiene Comtrade'")
    
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
        archivos_comprimidos = [f for f in archivos if f.lower().endswith(('.rar', '.zip', '.7z'))]
        archivos_dat_cfg = [f for f in archivos if f.lower().endswith(('.dat', '.cfg'))]
        
        print(f"📊 Total archivos en carpeta: {len(archivos)}")
        print(f"📦 Archivos comprimidos (.rar/.zip/.7z): {len(archivos_comprimidos)}")
        print(f"📄 Archivos .dat/.cfg sueltos: {len(archivos_dat_cfg)}")
        
        if len(archivos_comprimidos) == 0 and len(archivos_dat_cfg) == 0:
            print("⚠️ No se encontraron archivos relevantes (.rar/.zip/.7z/.dat/.cfg) en la carpeta")
            return False
        
        # Mostrar preview de archivos
        print(f"\n📋 Preview de archivos relevantes (primeros 10):")
        archivos_relevantes = archivos_comprimidos + archivos_dat_cfg
        for archivo in archivos_relevantes[:10]:
            print(f"   • {archivo}")
        
        if len(archivos_relevantes) > 10:
            print(f"   ... y {len(archivos_relevantes) - 10} archivos más")
        
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
        Analiza un archivo .rar, .zip o .7z para verificar si contiene .dat y .cfg
        
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
                    
            elif extension == '.7z':
                # Analizar archivo 7Z
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
    
    def buscar_archivos_correlativo(self, correlativo, nombres_anexos):
        """
        Busca TODOS los archivos relacionados con un correlativo específico
        
        Args:
            correlativo (str): Correlativo de la respuesta
            nombres_anexos (str): Nombres de anexos separados por ' | '
        
        Returns:
            dict: Archivos encontrados categorizados por tipo
        """
        archivos_encontrados = {
            'comprimidos': [],  # .rar, .zip, .7z
            'dat_sueltos': [],  # .dat individuales
            'cfg_sueltos': [],  # .cfg individuales
            'otros': []         # otros archivos
        }
        
        if not nombres_anexos or nombres_anexos in ['Sin anexos descargados', 'Sin anexos', 'Error']:
            return archivos_encontrados
        
        # Obtener todos los archivos de la carpeta
        todos_archivos = os.listdir(self.carpeta_anexos)
        
        # Buscar archivos que empiecen con el correlativo
        archivos_correlativo = []
        for archivo in todos_archivos:
            if archivo.startswith(correlativo):
                ruta_completa = os.path.join(self.carpeta_anexos, archivo)
                if os.path.isfile(ruta_completa):
                    archivos_correlativo.append(ruta_completa)
        
        # También buscar por nombres específicos mencionados en anexos_descargados
        nombres_lista = nombres_anexos.split(' | ')
        for nombre_anexo in nombres_lista:
            nombre_anexo = nombre_anexo.strip()
            
            # Buscar archivo exacto
            ruta_exacta = os.path.join(self.carpeta_anexos, nombre_anexo)
            if os.path.exists(ruta_exacta) and ruta_exacta not in archivos_correlativo:
                archivos_correlativo.append(ruta_exacta)
            
            # Buscar por coincidencia parcial
            for archivo in todos_archivos:
                if nombre_anexo in archivo or archivo in nombre_anexo:
                    ruta_completa = os.path.join(self.carpeta_anexos, archivo)
                    if ruta_completa not in archivos_correlativo and os.path.isfile(ruta_completa):
                        archivos_correlativo.append(ruta_completa)
        
        # Categorizar archivos encontrados
        for archivo_path in archivos_correlativo:
            nombre_archivo = os.path.basename(archivo_path).lower()
            
            if nombre_archivo.endswith(('.rar', '.zip', '.7z')):
                archivos_encontrados['comprimidos'].append(archivo_path)
            elif nombre_archivo.endswith('.dat'):
                archivos_encontrados['dat_sueltos'].append(archivo_path)
            elif nombre_archivo.endswith('.cfg'):
                archivos_encontrados['cfg_sueltos'].append(archivo_path)
            else:
                archivos_encontrados['otros'].append(archivo_path)
        
        return archivos_encontrados
    
    def determinar_contiene_comtrade(self, archivos_encontrados):
        """
        Determina si los archivos contienen un Comtrade válido (.dat Y .cfg)
        
        Args:
            archivos_encontrados (dict): Archivos categorizados por tipo
        
        Returns:
            tuple: (tiene_comtrade, detalles_analisis)
        """
        tiene_dat = False
        tiene_cfg = False
        detalles = []
        
        # 1. Verificar archivos .dat y .cfg sueltos
        if archivos_encontrados['dat_sueltos']:
            tiene_dat = True
            detalles.append(f"📄 .dat sueltos: {len(archivos_encontrados['dat_sueltos'])}")
        
        if archivos_encontrados['cfg_sueltos']:
            tiene_cfg = True
            detalles.append(f"📄 .cfg sueltos: {len(archivos_encontrados['cfg_sueltos'])}")
        
        # 2. Analizar archivos comprimidos
        for archivo_comprimido in archivos_encontrados['comprimidos']:
            nombre_archivo = os.path.basename(archivo_comprimido)
            resultado = self.analizar_archivo_comprimido(archivo_comprimido)
            
            if resultado['error']:
                detalles.append(f"❌ Error en {nombre_archivo}: {resultado['error']}")
            else:
                if resultado['tiene_dat']:
                    tiene_dat = True
                    detalles.append(f"📦 {nombre_archivo}: contiene .dat")
                
                if resultado['tiene_cfg']:
                    tiene_cfg = True
                    detalles.append(f"📦 {nombre_archivo}: contiene .cfg")
                
                if not resultado['tiene_dat'] and not resultado['tiene_cfg']:
                    archivos_internos = len(resultado['archivos_encontrados'])
                    detalles.append(f"📦 {nombre_archivo}: {archivos_internos} archivos, sin .dat/.cfg")
        
        # 3. Determinar resultado final
        contiene_comtrade = tiene_dat and tiene_cfg
        
        if contiene_comtrade:
            detalles.append("✅ RESULTADO: Contiene .dat Y .cfg")
        else:
            missing = []
            if not tiene_dat:
                missing.append(".dat")
            if not tiene_cfg:
                missing.append(".cfg")
            detalles.append(f"❌ RESULTADO: Falta {' y '.join(missing)}")
        
        return contiene_comtrade, detalles
    
    def paso3_analizar_y_crear_excel(self):
        """
        Paso 3: Analizar anexos y crear Excel final con lógica mejorada
        """
        print(f"\n🔍 PASO 3: ANALIZANDO ANEXOS CON LÓGICA MEJORADA")
        print("="*60)
        print("🔍 Nuevo análisis incluye:")
        print("   📦 Archivos comprimidos (.rar/.zip/.7z)")
        print("   📄 Archivos .dat/.cfg sueltos")
        print("   📋 Múltiples archivos por correlativo")
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
        
        # Contadores para estadísticas mejoradas
        stats = {
            'sin_anexos': 0,
            'con_comtrade': 0,
            'sin_comtrade': 0,
            'errores': 0,
            'archivos_sueltos': 0,
            'archivos_comprimidos': 0,
            'multiples_archivos': 0
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
            
            # Verificar anexos con lógica mejorada
            if anexos_descargados in ['Sin anexos descargados', 'Sin anexos', 'Error', '']:
                print(f"   ⚠️ Sin anexos descargados")
                fila_final['Contiene Comtrade'] = 'No'
                stats['sin_anexos'] += 1
            else:
                print(f"   📎 Anexos: {anexos_descargados}")
                
                # Buscar TODOS los archivos relacionados con este correlativo
                archivos_encontrados = self.buscar_archivos_correlativo(correlativo, anexos_descargados)
                
                total_archivos = (len(archivos_encontrados['comprimidos']) + 
                                len(archivos_encontrados['dat_sueltos']) + 
                                len(archivos_encontrados['cfg_sueltos']) + 
                                len(archivos_encontrados['otros']))
                
                if total_archivos == 0:
                    print(f"   ❌ No se encontraron archivos en carpeta")
                    fila_final['Contiene Comtrade'] = 'No'
                    stats['errores'] += 1
                else:
                    print(f"   📁 Encontrados {total_archivos} archivos:")
                    print(f"      📦 Comprimidos: {len(archivos_encontrados['comprimidos'])}")
                    print(f"      📄 .dat sueltos: {len(archivos_encontrados['dat_sueltos'])}")
                    print(f"      📄 .cfg sueltos: {len(archivos_encontrados['cfg_sueltos'])}")
                    print(f"      📋 Otros: {len(archivos_encontrados['otros'])}")
                    
                    # Actualizar estadísticas
                    if total_archivos > 1:
                        stats['multiples_archivos'] += 1
                    if archivos_encontrados['dat_sueltos'] or archivos_encontrados['cfg_sueltos']:
                        stats['archivos_sueltos'] += 1
                    if archivos_encontrados['comprimidos']:
                        stats['archivos_comprimidos'] += 1
                    
                    # Determinar si contiene Comtrade
                    contiene_comtrade, detalles = self.determinar_contiene_comtrade(archivos_encontrados)
                    
                    # Mostrar detalles del análisis
                    for detalle in detalles:
                        print(f"      {detalle}")
                    
                    if contiene_comtrade:
                        fila_final['Contiene Comtrade'] = 'Si'
                        stats['con_comtrade'] += 1
                        print(f"   ✅ RESULTADO FINAL: Contiene Comtrade")
                    else:
                        fila_final['Contiene Comtrade'] = 'No'
                        stats['sin_comtrade'] += 1
                        print(f"   ❌ RESULTADO FINAL: No contiene Comtrade válido")
            
            # Agregar fila al DataFrame final
            self.df_final = pd.concat([self.df_final, pd.DataFrame([fila_final])], ignore_index=True)
        
        # Mostrar estadísticas finales mejoradas
        print(f"\n📈 ESTADÍSTICAS FINALES MEJORADAS:")
        print("="*50)
        print(f"📊 Total respuestas procesadas: {total_filas}")
        print(f"✅ Con Comtrade válido: {stats['con_comtrade']}")
        print(f"❌ Sin Comtrade válido: {stats['sin_comtrade']}")
        print(f"⚠️ Sin anexos: {stats['sin_anexos']}")
        print(f"🔧 Errores: {stats['errores']}")
        print(f"\n📋 DETALLES DE ANÁLISIS:")
        print(f"📄 Respuestas con archivos sueltos (.dat/.cfg): {stats['archivos_sueltos']}")
        print(f"📦 Respuestas con archivos comprimidos: {stats['archivos_comprimidos']}")
        print(f"📂 Respuestas con múltiples archivos: {stats['multiples_archivos']}")
        
        return True
    
    def paso4_exportar_excel_final(self):
        """
        Paso 4: Exportar Excel final
        """
        print(f"\n📋 PASO 4: EXPORTANDO EXCEL FINAL")
        print("="*50)
        
        # Generar nombre de archivo
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        nombre_archivo = f"analisis_comtrade_final_mejorado_{timestamp}.xlsx"
        
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
    print("🚀 INICIANDO POST-PROCESADOR DE ANÁLISIS COMTRADE MEJORADO")
    print("="*60)
    
    procesador = PostProcesadorComtradeMejorado()
    
    try:
        exito = procesador.ejecutar_proceso_completo()
        
        if exito:
            print("\n✅ PROCESO EXITOSO")
            print("📋 Se ha creado el Excel final con la columna 'Contiene Comtrade'")
            print("🔍 Análisis mejorado incluye:")
            print("   • Soporte para archivos .7z")
            print("   • Detección de .dat/.cfg sueltos")
            print("   • Manejo de múltiples archivos por correlativo")
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