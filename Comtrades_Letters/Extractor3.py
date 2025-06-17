import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
import re
from datetime import datetime

class ExtractorRespondidoPor:
    def __init__(self, headless=False):
        """
        Extractor especializado para seguir enlaces de "Respondido por"
        """
        self.driver = None
        self.is_logged_in = False
        self.setup_driver(headless)
        
    def setup_driver(self, headless=False):
        """Configura el driver optimizado con múltiples métodos"""
        chrome_options = Options()
        if headless:
            chrome_options.add_argument("--headless")
        chrome_options.add_argument("--no-sandbox")
        chrome_options.add_argument("--disable-dev-shm-usage")
        chrome_options.add_argument("--disable-gpu")
        chrome_options.add_argument("--window-size=1920,1080")
        chrome_options.add_argument("--disable-blink-features=AutomationControlled")
        
        # Método 1: Intentar con WebDriver Manager
        try:
            print("🔄 Intentando método 1: WebDriver Manager...")
            service = Service(ChromeDriverManager().install())
            self.driver = webdriver.Chrome(service=service, options=chrome_options)
            self.driver.implicitly_wait(10)
            print("✅ Driver configurado correctamente con WebDriver Manager")
            return
        except Exception as e:
            print(f"⚠️ Método 1 falló: {e}")
        
        # Método 2: Intentar sin service (Chrome en PATH)
        try:
            print("🔄 Intentando método 2: Chrome en PATH...")
            self.driver = webdriver.Chrome(options=chrome_options)
            self.driver.implicitly_wait(10)
            print("✅ Driver configurado correctamente usando Chrome en PATH")
            return
        except Exception as e:
            print(f"⚠️ Método 2 falló: {e}")
        
        # Método 3: ChromeDriver manual
        try:
            print("🔄 Intentando método 3: ChromeDriver manual...")
            # Rutas comunes donde puede estar ChromeDriver
            rutas_chromedriver = [
                "./chromedriver.exe",
                "./chromedriver",
                "C:/chromedriver/chromedriver.exe",
                "C:/Windows/System32/chromedriver.exe",
                "/usr/local/bin/chromedriver",
                "/usr/bin/chromedriver"
            ]
            
            for ruta in rutas_chromedriver:
                try:
                    service = Service(ruta)
                    self.driver = webdriver.Chrome(service=service, options=chrome_options)
                    self.driver.implicitly_wait(10)
                    print(f"✅ Driver configurado correctamente usando: {ruta}")
                    return
                except:
                    continue
                    
        except Exception as e:
            print(f"⚠️ Método 3 falló: {e}")
        
        # Si todo falla, mostrar instrucciones
        print("❌ No se pudo configurar ChromeDriver automáticamente")
        print("\n📋 SOLUCIONES:")
        print("1. Descarga ChromeDriver manualmente:")
        print("   - Ve a: https://chromedriver.chromium.org/downloads")
        print("   - Descarga la versión que coincida con tu Chrome")
        print("   - Coloca chromedriver.exe en la misma carpeta que este script")
        print("\n2. O verifica tu versión de Chrome:")
        print("   - Abre Chrome → Menú → Ayuda → Información de Google Chrome")
        print("   - Anota la versión")
        print("\n3. Verificar conexión a internet")
        
        self.driver = None

    def login_interactivo(self):
        """Login manual paso a paso"""
        if not self.driver:
            print("❌ Driver no configurado. No se puede continuar.")
            return False
            
        print("\n🖱️  MODO LOGIN PASO A PASO")
        print("="*50)
        print("PASO 1: Abrir página principal")
        print("PASO 2: Hacer login manual")
        print("PASO 3: Confirmar que estás autenticado")
        
        try:
            # PASO 1: Ir a la página principal
            print(f"\n🔄 PASO 1: Abriendo https://correspondencia.coordinador.cl/")
            self.driver.get("https://correspondencia.coordinador.cl/")
            time.sleep(3)
            
            print("✅ Página principal cargada")
            print("\n👤 PASO 2: HAZ LOGIN MANUALMENTE EN EL NAVEGADOR")
            print("- Ingresa tu usuario")
            print("- Ingresa tu contraseña") 
            print("- Haz click en 'Iniciar Sesión'")
            print("- Espera a que aparezca el dashboard/menú principal")
            
            input("\n⏳ Presiona ENTER después de hacer login exitoso...")
            
            # PASO 3: Verificar que el login fue exitoso
            print(f"\n🔍 PASO 3: Verificando login...")
            current_url = self.driver.current_url
            page_source = self.driver.page_source.lower()
            
            print(f"URL actual: {current_url}")
            
            # Indicadores de login exitoso
            login_exitoso = False
            
            if "login" not in current_url.lower() and "correspondencia.coordinador.cl" in current_url:
                print("✅ URL indica login exitoso")
                login_exitoso = True
            
            # Buscar elementos que aparecen después del login
            try:
                # Buscar elementos típicos de usuario logueado
                elementos_logueado = [
                    "cerrar sesión",
                    "logout",
                    "dashboard",
                    "menu",
                    "usuario"
                ]
                
                for elemento in elementos_logueado:
                    if elemento in page_source:
                        print(f"✅ Encontrado indicador de login: '{elemento}'")
                        login_exitoso = True
                        break
                        
            except Exception as e:
                print(f"⚠️ Error verificando elementos: {e}")
            
            if login_exitoso:
                print("✅ Login confirmado automáticamente!")
                self.is_logged_in = True
                return True
            else:
                print("⚠️ No se pudo verificar el login automáticamente")
                respuesta = input("❓ ¿Hiciste login correctamente? ¿Puedes ver el menú principal? (s/n): ")
                if respuesta.lower() in ['s', 'sí', 'si', 'y', 'yes']:
                    print("✅ Login confirmado manualmente!")
                    self.is_logged_in = True
                    return True
                else:
                    print("❌ Login no confirmado")
                    return False
                    
        except Exception as e:
            print(f"❌ Error en proceso de login: {e}")
            return False
    
    def extraer_enlaces_respondido_por(self, url_envio):
        """
        Extrae todos los enlaces de la sección "Respondido por"
        
        Args:
            url_envio (str): URL del envío principal (ej: DE01746-25)
        
        Returns:
            list: Lista de URLs de respuestas
        """
        if not self.is_logged_in:
            print("❌ Necesitas hacer login primero")
            return []
        
        try:
            print(f"🔍 Accediendo a: {url_envio}")
            self.driver.get(url_envio)
            time.sleep(3)
            
            # Buscar la sección "Respondido por"
            enlaces_respondido = []
            
            try:
                # Buscar el texto "Respondido por:" y extraer enlaces siguientes
                wait = WebDriverWait(self.driver, 10)
                
                # Método 1: Buscar por texto "Respondido por"
                try:
                    seccion_respondido = self.driver.find_element(
                        By.XPATH, "//dt[contains(text(), 'Respondido por:')]"
                    )
                    
                    # Encontrar el dd siguiente que contiene los enlaces
                    dd_enlaces = seccion_respondido.find_element(By.XPATH, "./following-sibling::dd[1]")
                    enlaces = dd_enlaces.find_elements(By.TAG_NAME, "a")
                    
                    for enlace in enlaces:
                        href = enlace.get_attribute('href')
                        texto = enlace.text.strip()
                        if href and 'correspondencia.coordinador.cl' in href:
                            enlaces_respondido.append({
                                'url': href,
                                'correlativo': texto,
                                'texto_enlace': enlace.get_attribute('outerHTML')
                            })
                            
                    print(f"✅ Método 1: Encontrados {len(enlaces_respondido)} enlaces")
                    
                except NoSuchElementException:
                    print("⚠️ Método 1 falló, probando método 2...")
                
                # Método 2: Buscar directamente enlaces que contengan correlativos
                if not enlaces_respondido:
                    todos_enlaces = self.driver.find_elements(By.TAG_NAME, "a")
                    
                    for enlace in todos_enlaces:
                        href = enlace.get_attribute('href')
                        texto = enlace.text.strip()
                        
                        # Filtrar enlaces que parecen correlativos y van a páginas de respuesta
                        if (href and 
                            'correspondencia.coordinador.cl' in href and
                            '/show/recibido/' in href and
                            re.match(r'DE\d{5}-\d{2}', texto)):
                            
                            enlaces_respondido.append({
                                'url': href,
                                'correlativo': texto,
                                'texto_enlace': enlace.get_attribute('outerHTML')
                            })
                    
                    # Eliminar duplicados
                    enlaces_unicos = {}
                    for enlace in enlaces_respondido:
                        if enlace['correlativo'] not in enlaces_unicos:
                            enlaces_unicos[enlace['correlativo']] = enlace
                    
                    enlaces_respondido = list(enlaces_unicos.values())
                    print(f"✅ Método 2: Encontrados {len(enlaces_respondido)} enlaces únicos")
                
                if enlaces_respondido:
                    print(f"📋 Enlaces encontrados:")
                    for i, enlace in enumerate(enlaces_respondido, 1):
                        print(f"  {i}. {enlace['correlativo']} -> {enlace['url']}")
                else:
                    print("❌ No se encontraron enlaces de 'Respondido por'")
                
                return enlaces_respondido
                
            except Exception as e:
                print(f"❌ Error extrayendo enlaces: {e}")
                return []
                
        except Exception as e:
            print(f"❌ Error accediendo a la página: {e}")
            return []
    
    def extraer_datos_respuesta_individual(self, url_respuesta, correlativo):
        """
        Extrae datos de una página individual de respuesta
        
        Args:
            url_respuesta (str): URL de la respuesta individual
            correlativo (str): Correlativo de la respuesta
        
        Returns:
            dict: Datos extraídos de la respuesta
        """
        try:
            print(f"  📄 Extrayendo: {correlativo}")
            self.driver.get(url_respuesta)
            time.sleep(2)
            
            # Extraer datos básicos (solo los que necesitamos)
            datos = {
                'correlativo': correlativo,
                'url': url_respuesta,
                'fecha_envio': '',
                'empresa': '',
                'responde_a': '',
                'anexos': '',
                'comtrade_respuesta': 'Sin información'
            }
            
            # Extraer fecha de envío y formatear (sin hora)
            try:
                fecha_element = self.driver.find_element(
                    By.XPATH, "//dt[contains(text(), 'Fecha Envío')]/following-sibling::dd[1]"
                )
                fecha_completa = fecha_element.text.strip()
                
                # Extraer solo la fecha (DD/MM/YYYY) sin la hora
                match_fecha = re.search(r'(\d{2}/\d{2}/\d{4})', fecha_completa)
                if match_fecha:
                    datos['fecha_envio'] = match_fecha.group(1)
                else:
                    datos['fecha_envio'] = fecha_completa.split(' ')[0] if fecha_completa else ''
                    
            except NoSuchElementException:
                pass
            
            # Extraer empresa
            try:
                empresa_element = self.driver.find_element(
                    By.XPATH, "//dt[contains(text(), 'Empresa')]/following-sibling::dd[1]"
                )
                empresa_texto = empresa_element.text.strip()
                # Limpiar texto de empresa (quitar saltos de línea extra)
                datos['empresa'] = ' '.join(empresa_texto.split())
            except NoSuchElementException:
                pass
            
            # Extraer "Responde a:" (en lugar de referencia)
            try:
                # Buscar la sección "Responde a:"
                responde_element = self.driver.find_element(
                    By.XPATH, "//dt[contains(text(), 'Responde a:')]/following-sibling::dd[1]"
                )
                
                # Buscar enlaces dentro de la sección "Responde a:"
                enlaces_responde = responde_element.find_elements(By.TAG_NAME, "a")
                
                correlativos_responde = []
                for enlace in enlaces_responde:
                    texto_enlace = enlace.text.strip()
                    # Buscar patrón de correlativo (ej: DE01746-25)
                    if re.match(r'[A-Z]{2}\d{5}-\d{2}', texto_enlace):
                        correlativos_responde.append(texto_enlace)
                
                # Si encontró correlativos en los enlaces, usarlos
                if correlativos_responde:
                    datos['responde_a'] = ' | '.join(correlativos_responde)
                else:
                    # Si no hay enlaces, buscar correlativos en el texto
                    texto_responde = responde_element.text.strip()
                    matches = re.findall(r'[A-Z]{2}\d{5}-\d{2}', texto_responde)
                    if matches:
                        datos['responde_a'] = ' | '.join(matches)
                    else:
                        datos['responde_a'] = texto_responde if texto_responde else ''
                        
            except NoSuchElementException:
                # Si no encuentra "Responde a:", buscar en otros lugares
                try:
                    # Buscar en toda la página correlativos que no sean el actual
                    page_text = self.driver.page_source
                    correlativos_encontrados = re.findall(r'[A-Z]{2}\d{5}-\d{2}', page_text)
                    
                    # Filtrar para excluir el correlativo actual
                    correlativos_diferentes = [c for c in correlativos_encontrados if c != correlativo]
                    
                    # Tomar el más común (probablemente el que responde)
                    if correlativos_diferentes:
                        from collections import Counter
                        mas_comun = Counter(correlativos_diferentes).most_common(1)[0][0]
                        datos['responde_a'] = mas_comun
                        
                except:
                    pass
            
            # Extraer información de anexos (clave para Comtrade)
            try:
                # Buscar sección de anexos
                anexos_elements = self.driver.find_elements(
                    By.XPATH, "//dt[contains(text(), 'Anexos')]/following-sibling::dd[1]//a"
                )
                
                anexos_texto = []
                for anexo in anexos_elements:
                    anexo_texto = anexo.text.strip()
                    # Limpiar texto del anexo (quitar "Descargar" si aparece)
                    anexo_limpio = anexo_texto.replace('Descargar', '').replace('Anexo', '').strip()
                    if anexo_limpio:
                        anexos_texto.append(anexo_limpio)
                
                datos['anexos'] = ' | '.join(anexos_texto)
                
                # Determinar respuesta Comtrade basada en anexos
                datos['comtrade_respuesta'] = self.determinar_comtrade_desde_anexos(
                    datos['anexos']
                )
                
            except NoSuchElementException:
                pass
            
            return datos
            
        except Exception as e:
            print(f"  ❌ Error extrayendo {correlativo}: {e}")
            return {
                'correlativo': correlativo,
                'url': url_respuesta,
                'fecha_envio': '',
                'empresa': 'ERROR AL EXTRAER',
                'responde_a': '',
                'anexos': '',
                'comtrade_respuesta': 'Error al extraer'
            }
    
    def determinar_comtrade_desde_anexos(self, anexos_texto):
        """
        Determina la respuesta de Comtrade basándose en anexos
        
        Args:
            anexos_texto (str): Texto de los anexos
        
        Returns:
            str: Estado de respuesta Comtrade
        """
        # Solo usar anexos para determinar Comtrade
        texto_completo = anexos_texto.upper()
        
        # Patrones que indican envío de Comtrade
        patrones_si_envia = [
            'COMTRADE',
            '.RAR',
            '.ZIP',
            'ARCHIVO',
            'DATOS',
            'REGISTROS'
        ]
        
        # Patrones que indican NO envío
        patrones_no_envia = [
            'SIN ANEXO',
            'NO ANEXO',
            'NO APLICA',
            'NO CORRESPONDE',
            'INFORME DE FALLA'
        ]
        
        # Verificar patrones
        if any(patron in texto_completo for patron in patrones_si_envia):
            return 'Sí envía'
        elif any(patron in texto_completo for patron in patrones_no_envia):
            return 'No envía'
        elif not anexos_texto or anexos_texto.strip() == '':
            return 'Sin anexos'
        else:
            return 'Sin información'
    
    def procesar_envio_completo(self, url_envio_principal=None, limite_enlaces=None):
        """
        Procesa un envío completo: navega a la URL y extrae todas las respuestas
        
        Args:
            url_envio_principal (str): URL del envío principal (opcional)
            limite_enlaces (int): Cantidad máxima de enlaces a procesar (opcional)
        
        Returns:
            list: Lista con todos los datos extraídos
        """
        print(f"\n🚀 PROCESANDO ENVÍO COMPLETO")
        print("="*70)
        
        # Si no se proporciona URL, pedirla
        if not url_envio_principal:
            print(f"📍 Ingresa la URL del envío principal:")
            print(f"Ejemplo: https://correspondencia.coordinador.cl/correspondencia/show/envio/67df373635635739b703ab3f")
            url_envio_principal = input("URL del envío: ").strip()
            
            if not url_envio_principal:
                # URL por defecto
                url_envio_principal = "https://correspondencia.coordinador.cl/correspondencia/show/envio/67df373635635739b703ab3f"
                print(f"✅ Usando URL por defecto: {url_envio_principal}")
        
        # Preguntar por límite si no se especifica
        if limite_enlaces is None:
            print(f"\n📊 ¿Quieres limitar la cantidad de enlaces para pruebas?")
            respuesta = input("Ingresa número máximo de enlaces (o ENTER para todos): ").strip()
            if respuesta.isdigit():
                limite_enlaces = int(respuesta)
                print(f"✅ Limitando a {limite_enlaces} enlaces")
            else:
                print("✅ Procesando todos los enlaces encontrados")
        
        print(f"📍 URL a procesar: {url_envio_principal}")
        if limite_enlaces:
            print(f"🔢 Límite de enlaces: {limite_enlaces}")
        
        # Verificar que estamos logueados
        if not self.is_logged_in:
            print("❌ Necesitas hacer login primero")
            return []
        
        # Paso 1: Navegar a la URL del envío principal
        print(f"\n🔄 Navegando a la página del envío...")
        try:
            self.driver.get(url_envio_principal)
            time.sleep(4)  # Esperar a que cargue completamente
            
            current_url = self.driver.current_url
            if "login" in current_url.lower():
                print("❌ La sesión expiró. Necesitas hacer login nuevamente.")
                return []
            
            print(f"✅ Página del envío cargada: {current_url}")
            
        except Exception as e:
            print(f"❌ Error navegando a la URL: {e}")
            return []
        
        # Paso 2: Extraer enlaces de "Respondido por"
        enlaces = self.extraer_enlaces_respondido_por(url_envio_principal)
        
        if not enlaces:
            print("❌ No se encontraron enlaces para procesar")
            print("🔍 Verificando si la página tiene la estructura esperada...")
            
            # Debug: mostrar información de la página actual
            try:
                title = self.driver.title
                print(f"📄 Título de página: {title}")
                
                # Buscar texto "Respondido por" en toda la página
                page_text = self.driver.page_source.lower()
                if "respondido por" in page_text:
                    print("✅ Se encontró texto 'Respondido por' en la página")
                else:
                    print("❌ No se encontró texto 'Respondido por' en la página")
                
                # Contar enlaces totales
                all_links = self.driver.find_elements(By.TAG_NAME, "a")
                print(f"🔗 Total de enlaces en la página: {len(all_links)}")
                
            except Exception as debug_error:
                print(f"⚠️ Error en debug: {debug_error}")
            
            return []
        
        # Paso 3: Aplicar límite si se especifica
        total_encontrados = len(enlaces)
        if limite_enlaces and limite_enlaces < total_encontrados:
            enlaces = enlaces[:limite_enlaces]
            print(f"🔢 Limitando de {total_encontrados} a {limite_enlaces} enlaces para pruebas")
        
        # Paso 4: Procesar cada enlace individualmente
        todos_resultados = []
        total_enlaces = len(enlaces)
        
        print(f"\n📊 Procesando {total_enlaces} respuestas...")
        print("="*50)
        
        for i, enlace_info in enumerate(enlaces, 1):
            print(f"\n[{i}/{total_enlaces}] Procesando: {enlace_info['correlativo']}")
            
            try:
                datos = self.extraer_datos_respuesta_individual(
                    enlace_info['url'], 
                    enlace_info['correlativo']
                )
                
                # Agregar información del envío principal
                datos['envio_principal'] = url_envio_principal
                datos['numero_respuesta'] = i
                datos['total_encontrados'] = total_encontrados
                datos['limite_aplicado'] = limite_enlaces if limite_enlaces else 'Sin límite'
                
                todos_resultados.append(datos)
                
                print(f"  ✅ {datos['correlativo']} - {datos['empresa']} - {datos['comtrade_respuesta']}")
                
            except Exception as e:
                print(f"  ❌ Error procesando {enlace_info['correlativo']}: {e}")
                # Agregar resultado con error para no perder la info
                datos_error = {
                    'correlativo': enlace_info['correlativo'],
                    'url': enlace_info['url'],
                    'envio_principal': url_envio_principal,
                    'numero_respuesta': i,
                    'fecha_envio': 'ERROR',
                    'empresa': 'ERROR AL EXTRAER',
                    'responde_a': '',
                    'anexos': '',
                    'comtrade_respuesta': 'Error al extraer',
                    'total_encontrados': total_encontrados,
                    'limite_aplicado': limite_enlaces if limite_enlaces else 'Sin límite'
                }
                todos_resultados.append(datos_error)
            
            # Pausa entre solicitudes para no sobrecargar el servidor
            time.sleep(1.5)
        
        print(f"\n🎉 Procesamiento completado!")
        print(f"📊 Enlaces encontrados: {total_encontrados}")
        print(f"📊 Enlaces procesados: {len(todos_resultados)}")
        if limite_enlaces:
            print(f"🔢 Límite aplicado: {limite_enlaces}")
        
        return todos_resultados
    
    def exportar_resultados(self, resultados, nombre_archivo="respuestas_extraidas.xlsx"):
        """Exporta los resultados a Excel con formato específico - SOLO COLUMNAS NECESARIAS"""
        try:
            if not resultados:
                print("❌ No hay resultados para exportar")
                # Crear archivo vacío con headers simplificados
                df_vacio = pd.DataFrame(columns=[
                    'Envío Principal', 'Número Respuesta', 'Correlativo Respuesta',
                    'Fecha de Envío', 'Empresa', 'Responde a', 'Anexos', 'Envía Comtrade'
                ])
                df_vacio.to_excel(nombre_archivo, index=False, engine='openpyxl')
                print(f"📄 Archivo Excel vacío creado: {nombre_archivo}")
                return
            
            # Crear DataFrame
            df = pd.DataFrame(resultados)
            
            # Definir SOLO las columnas que necesitas (con Responde a en lugar de Referencia)
            columnas_finales = {
                'envio_principal': 'Envío Principal',
                'numero_respuesta': 'Número Respuesta', 
                'correlativo': 'Correlativo Respuesta',
                'fecha_envio': 'Fecha de Envío',
                'empresa': 'Empresa',
                'responde_a': 'Responde a',
                'anexos': 'Anexos',
                'comtrade_respuesta': 'Envía Comtrade'
            }
            
            # Crear DataFrame final SOLO con las columnas necesarias
            df_final = pd.DataFrame()
            for col_orig, col_nuevo in columnas_finales.items():
                if col_orig in df.columns:
                    df_final[col_nuevo] = df[col_orig]
                else:
                    df_final[col_nuevo] = ''
            
            # Limpiar datos vacíos
            df_final = df_final.fillna('')
            
            # Formatear fechas: asegurar que solo sea DD/MM/YYYY
            if 'Fecha de Envío' in df_final.columns:
                def limpiar_fecha(fecha_str):
                    if not fecha_str or fecha_str == '':
                        return ''
                    
                    # Buscar patrón DD/MM/YYYY en el texto
                    match = re.search(r'(\d{2}/\d{2}/\d{4})', str(fecha_str))
                    if match:
                        return match.group(1)
                    else:
                        # Si no encuentra el patrón, tomar solo la primera parte antes del espacio
                        return str(fecha_str).split(' ')[0] if fecha_str else ''
                
                df_final['Fecha de Envío'] = df_final['Fecha de Envío'].apply(limpiar_fecha)
            
            # Limpiar anexos: quitar texto extra como "Descargar"
            if 'Anexos' in df_final.columns:
                def limpiar_anexos(anexo_str):
                    if not anexo_str or anexo_str == '':
                        return ''
                    
                    # Limpiar texto común de los anexos
                    anexo_limpio = str(anexo_str)
                    anexo_limpio = anexo_limpio.replace('Descargar', '')
                    anexo_limpio = anexo_limpio.replace('Anexo', '')
                    anexo_limpio = anexo_limpio.replace('  ', ' ')  # Dobles espacios
                    return anexo_limpio.strip()
                
                df_final['Anexos'] = df_final['Anexos'].apply(limpiar_anexos)
            
            # Limpiar empresa: quitar saltos de línea extra
            if 'Empresa' in df_final.columns:
                def limpiar_empresa(empresa_str):
                    if not empresa_str or empresa_str == '':
                        return ''
                    
                    # Limpiar saltos de línea y espacios extra
                    empresa_limpia = ' '.join(str(empresa_str).split())
                    return empresa_limpia
                
                df_final['Empresa'] = df_final['Empresa'].apply(limpiar_empresa)
            
            # Limpiar "Responde a": asegurar que solo tenga correlativos limpios
            if 'Responde a' in df_final.columns:
                def limpiar_responde_a(responde_str):
                    if not responde_str or responde_str == '':
                        return ''
                    
                    # Buscar solo correlativos válidos en el texto
                    correlativos = re.findall(r'[A-Z]{2}\d{5}-\d{2}', str(responde_str))
                    
                    if correlativos:
                        # Eliminar duplicados manteniendo el orden
                        correlativos_unicos = []
                        for c in correlativos:
                            if c not in correlativos_unicos:
                                correlativos_unicos.append(c)
                        return ' | '.join(correlativos_unicos)
                    else:
                        return str(responde_str).strip()
                
                df_final['Responde a'] = df_final['Responde a'].apply(limpiar_responde_a)
            
            # Exportar a Excel
            df_final.to_excel(nombre_archivo, index=False, engine='openpyxl')
            
            print(f"\n✅ EXCEL CON 'RESPONDE A' EXPORTADO!")
            print(f"📁 Archivo: {nombre_archivo}")
            print(f"📊 Total de respuestas: {len(df_final)}")
            print(f"📋 Columnas: {len(df_final.columns)} (Referencia → Responde a)")
            
            # Estadísticas simplificadas
            if len(df_final) > 0:
                print(f"\n📈 ESTADÍSTICAS:")
                print("="*40)
                
                # Estadísticas de Comtrade
                if 'Envía Comtrade' in df_final.columns:
                    comtrade_stats = df_final['Envía Comtrade'].value_counts()
                    print(f"📊 Respuestas Comtrade:")
                    for respuesta, count in comtrade_stats.items():
                        print(f"   • {respuesta}: {count}")
                
                # Estadísticas de "Responde a"
                if 'Responde a' in df_final.columns:
                    responde_stats = df_final['Responde a'].value_counts()
                    print(f"\n📝 Cartas a las que responden:")
                    for responde, count in responde_stats.head(5).items():
                        if responde and responde != '':
                            print(f"   • {responde}: {count} respuestas")
                
                # Empresas únicas
                if 'Empresa' in df_final.columns:
                    empresas_no_error = df_final[df_final['Empresa'] != 'ERROR AL EXTRAER']['Empresa'].nunique()
                    print(f"\n🏢 Empresas únicas: {empresas_no_error}")
                
                # Preview de los primeros resultados (formato más limpio)
                print(f"\n📋 PREVIEW (Primeros 3 resultados):")
                print("="*50)
                for index, row in df_final.head(3).iterrows():
                    correlativo = row.get('Correlativo Respuesta', 'N/A')
                    empresa = row.get('Empresa', 'N/A')
                    comtrade = row.get('Envía Comtrade', 'N/A')
                    fecha = row.get('Fecha de Envío', 'N/A')
                    responde_a = row.get('Responde a', 'N/A')
                    anexos = row.get('Anexos', 'N/A')
                    
                    print(f"{index+1}. {correlativo} | {fecha}")
                    print(f"   🏢 {empresa}")
                    print(f"   ↪️  Responde a: {responde_a}")
                    print(f"   📎 {anexos}")
                    print(f"   📊 {comtrade}")
                    print()
                
                if len(df_final) > 3:
                    print(f"... y {len(df_final) - 3} resultados más en el Excel")
            
            print(f"\n🎯 ¡Excel con columna 'Responde a' listo!")
            print(f"📋 Cambio: Referencia → Responde a (correlativos específicos)")
            
        except Exception as e:
            print(f"❌ Error exportando: {e}")
            # Intentar exportación de emergencia simplificada
            try:
                df_backup = pd.DataFrame(resultados) if resultados else pd.DataFrame()
                # Solo columnas básicas en backup
                columnas_backup = ['correlativo', 'fecha_envio', 'empresa', 'responde_a', 'anexos', 'comtrade_respuesta']
                df_backup_simple = df_backup[columnas_backup] if not df_backup.empty else df_backup
                
                nombre_backup = f"backup_responde_a_{nombre_archivo}"
                df_backup_simple.to_excel(nombre_backup, index=False)
                print(f"💾 Exportación de emergencia con 'Responde a': {nombre_backup}")
            except Exception as backup_error:
                print(f"❌ Error en exportación de emergencia: {backup_error}")
    
    def cerrar(self):
        """Cierra el driver"""
        if self.driver:
            self.driver.quit()

# Función principal
if __name__ == "__main__":
    extractor = ExtractorRespondidoPor(headless=False)
    
    try:
        print("🔗 EXTRACTOR DE ENLACES 'RESPONDIDO POR'")
        print("="*50)
        print("Este código:")
        print("1. Hace login manual")
        print("2. Va a la página del envío principal")
        print("3. Extrae todos los enlaces de 'Respondido por'")
        print("4. Visita cada enlace y extrae datos completos")
        print("5. Determina si envía Comtrade basado en anexos")
        print("6. Exporta todo a Excel con columna 'Responde a'")
        
        # Login
        login_ok = extractor.login_interactivo()
        
        if login_ok:
            print(f"\n🎯 PASO 4: NAVEGAR A LA PÁGINA DEL ENVÍO")
            print("="*50)
            
            # Opción rápida para pruebas
            print(f"🧪 MODO DE PRUEBA DISPONIBLE:")
            print(f"Para hacer pruebas rápidas, puedes limitar la cantidad de enlaces a procesar")
            
            # Procesar el envío completo (incluye navegación y límite)
            resultados = extractor.procesar_envio_completo()
            
            # SIEMPRE generar Excel, incluso si no hay resultados
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            nombre_archivo = f"respuestas_comtrade_{timestamp}.xlsx"
            
            print(f"\n📊 GENERANDO EXCEL...")
            extractor.exportar_resultados(resultados, nombre_archivo)
            
            if resultados:
                print(f"\n🎯 RESUMEN FINAL:")
                print(f"📁 Archivo generado: {nombre_archivo}")
                print(f"📊 Total respuestas procesadas: {len(resultados)}")
                
                # Mostrar preview de resultados
                print(f"\n📋 Preview de resultados:")
                for resultado in resultados[:3]:  # Mostrar primeros 3
                    print(f"  • {resultado['correlativo']} - {resultado['empresa']} - {resultado['comtrade_respuesta']}")
                
                if len(resultados) > 3:
                    print(f"  ... y {len(resultados) - 3} más")
                    
                # Estadísticas rápidas
                comtrade_si = len([r for r in resultados if 'sí' in r.get('comtrade_respuesta', '').lower()])
                comtrade_no = len([r for r in resultados if 'no' in r.get('comtrade_respuesta', '').lower()])
                print(f"\n📈 Resumen rápido:")
                print(f"   🟢 Envían Comtrade: {comtrade_si}")
                print(f"   🔴 No envían Comtrade: {comtrade_no}")
                print(f"   📊 Otros: {len(resultados) - comtrade_si - comtrade_no}")
                
            else:
                print(f"\n⚠️ No se extrajeron resultados, pero se generó archivo Excel vacío")
                print(f"📁 Archivo: {nombre_archivo}")
                print("\n🔍 Posibles causas:")
                print("1. La URL del envío no tiene enlaces 'Respondido por'")
                print("2. Los enlaces tienen una estructura diferente")
                print("3. Problemas de permisos en la página")
                print("4. La página no cargó completamente")
            
            input("\n⏳ Presiona ENTER para cerrar el navegador...")
        
        else:
            print("❌ No se pudo hacer login")
    
    except Exception as e:
        print(f"❌ Error en la ejecución: {e}")
        import traceback
        traceback.print_exc()
    
    finally:
        extractor.cerrar()