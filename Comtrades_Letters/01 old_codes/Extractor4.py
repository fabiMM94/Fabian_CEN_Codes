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
import os
import shutil
import urllib.parse
import requests

class ExtractorRespondidoPorConDescargas:
    def __init__(self, headless=False):
        """
        Extractor especializado para seguir enlaces de "Respondido por" con descarga de documentos y anexos
        """
        self.driver = None
        self.is_logged_in = False
        self.carpeta_cartas = "Cartas"
        self.carpeta_anexos = "Anexos"
        self.session = requests.Session()  # Para mantener cookies de selenium en requests
        self.setup_carpetas()
        self.setup_driver(headless)
    
    def setup_carpetas(self):
        """
        Crear y limpiar carpetas de descarga
        """
        print("\n📁 CONFIGURANDO CARPETAS DE DESCARGA")
        print("="*50)
        
        # Limpiar y crear carpeta de Cartas
        if os.path.exists(self.carpeta_cartas):
            print(f"🗑️ Limpiando carpeta existente: {self.carpeta_cartas}")
            shutil.rmtree(self.carpeta_cartas)
        
        os.makedirs(self.carpeta_cartas, exist_ok=True)
        print(f"✅ Carpeta creada: {self.carpeta_cartas}")
        
        # Limpiar y crear carpeta de Anexos
        if os.path.exists(self.carpeta_anexos):
            print(f"🗑️ Limpiando carpeta existente: {self.carpeta_anexos}")
            shutil.rmtree(self.carpeta_anexos)
        
        os.makedirs(self.carpeta_anexos, exist_ok=True)
        print(f"✅ Carpeta creada: {self.carpeta_anexos}")
        
        print("📁 Carpetas de descarga listas y limpias")
        
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
        
        # Configurar descarga automática
        download_path = os.path.abspath("temp_downloads")
        prefs = {
            "download.default_directory": download_path,
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
            "safebrowsing.enabled": True
        }
        chrome_options.add_experimental_option("prefs", prefs)
        
        # Crear carpeta temporal de descargas
        os.makedirs(download_path, exist_ok=True)
        
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
            
            # Sincronizar cookies de Selenium con requests
            self.sincronizar_cookies()
            
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
                    self.sincronizar_cookies()
                    return True
                else:
                    print("❌ Login no confirmado")
                    return False
                    
        except Exception as e:
            print(f"❌ Error en proceso de login: {e}")
            return False
    
    def sincronizar_cookies(self):
        """
        Sincroniza las cookies de Selenium con la sesión de requests
        """
        try:
            selenium_cookies = self.driver.get_cookies()
            for cookie in selenium_cookies:
                self.session.cookies.set(cookie['name'], cookie['value'])
            print("🍪 Cookies sincronizadas para descargas")
        except Exception as e:
            print(f"⚠️ Error sincronizando cookies: {e}")
    
    def descargar_documento_principal(self, correlativo):
        """
        Descarga el documento principal de la carta
        """
        try:
            print(f"  📄 Buscando documento principal para {correlativo}...")
            
            # Buscar el enlace de descarga del documento principal
            download_links = self.driver.find_elements(
                By.XPATH, 
                "//a[contains(@href, '/correspondencia/download_saved_file/') and contains(text(), 'Descargar Documento')]"
            )
            
            if not download_links:
                # Método alternativo: buscar por ID
                download_links = self.driver.find_elements(By.ID, "download_file")
            
            if download_links:
                download_url = download_links[0].get_attribute('href')
                print(f"  📥 Descargando documento principal: {download_url}")
                
                # Descargar usando requests con las cookies de selenium
                response = self.session.get(download_url)
                
                if response.status_code == 200:
                    # Determinar nombre del archivo
                    filename = f"{correlativo}_documento_principal.pdf"
                    
                    # Intentar obtener nombre del header Content-Disposition
                    if 'content-disposition' in response.headers:
                        import re
                        cd = response.headers['content-disposition']
                        filename_match = re.findall('filename="(.+)"', cd)
                        if filename_match:
                            filename = f"{correlativo}_{filename_match[0]}"
                    
                    # Guardar archivo en carpeta Cartas
                    filepath = os.path.join(self.carpeta_cartas, filename)
                    with open(filepath, 'wb') as f:
                        f.write(response.content)
                    
                    print(f"  ✅ Documento guardado: {filename}")
                    return filename
                else:
                    print(f"  ❌ Error descargando documento: HTTP {response.status_code}")
                    return None
            else:
                print(f"  ⚠️ No se encontró enlace de descarga del documento principal")
                return None
                
        except Exception as e:
            print(f"  ❌ Error descargando documento principal: {e}")
            return None
    
    def descargar_anexos(self, correlativo):
        """
        Descarga todos los anexos disponibles
        """
        anexos_descargados = []
        
        try:
            print(f"  📎 Buscando anexos para {correlativo}...")
            
            # Buscar enlaces de anexos
            anexo_links = self.driver.find_elements(
                By.XPATH, 
                "//a[contains(@href, '/cartas/download_anexos/') or contains(text(), 'Descargar Anexo')]"
            )
            
            if not anexo_links:
                print(f"  ⚠️ No se encontraron anexos para {correlativo}")
                return anexos_descargados
            
            print(f"  📎 Encontrados {len(anexo_links)} anexos")
            
            for i, anexo_link in enumerate(anexo_links, 1):
                try:
                    anexo_url = anexo_link.get_attribute('href')
                    anexo_text = anexo_link.text.strip()
                    
                    print(f"    📥 Descargando anexo {i}: {anexo_text}")
                    
                    # Descargar usando requests
                    response = self.session.get(anexo_url)
                    
                    if response.status_code == 200:
                        # Limpiar nombre del anexo
                        nombre_anexo = anexo_text.replace('Descargar Anexo', '').strip()
                        if not nombre_anexo:
                            nombre_anexo = f"anexo_{i}"
                        
                        # Asegurar extensión
                        if not any(nombre_anexo.lower().endswith(ext) for ext in ['.pdf', '.rar', '.zip', '.doc', '.docx', '.txt']):
                            # Intentar detectar tipo de archivo por content-type
                            content_type = response.headers.get('content-type', '').lower()
                            if 'pdf' in content_type:
                                nombre_anexo += '.pdf'
                            elif 'zip' in content_type or 'rar' in content_type:
                                nombre_anexo += '.rar'
                            else:
                                nombre_anexo += '.pdf'  # Por defecto
                        
                        # Nombre final con correlativo
                        filename = f"{correlativo}_{nombre_anexo}"
                        
                        # Guardar en carpeta Anexos
                        filepath = os.path.join(self.carpeta_anexos, filename)
                        with open(filepath, 'wb') as f:
                            f.write(response.content)
                        
                        anexos_descargados.append(filename)
                        print(f"    ✅ Anexo guardado: {filename}")
                    else:
                        print(f"    ❌ Error descargando anexo {i}: HTTP {response.status_code}")
                        
                except Exception as anexo_error:
                    print(f"    ❌ Error con anexo {i}: {anexo_error}")
                    continue
            
            if anexos_descargados:
                print(f"  ✅ Total anexos descargados: {len(anexos_descargados)}")
            
            return anexos_descargados
            
        except Exception as e:
            print(f"  ❌ Error descargando anexos: {e}")
            return anexos_descargados
    
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
        Extrae datos de una página individual de respuesta Y DESCARGA documentos/anexos
        
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
            
            # **NUEVA FUNCIONALIDAD: DESCARGAR DOCUMENTOS Y ANEXOS**
            print(f"  💾 Iniciando descargas para {correlativo}...")
            
            # Descargar documento principal
            documento_descargado = self.descargar_documento_principal(correlativo)
            
            # Descargar anexos
            anexos_descargados = self.descargar_anexos(correlativo)
            
            # Extraer datos básicos (como antes)
            datos = {
                'correlativo': correlativo,
                'url': url_respuesta,
                'fecha_envio': '',
                'empresa': '',
                'responde_a': '',
                'anexos': '',
                'comtrade_respuesta': 'Sin información',
                'documento_descargado': documento_descargado,  # NUEVO
                'anexos_descargados': ' | '.join(anexos_descargados) if anexos_descargados else 'Sin anexos'  # NUEVO
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
                    # Limpiar texto del anexo de forma más inteligente
                    # Solo quitar "Descargar" y "Anexo" si están al inicio
                    anexo_limpio = anexo_texto
                    
                    # Quitar "Descargar" si está al inicio
                    if anexo_limpio.startswith('Descargar'):
                        anexo_limpio = anexo_limpio[9:].strip()  # Quitar "Descargar"
                    
                    # Solo quitar "Anexo" si está al inicio y seguido de espacio
                    if anexo_limpio.startswith('Anexo '):
                        anexo_limpio = anexo_limpio[6:].strip()  # Quitar "Anexo "
                    
                    if anexo_limpio:
                        anexos_texto.append(anexo_limpio)
                
                datos['anexos'] = ' | '.join(anexos_texto)
                
                # Determinar respuesta Comtrade basada en anexos
                datos['comtrade_respuesta'] = self.determinar_comtrade_desde_anexos(
                    datos['anexos']
                )
                
            except NoSuchElementException:
                pass
            
            print(f"  ✅ {correlativo} procesado - Doc: {'✅' if documento_descargado else '❌'} - Anexos: {len(anexos_descargados)}")
            
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
                'comtrade_respuesta': 'Error al extraer',
                'documento_descargado': None,
                'anexos_descargados': 'Error'
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
        Procesa un envío completo: navega a la URL y extrae todas las respuestas CON DESCARGAS
        
        Args:
            url_envio_principal (str): URL del envío principal (opcional)
            limite_enlaces (int): Cantidad máxima de enlaces a procesar (opcional)
        
        Returns:
            list: Lista con todos los datos extraídos
        """
        print(f"\n🚀 PROCESANDO ENVÍO COMPLETO CON DESCARGAS")
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
        
        # **NUEVA FUNCIONALIDAD: PREPARAR CARPETAS DE DESCARGA**
        print(f"\n💾 PREPARANDO SISTEMA DE DESCARGAS")
        print("="*50)
        print(f"📁 Carpeta de documentos: {self.carpeta_cartas}")
        print(f"📎 Carpeta de anexos: {self.carpeta_anexos}")
        print("🔄 Las carpetas se limpiarán automáticamente antes de procesar")
        
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
        
        # Paso 4: Procesar cada enlace individualmente CON DESCARGAS
        todos_resultados = []
        total_enlaces = len(enlaces)
        
        print(f"\n📊 Procesando {total_enlaces} respuestas CON DESCARGAS...")
        print("="*50)
        print("💾 Cada respuesta descargará:")
        print("   📄 Documento principal → Carpeta 'Cartas'")
        print("   📎 Anexos → Carpeta 'Anexos'")
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
                    'documento_descargado': None,
                    'anexos_descargados': 'Error',
                    'total_encontrados': total_encontrados,
                    'limite_aplicado': limite_enlaces if limite_enlaces else 'Sin límite'
                }
                todos_resultados.append(datos_error)
            
            # Pausa entre solicitudes para no sobrecargar el servidor
            time.sleep(1.5)
        
        # **REPORTE FINAL DE DESCARGAS**
        print(f"\n🎉 PROCESAMIENTO Y DESCARGAS COMPLETADOS!")
        print("="*70)
        print(f"📊 Enlaces encontrados: {total_encontrados}")
        print(f"📊 Enlaces procesados: {len(todos_resultados)}")
        if limite_enlaces:
            print(f"🔢 Límite aplicado: {limite_enlaces}")
        
        # Contar archivos descargados
        try:
            cartas_count = len([f for f in os.listdir(self.carpeta_cartas) if os.path.isfile(os.path.join(self.carpeta_cartas, f))])
            anexos_count = len([f for f in os.listdir(self.carpeta_anexos) if os.path.isfile(os.path.join(self.carpeta_anexos, f))])
            
            print(f"\n💾 RESUMEN DE DESCARGAS:")
            print(f"📄 Documentos en carpeta 'Cartas': {cartas_count}")
            print(f"📎 Anexos en carpeta 'Anexos': {anexos_count}")
            print(f"📁 Total archivos descargados: {cartas_count + anexos_count}")
            
        except Exception as count_error:
            print(f"⚠️ Error contando archivos descargados: {count_error}")
        
        return todos_resultados
    
    def exportar_resultados(self, resultados, nombre_archivo="respuestas_extraidas_con_descargas.xlsx"):
        """Exporta los resultados a Excel con formato específico - INCLUYENDO información de descargas"""
        try:
            if not resultados:
                print("❌ No hay resultados para exportar")
                # Crear archivo vacío con headers que incluyen las nuevas columnas
                df_vacio = pd.DataFrame(columns=[
                    'Envío Principal', 'Número Respuesta', 'Correlativo Respuesta',
                    'Fecha de Envío', 'Empresa', 'Responde a', 'Anexos', 'Envía Comtrade',
                    'Documento Descargado', 'Anexos Descargados'  # NUEVAS COLUMNAS
                ])
                df_vacio.to_excel(nombre_archivo, index=False, engine='openpyxl')
                print(f"📄 Archivo Excel vacío creado: {nombre_archivo}")
                return
            
            # Crear DataFrame
            df = pd.DataFrame(resultados)
            
            # Definir columnas finales INCLUYENDO las nuevas de descarga
            columnas_finales = {
                'envio_principal': 'Envío Principal',
                'numero_respuesta': 'Número Respuesta', 
                'correlativo': 'Correlativo Respuesta',
                'fecha_envio': 'Fecha de Envío',
                'empresa': 'Empresa',
                'responde_a': 'Responde a',
                'anexos': 'Anexos',
                'comtrade_respuesta': 'Envía Comtrade',
                'documento_descargado': 'Documento Descargado',  # NUEVA
                'anexos_descargados': 'Anexos Descargados'      # NUEVA
            }
            
            # Crear DataFrame final con TODAS las columnas
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
                    
                    # Limpiar texto común de los anexos de forma más inteligente
                    anexo_limpio = str(anexo_str)
                    
                    # Quitar "Descargar" al inicio o en cualquier parte
                    anexo_limpio = anexo_limpio.replace('Descargar', '').strip()
                    
                    # Solo quitar "Anexo" si está al inicio y seguido de un espacio
                    if anexo_limpio.startswith('Anexo '):
                        anexo_limpio = anexo_limpio[6:]  # Quitar "Anexo " (6 caracteres)
                    
                    # Limpiar espacios dobles y al inicio/final
                    anexo_limpio = ' '.join(anexo_limpio.split())
                    
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
            
            # **NUEVA FUNCIONALIDAD: Formatear columnas de descarga**
            if 'Documento Descargado' in df_final.columns:
                def formatear_documento_descargado(doc_str):
                    if not doc_str or doc_str == '' or str(doc_str).lower() == 'none':
                        return 'No descargado'
                    else:
                        return str(doc_str)
                
                df_final['Documento Descargado'] = df_final['Documento Descargado'].apply(formatear_documento_descargado)
            
            if 'Anexos Descargados' in df_final.columns:
                def formatear_anexos_descargados(anexos_str):
                    if not anexos_str or anexos_str == '' or str(anexos_str).lower() in ['none', 'sin anexos', 'error']:
                        return 'Sin anexos descargados'
                    else:
                        return str(anexos_str)
                
                df_final['Anexos Descargados'] = df_final['Anexos Descargados'].apply(formatear_anexos_descargados)
            
            # Exportar a Excel
            df_final.to_excel(nombre_archivo, index=False, engine='openpyxl')
            
            print(f"\n✅ EXCEL CON INFORMACIÓN DE DESCARGAS EXPORTADO!")
            print(f"📁 Archivo: {nombre_archivo}")
            print(f"📊 Total de respuestas: {len(df_final)}")
            print(f"📋 Columnas: {len(df_final.columns)} (incluye info de descargas)")
            
            # Estadísticas extendidas con información de descargas
            if len(df_final) > 0:
                print(f"\n📈 ESTADÍSTICAS COMPLETAS:")
                print("="*50)
                
                # Estadísticas de Comtrade
                if 'Envía Comtrade' in df_final.columns:
                    comtrade_stats = df_final['Envía Comtrade'].value_counts()
                    print(f"📊 Respuestas Comtrade:")
                    for respuesta, count in comtrade_stats.items():
                        print(f"   • {respuesta}: {count}")
                
                # **NUEVAS ESTADÍSTICAS DE DESCARGAS**
                if 'Documento Descargado' in df_final.columns:
                    docs_descargados = len([d for d in df_final['Documento Descargado'] if d != 'No descargado'])
                    print(f"\n💾 Documentos descargados: {docs_descargados}/{len(df_final)}")
                
                if 'Anexos Descargados' in df_final.columns:
                    anexos_descargados = len([a for a in df_final['Anexos Descargados'] if a != 'Sin anexos descargados'])
                    print(f"📎 Respuestas con anexos descargados: {anexos_descargados}/{len(df_final)}")
                
                # Estadísticas de "Responde a"
                if 'Responde a' in df_final.columns:
                    responde_stats = df_final['Responde a'].value_counts()
                    print(f"\n📝 Cartas a las que responden (Top 5):")
                    for responde, count in responde_stats.head(5).items():
                        if responde and responde != '':
                            print(f"   • {responde}: {count} respuestas")
                
                # Empresas únicas
                if 'Empresa' in df_final.columns:
                    empresas_no_error = df_final[df_final['Empresa'] != 'ERROR AL EXTRAER']['Empresa'].nunique()
                    print(f"\n🏢 Empresas únicas: {empresas_no_error}")
                
                # **PREVIEW MEJORADO CON INFORMACIÓN DE DESCARGAS**
                print(f"\n📋 PREVIEW CON DESCARGAS (Primeros 3 resultados):")
                print("="*60)
                for index, row in df_final.head(3).iterrows():
                    correlativo = row.get('Correlativo Respuesta', 'N/A')
                    empresa = row.get('Empresa', 'N/A')
                    comtrade = row.get('Envía Comtrade', 'N/A')
                    fecha = row.get('Fecha de Envío', 'N/A')
                    responde_a = row.get('Responde a', 'N/A')
                    doc_descargado = row.get('Documento Descargado', 'N/A')
                    anexos_descargados = row.get('Anexos Descargados', 'N/A')
                    
                    print(f"{index+1}. {correlativo} | {fecha}")
                    print(f"   🏢 {empresa}")
                    print(f"   ↪️  Responde a: {responde_a}")
                    print(f"   📄 Documento: {doc_descargado}")
                    print(f"   📎 Anexos: {anexos_descargados}")
                    print(f"   📊 {comtrade}")
                    print()
                
                if len(df_final) > 3:
                    print(f"... y {len(df_final) - 3} resultados más en el Excel")
            
            print(f"\n🎯 ¡Excel con información de descargas listo!")
            print(f"📋 Nuevas columnas: 'Documento Descargado' y 'Anexos Descargados'")
            
        except Exception as e:
            print(f"❌ Error exportando: {e}")
            # Intentar exportación de emergencia simplificada
            try:
                df_backup = pd.DataFrame(resultados) if resultados else pd.DataFrame()
                nombre_backup = f"backup_con_descargas_{nombre_archivo}"
                df_backup.to_excel(nombre_backup, index=False)
                print(f"💾 Exportación de emergencia: {nombre_backup}")
            except Exception as backup_error:
                print(f"❌ Error en exportación de emergencia: {backup_error}")
    
    def cerrar(self):
        """Cierra el driver y limpia recursos"""
        if self.driver:
            self.driver.quit()
        
        # Limpiar carpeta temporal si existe
        try:
            temp_path = "temp_downloads"
            if os.path.exists(temp_path):
                shutil.rmtree(temp_path)
        except:
            pass

# Función principal
if __name__ == "__main__":
    extractor = ExtractorRespondidoPorConDescargas(headless=False)
    
    try:
        print("🔗 EXTRACTOR DE ENLACES 'RESPONDIDO POR' CON DESCARGAS")
        print("="*60)
        print("Este código:")
        print("1. Hace login manual")
        print("2. Va a la página del envío principal")
        print("3. Extrae todos los enlaces de 'Respondido por'")
        print("4. Visita cada enlace y extrae datos completos")
        print("5. 💾 DESCARGA documentos principales → Carpeta 'Cartas'")
        print("6. 📎 DESCARGA anexos → Carpeta 'Anexos'")
        print("7. Determina si envía Comtrade basado en anexos")
        print("8. Exporta todo a Excel con información de descargas")
        print("9. 🗑️ Limpia carpetas automáticamente al inicio")
        
        # Login
        login_ok = extractor.login_interactivo()
        
        if login_ok:
            print(f"\n🎯 PASO 4: NAVEGAR A LA PÁGINA DEL ENVÍO")
            print("="*50)
            
            # Opción rápida para pruebas
            print(f"🧪 MODO DE PRUEBA DISPONIBLE:")
            print(f"Para hacer pruebas rápidas, puedes limitar la cantidad de enlaces a procesar")
            
            # Procesar el envío completo (incluye navegación, límite Y DESCARGAS)
            resultados = extractor.procesar_envio_completo()
            
            # SIEMPRE generar Excel, incluso si no hay resultados
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            nombre_archivo = f"respuestas_comtrade_con_descargas_{timestamp}.xlsx"
            
            print(f"\n📊 GENERANDO EXCEL CON INFORMACIÓN DE DESCARGAS...")
            extractor.exportar_resultados(resultados, nombre_archivo)
            
            if resultados:
                print(f"\n🎯 RESUMEN FINAL CON DESCARGAS:")
                print(f"📁 Archivo generado: {nombre_archivo}")
                print(f"📊 Total respuestas procesadas: {len(resultados)}")
                
                # Mostrar preview de resultados con info de descargas
                print(f"\n📋 Preview de resultados:")
                for resultado in resultados[:3]:  # Mostrar primeros 3
                    doc_status = "✅" if resultado.get('documento_descargado') else "❌"
                    anexos_count = len(resultado.get('anexos_descargados', '').split('|')) if resultado.get('anexos_descargados', '') != 'Sin anexos' else 0
                    print(f"  • {resultado['correlativo']} - {resultado['empresa']}")
                    print(f"    📄 Doc: {doc_status} | 📎 Anexos: {anexos_count} | 📊 {resultado['comtrade_respuesta']}")
                
                if len(resultados) > 3:
                    print(f"  ... y {len(resultados) - 3} más")
                    
                # Estadísticas rápidas INCLUYENDO DESCARGAS
                comtrade_si = len([r for r in resultados if 'sí' in r.get('comtrade_respuesta', '').lower()])
                comtrade_no = len([r for r in resultados if 'no' in r.get('comtrade_respuesta', '').lower()])
                docs_descargados = len([r for r in resultados if r.get('documento_descargado')])
                anexos_descargados = len([r for r in resultados if r.get('anexos_descargados', '') not in ['Sin anexos', 'Error', '']])
                
                print(f"\n📈 Resumen rápido:")
                print(f"   🟢 Envían Comtrade: {comtrade_si}")
                print(f"   🔴 No envían Comtrade: {comtrade_no}")
                print(f"   📊 Otros: {len(resultados) - comtrade_si - comtrade_no}")
                print(f"   💾 Documentos descargados: {docs_descargados}/{len(resultados)}")
                print(f"   📎 Respuestas con anexos: {anexos_descargados}/{len(resultados)}")
                
                # Mostrar conteo de archivos en carpetas
                try:
                    cartas_files = len([f for f in os.listdir(extractor.carpeta_cartas) if os.path.isfile(os.path.join(extractor.carpeta_cartas, f))])
                    anexos_files = len([f for f in os.listdir(extractor.carpeta_anexos) if os.path.isfile(os.path.join(extractor.carpeta_anexos, f))])
                    
                    print(f"\n📁 ARCHIVOS EN CARPETAS:")
                    print(f"   📄 Carpeta 'Cartas': {cartas_files} archivos")
                    print(f"   📎 Carpeta 'Anexos': {anexos_files} archivos")
                    print(f"   📋 Total descargado: {cartas_files + anexos_files} archivos")
                    
                except Exception as e:
                    print(f"⚠️ Error contando archivos: {e}")
                
            else:
                print(f"\n⚠️ No se extrajeron resultados, pero se generó archivo Excel vacío")
                print(f"📁 Archivo: {nombre_archivo}")
                print("\n🔍 Posibles causas:")
                print("1. La URL del envío no tiene enlaces 'Respondido por'")
                print("2. Los enlaces tienen una estructura diferente")
                print("3. Problemas de permisos en la página")
                print("4. La página no cargó completamente")
                print("5. Problemas de descarga de archivos")
            
            # Mostrar ubicación de las carpetas
            print(f"\n📂 UBICACIÓN DE ARCHIVOS DESCARGADOS:")
            print(f"📄 Documentos: {os.path.abspath(extractor.carpeta_cartas)}")
            print(f"📎 Anexos: {os.path.abspath(extractor.carpeta_anexos)}")
            print(f"📊 Excel: {os.path.abspath(nombre_archivo)}")
            
            input("\n⏳ Presiona ENTER para cerrar el navegador...")
        
        else:
            print("❌ No se pudo hacer login")
    
    except Exception as e:
        print(f"❌ Error en la ejecución: {e}")
        import traceback
        traceback.print_exc()
    
    finally:
        extractor.cerrar()