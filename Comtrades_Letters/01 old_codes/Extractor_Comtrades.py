import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import TimeoutException, NoSuchElementException
import re
from datetime import datetime

class CorrespondenciaExtractor:
    def __init__(self, headless=True):
        """
        Inicializa el extractor de correspondencia
        
        Args:
            headless (bool): Si True, ejecuta el navegador sin interfaz gráfica
        """
        self.driver = None
        self.is_logged_in = False
        self.setup_driver(headless)
        
    def setup_driver(self, headless=True):
        """Configura el driver de Chrome con las opciones necesarias"""
        chrome_options = Options()
        if headless:
            chrome_options.add_argument("--headless")
        chrome_options.add_argument("--no-sandbox")
        chrome_options.add_argument("--disable-dev-shm-usage")
        chrome_options.add_argument("--disable-gpu")
        chrome_options.add_argument("--window-size=1920,1080")
        
        try:
            self.driver = webdriver.Chrome(options=chrome_options)
            self.driver.implicitly_wait(10)
        except Exception as e:
            print(f"Error al inicializar el driver: {e}")
            print("Asegúrate de tener ChromeDriver instalado y en el PATH")
    
    def iniciar_sesion(self, usuario=None, password=None, url_login=None):
        """
        Inicia sesión en el sistema de correspondencia
        
        Args:
            usuario (str): Nombre de usuario (opcional, se solicitará si no se proporciona)
            password (str): Contraseña (opcional, se solicitará si no se proporciona)
            url_login (str): URL de login (opcional, detectará automáticamente)
        
        Returns:
            bool: True si el login fue exitoso
        """
        try:
            # URLs comunes de login para el sistema
            posibles_urls_login = [
                "https://correspondencia.coordinador.cl/login",
                "https://correspondencia.coordinador.cl/auth/login",
                "https://correspondencia.coordinador.cl/iniciar-sesion",
                "https://correspondencia.coordinador.cl"
            ]
            
            # Si no se proporciona URL, probar las comunes
            if url_login:
                urls_a_probar = [url_login]
            else:
                urls_a_probar = posibles_urls_login
            
            print("🔐 Iniciando proceso de autenticación...")
            
            for url in urls_a_probar:
                try:
                    print(f"Probando URL de login: {url}")
                    self.driver.get(url)
                    time.sleep(3)
                    
                    # Buscar formulario de login
                    if self.detectar_formulario_login():
                        print(f"✅ Formulario de login encontrado en: {url}")
                        break
                except Exception as e:
                    print(f"❌ Error con URL {url}: {e}")
                    continue
            else:
                print("❌ No se encontró formulario de login en ninguna URL")
                return False
            
            # Solicitar credenciales si no se proporcionaron
            if not usuario:
                usuario = input("👤 Usuario: ")
            if not password:
                import getpass
                password = getpass.getpass("🔑 Contraseña: ")
            
            # Intentar hacer login
            return self.realizar_login(usuario, password)
            
        except Exception as e:
            print(f"❌ Error durante el login: {e}")
            return False
    
    def detectar_formulario_login(self):
        """
        Detecta si hay un formulario de login en la página actual
        
        Returns:
            bool: True si encuentra formulario de login
        """
        try:
            # Selectores comunes para formularios de login
            selectores_login = [
                "input[type='email']",
                "input[name='username']",
                "input[name='usuario']",
                "input[name='user']",
                "input[name='email']",
                "input[id='username']",
                "input[id='usuario']",
                "input[id='email']",
                "input[placeholder*='usuario']",
                "input[placeholder*='email']",
                "form[class*='login']",
                "form[id*='login']"
            ]
            
            for selector in selectores_login:
                elementos = self.driver.find_elements(By.CSS_SELECTOR, selector)
                if elementos:
                    return True
            
            return False
            
        except Exception:
            return False
    
    def realizar_login(self, usuario, password):
        """
        Realiza el proceso de login con las credenciales proporcionadas
        
        Args:
            usuario (str): Nombre de usuario
            password (str): Contraseña
        
        Returns:
            bool: True si el login fue exitoso
        """
        try:
            # Buscar campo de usuario
            campo_usuario = self.encontrar_campo_usuario()
            if not campo_usuario:
                print("❌ No se encontró campo de usuario")
                return False
            
            # Buscar campo de contraseña
            campo_password = self.encontrar_campo_password()
            if not campo_password:
                print("❌ No se encontró campo de contraseña")
                return False
            
            # Limpiar y llenar campos
            print("📝 Llenando credenciales...")
            campo_usuario.clear()
            campo_usuario.send_keys(usuario)
            
            campo_password.clear()
            campo_password.send_keys(password)
            
            # Buscar y hacer click en botón de login
            boton_login = self.encontrar_boton_login()
            if boton_login:
                print("🔄 Enviando formulario...")
                boton_login.click()
            else:
                # Si no encuentra botón, intentar enviar con Enter
                print("🔄 Enviando con Enter...")
                campo_password.send_keys("\n")
            
            # Esperar respuesta
            time.sleep(5)
            
            # Verificar si el login fue exitoso
            if self.verificar_login_exitoso():
                print("✅ Login exitoso!")
                self.is_logged_in = True
                return True
            else:
                print("❌ Login falló - verificar credenciales")
                return False
                
        except Exception as e:
            print(f"❌ Error realizando login: {e}")
            return False
    
    def encontrar_campo_usuario(self):
        """Encuentra el campo de usuario usando diferentes selectores"""
        selectores_usuario = [
            "input[type='email']",
            "input[name='username']",
            "input[name='usuario']",
            "input[name='user']",
            "input[name='email']",
            "input[name='login']",
            "input[id='username']",
            "input[id='usuario']",
            "input[id='user']",
            "input[id='email']",
            "input[id='login']",
            "input[placeholder*='usuario' i]",
            "input[placeholder*='email' i]",
            "input[placeholder*='user' i]"
        ]
        
        for selector in selectores_usuario:
            try:
                elemento = self.driver.find_element(By.CSS_SELECTOR, selector)
                if elemento.is_displayed():
                    return elemento
            except NoSuchElementException:
                continue
        
        return None
    
    def encontrar_campo_password(self):
        """Encuentra el campo de contraseña"""
        selectores_password = [
            "input[type='password']",
            "input[name='password']",
            "input[name='contraseña']",
            "input[name='pass']",
            "input[id='password']",
            "input[id='contraseña']",
            "input[id='pass']",
            "input[placeholder*='contraseña' i]",
            "input[placeholder*='password' i]"
        ]
        
        for selector in selectores_password:
            try:
                elemento = self.driver.find_element(By.CSS_SELECTOR, selector)
                if elemento.is_displayed():
                    return elemento
            except NoSuchElementException:
                continue
        
        return None
    
    def encontrar_boton_login(self):
        """Encuentra el botón de login/enviar"""
        selectores_boton = [
            "button[type='submit']",
            "input[type='submit']",
            "button[name='login']",
            "button[id='login']",
            "button[class*='login']",
            "input[value*='Iniciar' i]",
            "input[value*='Entrar' i]",
            "input[value*='Login' i]",
            "button:contains('Iniciar')",
            "button:contains('Entrar')",
            "button:contains('Login')",
            ".btn-login",
            ".login-button"
        ]
        
        for selector in selectores_boton:
            try:
                if ":contains(" in selector:
                    # Para selectores que usan :contains, usar XPath
                    texto = selector.split("'")[1]
                    xpath = f"//button[contains(text(), '{texto}')]"
                    elemento = self.driver.find_element(By.XPATH, xpath)
                else:
                    elemento = self.driver.find_element(By.CSS_SELECTOR, selector)
                
                if elemento.is_displayed():
                    return elemento
            except NoSuchElementException:
                continue
        
        return None
    
    def verificar_login_exitoso(self):
        """
        Verifica si el login fue exitoso
        
        Returns:
            bool: True si está autenticado
        """
        try:
            # Indicadores de login exitoso
            indicadores_exito = [
                # URL cambió a dashboard/home
                "dashboard" in self.driver.current_url.lower(),
                "home" in self.driver.current_url.lower(),
                "correspondencia" in self.driver.current_url.lower(),
                
                # Elementos que aparecen después del login
                len(self.driver.find_elements(By.CSS_SELECTOR, "[class*='logout']")) > 0,
                len(self.driver.find_elements(By.CSS_SELECTOR, "[href*='logout']")) > 0,
                len(self.driver.find_elements(By.CSS_SELECTOR, ".user-menu")) > 0,
                len(self.driver.find_elements(By.CSS_SELECTOR, ".navbar-user")) > 0,
                
                # Ya no hay formulario de login
                len(self.driver.find_elements(By.CSS_SELECTOR, "input[type='password']")) == 0
            ]
            
            # Indicadores de login fallido
            indicadores_fallo = [
                "login" in self.driver.current_url.lower(),
                len(self.driver.find_elements(By.CSS_SELECTOR, ".error")) > 0,
                len(self.driver.find_elements(By.CSS_SELECTOR, ".alert-danger")) > 0,
                "error" in self.driver.page_source.lower(),
                "incorrecto" in self.driver.page_source.lower(),
                "inválido" in self.driver.page_source.lower()
            ]
            
            # Si hay indicadores de fallo, retornar False
            if any(indicadores_fallo):
                return False
            
            # Si hay indicadores de éxito, retornar True
            return any(indicadores_exito)
            
        except Exception as e:
            print(f"Error verificando login: {e}")
            return False
    
    def login_interactivo(self):
        """
        Permite al usuario hacer login manualmente mientras observa el navegador
        
        Returns:
            bool: True si el usuario confirma que hizo login exitoso
        """
        print("\n🖱️  MODO LOGIN INTERACTIVO")
        print("="*50)
        print("El navegador se abrirá y podrás hacer login manualmente.")
        print("Después de hacer login exitoso, vuelve a esta consola.")
        
        try:
            # Ir a la página principal
            self.driver.get("https://correspondencia.coordinador.cl")
            
            input("\n⏳ Presiona ENTER después de hacer login exitoso en el navegador...")
            
            # Verificar si el login fue exitoso
            if self.verificar_login_exitoso():
                print("✅ Login confirmado!")
                self.is_logged_in = True
                return True
            else:
                respuesta = input("❓ ¿Estás seguro de que hiciste login? (s/n): ")
                if respuesta.lower() in ['s', 'sí', 'si', 'y', 'yes']:
                    self.is_logged_in = True
                    return True
                else:
                    return False
        except Exception as e:
            print("aaaaaaaaaaaaaaaaa")
            return False        
                    
    def buscar_correspondencia(self, query, period="", doc_type="T"):
        """
        Busca correspondencia en el sistema
        
        Args:
            query (str): Término de búsqueda (ej: "DE01746-25")
            period (str): Período de búsqueda (opcional)
            doc_type (str): Tipo de documento (por defecto "T")
        
        Returns:
            list: Lista de diccionarios con la información extraída
        """
        # Verificar si está autenticado
        if not self.is_logged_in:
            print("⚠️  No estás autenticado. Inicia sesión primero.")
            return []
        
        try:
            print(f"🔍 Buscando: {query}")
            
            # Paso 1: Ir a la página de búsqueda principal
            print("📍 Navegando a la página de búsqueda...")
            self.driver.get("https://correspondencia.coordinador.cl/correspondencia/busqueda")
            time.sleep(3)
            
            # Verificar si la página redirige al login (sesión expirada)
            if "login" in self.driver.current_url.lower():
                print("⚠️  Sesión expirada. Necesitas autenticarte nuevamente.")
                self.is_logged_in = False
                return []
            
            # Paso 2: Llenar el formulario de búsqueda
            print("📝 Llenando formulario de búsqueda...")
            success = self.llenar_formulario_busqueda(query, period, doc_type)
            
            if not success:
                print("❌ No se pudo llenar el formulario de búsqueda")
                return []
            
            # Paso 3: Hacer click en BUSCAR
            print("🔍 Ejecutando búsqueda...")
            if not self.ejecutar_busqueda():
                print("❌ No se pudo ejecutar la búsqueda")
                return []
            
            # Paso 4: Esperar y extraer resultados
            print("📊 Extrayendo resultados...")
            time.sleep(5)  # Esperar a que carguen los resultados
            
            return self.extraer_datos_tabla_correspondencia()
            
        except Exception as e:
            print(f"❌ Error al buscar correspondencia: {e}")
            return []
    
    def llenar_formulario_busqueda(self, query, period, doc_type):
        """
        Llena el formulario de búsqueda en la página
        
        Returns:
            bool: True si se llenó correctamente
        """
        try:
            # Buscar campo de términos de búsqueda
            campo_query = None
            selectores_query = [
                "input[placeholder*='Búsqueda']",
                "input[name*='query']",
                "input[id*='search']",
                "input[type='text']",
                "#searchTerms",
                ".search-input"
            ]
            
            for selector in selectores_query:
                try:
                    campo_query = self.driver.find_element(By.CSS_SELECTOR, selector)
                    if campo_query.is_displayed():
                        break
                except NoSuchElementException:
                    continue
            
            if not campo_query:
                print("❌ No se encontró el campo de búsqueda")
                return False
            
            # Limpiar y llenar campo de búsqueda
            campo_query.clear()
            campo_query.send_keys(query)
            print(f"✅ Campo de búsqueda llenado con: {query}")
            
            # Llenar período si se proporciona
            if period:
                campo_period = None
                selectores_period = [
                    "input[name*='period']",
                    "input[placeholder*='Período']",
                    "#period"
                ]
                
                for selector in selectores_period:
                    try:
                        campo_period = self.driver.find_element(By.CSS_SELECTOR, selector)
                        if campo_period.is_displayed():
                            campo_period.clear()
                            campo_period.send_keys(period)
                            print(f"✅ Período llenado con: {period}")
                            break
                    except NoSuchElementException:
                        continue
            
            # Seleccionar tipo de documento si es necesario
            if doc_type and doc_type != "T":
                try:
                    select_doc_type = self.driver.find_element(By.CSS_SELECTOR, "select")
                    select_doc_type.click()
                    time.sleep(1)
                    
                    # Buscar la opción correcta
                    option = self.driver.find_element(By.CSS_SELECTOR, f"option[value='{doc_type}']")
                    option.click()
                    print(f"✅ Tipo de documento seleccionado: {doc_type}")
                except NoSuchElementException:
                    print("⚠️  No se pudo cambiar el tipo de documento")
            
            return True
            
        except Exception as e:
            print(f"❌ Error llenando formulario: {e}")
            return False
    
    def ejecutar_busqueda(self):
        """
        Hace click en el botón BUSCAR
        
        Returns:
            bool: True si se ejecutó correctamente
        """
        try:
            # Buscar botón de búsqueda
            selectores_boton = [
                "button:contains('BUSCAR')",
                "input[value*='BUSCAR']",
                "button[type='submit']",
                ".btn-search",
                "#searchButton",
                "button.btn-primary"
            ]
            
            boton_buscar = None
            for selector in selectores_boton:
                try:
                    if ":contains(" in selector:
                        # Usar XPath para texto
                        boton_buscar = self.driver.find_element(By.XPATH, "//button[contains(text(), 'BUSCAR')]")
                    else:
                        boton_buscar = self.driver.find_element(By.CSS_SELECTOR, selector)
                    
                    if boton_buscar.is_displayed():
                        break
                except NoSuchElementException:
                    continue
            
            if not boton_buscar:
                print("❌ No se encontró el botón BUSCAR")
                return False
            
            # Hacer click en el botón
            boton_buscar.click()
            print("✅ Botón BUSCAR presionado")
            return True
            
        except Exception as e:
            print(f"❌ Error ejecutando búsqueda: {e}")
            return False
    
    def extraer_datos_tabla_correspondencia(self):
        """
        Extrae los datos específicos de la tabla de correspondencia
        
        Returns:
            list: Lista de diccionarios con los datos extraídos
        """
        resultados = []
        
        try:
            # Esperar a que aparezcan los resultados
            wait = WebDriverWait(self.driver, 15)
            
            # Buscar la sección de resultados
            try:
                resultados_seccion = wait.until(
                    EC.presence_of_element_located((By.XPATH, "//h2[contains(text(), 'RESULTADOS')]"))
                )
                print("✅ Sección de resultados encontrada")
            except TimeoutException:
                print("⚠️  No se encontró la sección de resultados, intentando extraer datos disponibles...")
            
            # Buscar la tabla de resultados
            tabla = None
            selectores_tabla = [
                "table",
                ".table",
                ".results-table",
                "[class*='table']"
            ]
            
            for selector in selectores_tabla:
                try:
                    tabla = self.driver.find_element(By.CSS_SELECTOR, selector)
                    if tabla.is_displayed():
                        print(f"✅ Tabla encontrada con selector: {selector}")
                        break
                except NoSuchElementException:
                    continue
            
            if not tabla:
                print("❌ No se encontró tabla de resultados")
                return []
            
            # Extraer header de la tabla para identificar columnas
            headers = []
            try:
                header_row = tabla.find_element(By.TAG_NAME, "thead").find_element(By.TAG_NAME, "tr")
                header_cells = header_row.find_elements(By.TAG_NAME, "th")
                headers = [cell.text.strip() for cell in header_cells]
                print(f"✅ Headers encontrados: {headers}")
            except NoSuchElementException:
                print("⚠️  No se encontraron headers de tabla")
            
            # Extraer filas de datos
            tbody = tabla.find_element(By.TAG_NAME, "tbody")
            filas = tbody.find_elements(By.TAG_NAME, "tr")
            
            print(f"📊 Encontradas {len(filas)} filas de resultados")
            
            for i, fila in enumerate(filas):
                try:
                    celdas = fila.find_elements(By.TAG_NAME, "td")
                    
                    if len(celdas) == 0:
                        continue
                    
                    # Extraer datos de cada celda
                    datos_fila = {}
                    
                    for j, celda in enumerate(celdas):
                        texto_celda = celda.text.strip()
                        
                        # Asignar datos según la posición y headers conocidos
                        if j < len(headers):
                            header = headers[j]
                        else:
                            header = f"Columna_{j+1}"
                        
                        datos_fila[header] = texto_celda
                        
                        # También buscar enlaces dentro de la celda
                        enlaces = celda.find_elements(By.TAG_NAME, "a")
                        if enlaces:
                            datos_fila[f"{header}_enlaces"] = [enlace.get_attribute('href') for enlace in enlaces]
                    
                    # Procesar y limpiar los datos extraídos
                    resultado_procesado = self.procesar_fila_correspondencia(datos_fila, i+1)
                    
                    if resultado_procesado:
                        resultados.append(resultado_procesado)
                        print(f"✅ Fila {i+1} procesada: {resultado_procesado.get('correlativo', 'N/A')}")
                
                except Exception as e:
                    print(f"⚠️  Error procesando fila {i+1}: {e}")
                    continue
            
            print(f"🎉 Total de resultados extraídos: {len(resultados)}")
            return resultados
            
        except Exception as e:
            print(f"❌ Error extrayendo datos de tabla: {e}")
            return []
    
    def procesar_fila_correspondencia(self, datos_fila, numero_fila):
        """
        Procesa y estructura los datos de una fila de correspondencia
        
        Args:
            datos_fila (dict): Datos raw de la fila
            numero_fila (int): Número de fila
        
        Returns:
            dict: Datos estructurados
        """
        try:
            # Mapear campos conocidos
            resultado = {
                'numero_fila': numero_fila,
                'correlativo': self.extraer_campo(datos_fila, ['Correlativo', 'correlativo', 'Código']),
                'tipo_documento': self.extraer_campo(datos_fila, ['Tipo Documento', 'Tipo', 'tipo']),
                'fecha_envio': self.extraer_campo(datos_fila, ['Fecha', 'fecha', 'Fecha Envío']),
                'empresa': self.extraer_campo(datos_fila, ['Empresa(s)', 'Empresa', 'empresa', 'Empresas']),
                'remitente': self.extraer_campo(datos_fila, ['Remitente', 'remitente']),
                'materia_macro': self.extraer_campo(datos_fila, ['Materia Macro', 'Materia', 'materia']),
                'materia_micro': self.extraer_campo(datos_fila, ['Materia Micro', 'Micro', 'micro']),
                'referencia': self.extraer_campo(datos_fila, ['Referencia', 'referencia', 'Ref']),
                'comtrade_respuesta': self.determinar_comtrade_respuesta(datos_fila),
                'datos_completos': datos_fila
            }
            
            return resultado
            
        except Exception as e:
            print(f"Error procesando fila {numero_fila}: {e}")
            return None
    
    def extraer_campo(self, datos_fila, posibles_nombres):
        """
        Extrae un campo usando múltiples nombres posibles
        
        Args:
            datos_fila (dict): Datos de la fila
            posibles_nombres (list): Lista de nombres posibles para el campo
        
        Returns:
            str: Valor del campo o None
        """
        for nombre in posibles_nombres:
            if nombre in datos_fila and datos_fila[nombre]:
                return datos_fila[nombre]
        return None
    
    def determinar_comtrade_respuesta(self, datos_fila):
        """
        Determina la respuesta de Comtrade basado en todos los datos de la fila
        
        Args:
            datos_fila (dict): Datos completos de la fila
        
        Returns:
            str: Estado de respuesta Comtrade
        """
        # Combinar todo el texto de la fila
        texto_completo = ' '.join([str(valor) for valor in datos_fila.values() if valor]).upper()
        
        # Buscar patrones específicos
        if 'COMTRADE' in texto_completo:
            if any(palabra in texto_completo for palabra in ['ENVÍA', 'ENVIA', 'SÍ', 'SI', 'RESPONDE']):
                return 'Sí envía'
            elif any(palabra in texto_completo for palabra in ['NO ENVÍA', 'NO ENVIA', 'NO RESPONDE', 'NO']):
                return 'No envía'
            else:
                return 'Menciona Comtrade'
        
        # Buscar en referencia específicamente
        referencia = self.extraer_campo(datos_fila, ['Referencia', 'referencia', 'Ref'])
        if referencia:
            ref_upper = referencia.upper()
            if 'RESPONDE' in ref_upper and 'CARTA' in ref_upper:
                return 'Sí envía'
            elif 'EVENTO' in ref_upper:
                return 'Evento (verificar Comtrade)'
        
        return 'Sin información'
        """
        Busca correspondencia en el sistema
        
        Args:
            query (str): Término de búsqueda (ej: "DE01746-25")
            period (str): Período de búsqueda (opcional)
            doc_type (str): Tipo de documento (por defecto "T")
        
        Returns:
            list: Lista de diccionarios con la información extraída
        """
        url = f"https://correspondencia.coordinador.cl/correspondencia/busqueda?query={query}&period={period}&doc_type={doc_type}"
        
        try:
            print(f"Accediendo a: {url}")
            self.driver.get(url)
            
            # Esperar a que la página cargue completamente
            time.sleep(3)
            
            # Buscar la tabla de resultados
            return self.extraer_datos_tabla()
            
        except Exception as e:
            print(f"Error al buscar correspondencia: {e}")
            return []
    
    def extraer_datos_tabla(self):
        """
        Extrae los datos de la tabla de resultados
        
        Returns:
            list: Lista de diccionarios con los datos extraídos
        """
        resultados = []
        
        try:
            # Esperar a que aparezca la tabla de resultados
            wait = WebDriverWait(self.driver, 15)
            
            # Intentar encontrar diferentes selectores posibles para la tabla
            posibles_selectores = [
                "table",
                ".table",
                "[class*='table']",
                ".results-table",
                "#results-table",
                ".correspondence-table"
            ]
            
            tabla = None
            for selector in posibles_selectores:
                try:
                    tabla = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, selector)))
                    print(f"Tabla encontrada con selector: {selector}")
                    break
                except TimeoutException:
                    continue
            
            if not tabla:
                # Si no encuentra tabla, buscar filas directamente
                filas = self.driver.find_elements(By.CSS_SELECTOR, "tr, .row, .result-item")
                if filas:
                    print(f"Encontradas {len(filas)} filas sin tabla contenedora")
                    return self.procesar_filas(filas)
                else:
                    print("No se encontraron resultados en la página")
                    # Guardar screenshot para debugging
                    self.driver.save_screenshot("debug_no_results.png")
                    return []
            
            # Extraer filas de la tabla
            filas = tabla.find_elements(By.TAG_NAME, "tr")
            
            if len(filas) <= 1:  # Solo header o sin datos
                print("No se encontraron datos en la tabla")
                return []
            
            # Procesar filas (saltando el header)
            for i, fila in enumerate(filas[1:], 1):
                try:
                    datos = self.extraer_datos_fila(fila, i)
                    if datos:
                        resultados.append(datos)
                except Exception as e:
                    print(f"Error procesando fila {i}: {e}")
                    continue
            
        except TimeoutException:
            print("Timeout esperando la tabla de resultados")
            # Intentar extraer cualquier información visible
            return self.extraer_datos_alternativos()
        except Exception as e:
            print(f"Error extrayendo datos de la tabla: {e}")
        
        return resultados
    
    def procesar_filas(self, filas):
        """Procesa filas encontradas directamente"""
        resultados = []
        for i, fila in enumerate(filas):
            try:
                datos = self.extraer_datos_fila(fila, i)
                if datos:
                    resultados.append(datos)
            except Exception as e:
                print(f"Error procesando fila {i}: {e}")
        return resultados
    
    def extraer_datos_fila(self, fila, numero_fila):
        """
        Extrae datos de una fila específica
        
        Args:
            fila: Elemento WebElement de la fila
            numero_fila: Número de fila para referencia
        
        Returns:
            dict: Diccionario con los datos extraídos
        """
        try:
            # Extraer todas las celdas de la fila
            celdas = fila.find_elements(By.TAG_NAME, "td")
            
            if not celdas:
                # Intentar con otros selectores
                celdas = fila.find_elements(By.CSS_SELECTOR, "div, span")
            
            if not celdas:
                return None
            
            # Extraer texto de todas las celdas
            textos_celdas = [celda.text.strip() for celda in celdas if celda.text.strip()]
            
            # Buscar enlaces para obtener más detalles
            enlaces = fila.find_elements(By.TAG_NAME, "a")
            
            # Crear diccionario con los datos
            datos = {
                'numero_fila': numero_fila,
                'fecha_envio': self.extraer_fecha(textos_celdas),
                'empresa': self.extraer_empresa(textos_celdas),
                'correlativo': self.extraer_correlativo(textos_celdas),
                'comtrade_respuesta': self.verificar_comtrade(textos_celdas),
                'texto_completo': ' | '.join(textos_celdas),
                'enlaces': [enlace.get_attribute('href') for enlace in enlaces if enlace.get_attribute('href')]
            }
            
            print(f"Fila {numero_fila} procesada: {datos['correlativo']}")
            return datos
            
        except Exception as e:
            print(f"Error extrayendo datos de fila {numero_fila}: {e}")
            return None
    
    def extraer_fecha(self, textos):
        """Extrae fecha de envío de los textos"""
        patrones_fecha = [
            r'\d{2}/\d{2}/\d{4}',
            r'\d{2}-\d{2}-\d{4}',
            r'\d{4}/\d{2}/\d{2}',
            r'\d{4}-\d{2}-\d{2}'
        ]
        
        for texto in textos:
            for patron in patrones_fecha:
                match = re.search(patron, texto)
                if match:
                    return match.group()
        return None
    
    def extraer_empresa(self, textos):
        """Extrae nombre de empresa de los textos"""
        # Buscar patrones comunes de nombres de empresa
        for texto in textos:
            # Si contiene palabras clave de empresa
            if any(palabra in texto.upper() for palabra in ['S.A.', 'LTDA', 'SPA', 'ENEL', 'CGE', 'CHILQUINTA']):
                return texto
        
        # Si no encuentra patrón específico, devolver el texto más largo (probable empresa)
        if textos:
            return max(textos, key=len) if len(max(textos, key=len)) > 10 else None
        return None
    
    def extraer_correlativo(self, textos):
        """Extrae código correlativo de los textos"""
        patrones_correlativo = [
            r'[A-Z]{2}\d{5}-\d{2}',  # Ej: DE01746-25
            r'[A-Z]+\d+-\d+',
            r'\d{4,}-\d{2,}'
        ]
        
        for texto in textos:
            for patron in patrones_correlativo:
                match = re.search(patron, texto)
                if match:
                    return match.group()
        return None
    
    def verificar_comtrade(self, textos):
        """Verifica si hay mención de Comtrade en los textos"""
        texto_completo = ' '.join(textos).upper()
        
        if 'COMTRADE' in texto_completo:
            if any(palabra in texto_completo for palabra in ['ENVÍA', 'ENVIA', 'SÍ', 'SI', 'POSITIVO']):
                return 'Sí envía'
            elif any(palabra in texto_completo for palabra in ['NO ENVÍA', 'NO ENVIA', 'NO', 'NEGATIVO']):
                return 'No envía'
            else:
                return 'Menciona Comtrade'
        
        return 'Sin información'
    
    def extraer_datos_alternativos(self):
        """Método alternativo para extraer datos cuando no se encuentra tabla"""
        try:
            # Buscar cualquier texto que contenga información útil
            elementos_texto = self.driver.find_elements(By.CSS_SELECTOR, "div, p, span")
            textos = [elem.text.strip() for elem in elementos_texto if elem.text.strip()]
            
            if textos:
                return [{
                    'numero_fila': 1,
                    'fecha_envio': self.extraer_fecha(textos),
                    'empresa': self.extraer_empresa(textos),
                    'correlativo': self.extraer_correlativo(textos),
                    'comtrade_respuesta': self.verificar_comtrade(textos),
                    'texto_completo': ' | '.join(textos[:10]),  # Primeros 10 elementos
                    'enlaces': []
                }]
        except Exception as e:
            print(f"Error en extracción alternativa: {e}")
        
        return []
    
    def buscar_multiples_correlativos(self, correlativos_lista):
        """
        Busca múltiples correlativos y consolida los resultados
        
        Args:
            correlativos_lista (list): Lista de correlativos a buscar
        
        Returns:
            pandas.DataFrame: DataFrame con todos los resultados
        """
        todos_resultados = []
        
        for correlativo in correlativos_lista:
            print(f"\nBuscando: {correlativo}")
            resultados = self.buscar_correspondencia(correlativo)
            
            for resultado in resultados:
                resultado['correlativo_buscado'] = correlativo
                todos_resultados.append(resultado)
            
            time.sleep(2)  # Pausa entre búsquedas
        
        return pd.DataFrame(todos_resultados)
    
    def exportar_resultados(self, resultados, nombre_archivo="correspondencia_extraida.xlsx"):
        """
        Exporta los resultados a un archivo Excel
        
        Args:
            resultados: Lista de diccionarios o DataFrame
            nombre_archivo (str): Nombre del archivo a crear
        """
        if isinstance(resultados, list):
            df = pd.DataFrame(resultados)
        else:
            df = resultados
        
        if not df.empty:
            # Agregar timestamp
            df['fecha_extraccion'] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            
            # Reordenar columnas
            columnas_orden = ['correlativo_buscado', 'correlativo', 'fecha_envio', 'empresa', 
                            'comtrade_respuesta', 'numero_fila', 'fecha_extraccion', 
                            'texto_completo', 'enlaces']
            
            columnas_existentes = [col for col in columnas_orden if col in df.columns]
            df = df[columnas_existentes]
            
            df.to_excel(nombre_archivo, index=False)
            print(f"Resultados exportados a: {nombre_archivo}")
        else:
            print("No hay resultados para exportar")
    
    def cerrar(self):
        """Cierra el driver del navegador"""
        if self.driver:
            self.driver.quit()

# Ejemplo de uso
if __name__ == "__main__":
    # Crear extractor
    extractor = CorrespondenciaExtractor(headless=False)
    
    try:
        print("🌐 Abriendo navegador para login manual...")
        print("="*50)
        print("1. Se abrirá el navegador")
        print("2. Ingresa tus credenciales manualmente")
        print("3. Una vez que hayas hecho login, vuelve aquí")
        print("4. Presiona ENTER para continuar con la extracción")
        print("="*50)
        
        # Login interactivo directo (sin intentar automático)
        login_exitoso = extractor.login_interactivo()
        
        # Una vez autenticado, buscar correspondencia
        if login_exitoso:
            print("\n🔍 Opciones de búsqueda:")
            print("1. Buscar correlativo específico (ej: DE01746-25)")
            print("2. Buscar múltiples correlativos")
            print("3. Extraer TODOS los resultados de una búsqueda amplia")
            
            opcion = input("\nSelecciona una opción (1, 2, o 3): ").strip()
            
            if opcion == "1":
                # Búsqueda específica
                correlativo = input("Ingresa el correlativo a buscar (ej: DE01746-25): ").strip()
                if not correlativo:
                    correlativo = "DE01746-25"  # Default
                
                print(f"\n🔍 Buscando: {correlativo}")
                resultados = extractor.buscar_correspondencia(correlativo)
                
                if resultados:
                    nombre_archivo = f"correspondencia_{correlativo.replace('-', '_')}.xlsx"
                    extractor.exportar_resultados(resultados, nombre_archivo)
                    print(f"✅ Excel generado: {nombre_archivo}")
                else:
                    print("❌ No se encontraron resultados")
            
            elif opcion == "2":
                # Búsqueda múltiple
                print("\n📝 Ingresa los correlativos separados por comas:")
                print("Ejemplo: DE01746-25, DE01747-25, DE01748-25")
                entrada_correlativos = input("Correlativos: ").strip()
                
                if entrada_correlativos:
                    correlativos_lista = [c.strip() for c in entrada_correlativos.split(',')]
                    print(f"\n🔍 Buscando {len(correlativos_lista)} correlativos...")
                    
                    todos_resultados = extractor.buscar_multiples_correlativos(correlativos_lista)
                    
                    if todos_resultados:
                        nombre_archivo = "correspondencia_multiple.xlsx"
                        extractor.exportar_resultados(todos_resultados, nombre_archivo)
                        print(f"✅ Excel múltiple generado: {nombre_archivo}")
                    else:
                        print("❌ No se encontraron resultados")
                else:
                    print("❌ No se ingresaron correlativos")
            
            elif opcion == "3":
                # Extracción completa de todos los resultados
                print("\n🎯 EXTRACCIÓN COMPLETA DE RESULTADOS")
                print("="*50)
                print("Esta opción extraerá TODOS los resultados visibles en la búsqueda actual.")
                print("Perfecto para obtener los 193 resultados que mencionas.")
                
                # Opciones de búsqueda amplia
                print("\nOpciones de búsqueda:")
                print("a) Buscar por término específico (ej: DE01746-25)")
                print("b) Buscar todos los documentos de un período")
                print("c) Extraer de la página actual (si ya tienes resultados abiertos)")
                
                sub_opcion = input("\nSelecciona (a, b, o c): ").strip().lower()
                
                if sub_opcion == "a":
                    termino = input("Ingresa el término de búsqueda (ej: DE01746-25): ").strip()
                    if not termino:
                        termino = "DE01746-25"
                    
                    print(f"\n🔍 Extrayendo TODOS los resultados para: {termino}")
                    resultados = extractor.buscar_correspondencia(termino)
                    
                elif sub_opcion == "b":
                    # Búsqueda por período
                    print("\n📅 Búsqueda por período:")
                    period = input("Ingresa período (ej: 2025, 2025-02, o deja vacío): ").strip()
                    doc_type = input("Tipo de documento (T para todos, o deja vacío): ").strip()
                    if not doc_type:
                        doc_type = "T"
                    
                    print(f"\n🔍 Extrayendo resultados del período: {period or 'todos'}")
                    resultados = extractor.buscar_correspondencia("", period, doc_type)
                    
                elif sub_opcion == "c":
                    # Extraer de página actual
                    print("\n📄 Extrayendo de la página actual...")
                    print("Asegúrate de que ya tengas los resultados abiertos en el navegador.")
                    input("Presiona ENTER cuando estés listo...")
                    
                    # Ir directamente a extraer datos sin hacer nueva búsqueda
                    resultados = extractor.extraer_datos_tabla_correspondencia()
                    
                else:
                    print("❌ Opción no válida")
                    resultados = []
                
                # Procesar resultados de extracción completa
                if resultados:
                    print(f"\n🎉 ¡Extracción exitosa!")
                    print(f"📊 Total de resultados extraídos: {len(resultados)}")
                    
                    # Generar nombre de archivo descriptivo
                    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                    nombre_archivo = f"correspondencia_completa_{timestamp}.xlsx"
                    
                    extractor.exportar_resultados(resultados, nombre_archivo)
                    print(f"✅ Excel completo generado: {nombre_archivo}")
                    
                    # Mostrar estadísticas
                    print(f"\n📈 Estadísticas de la extracción:")
                    print(f"   • Total de registros: {len(resultados)}")
                    
                    # Contar respuestas Comtrade
                    comtrade_stats = {}
                    for resultado in resultados:
                        respuesta = resultado.get('comtrade_respuesta', 'Sin información')
                        comtrade_stats[respuesta] = comtrade_stats.get(respuesta, 0) + 1
                    
                    print(f"   • Respuestas Comtrade:")
                    for respuesta, count in comtrade_stats.items():
                        print(f"     - {respuesta}: {count}")
                    
                    # Empresas únicas
                    empresas = set()
                    for resultado in resultados:
                        empresa = resultado.get('empresa', '')
                        if empresa:
                            empresas.add(empresa)
                    print(f"   • Empresas únicas: {len(empresas)}")
                    
                else:
                    print("❌ No se encontraron resultados para extraer")
            
            else:
                print("❌ Opción no válida")
            
            # Opción para mantener navegador abierto
            print(f"\n🌐 El navegador permanece abierto para que puedas verificar los resultados.")
            continuar = input("¿Quieres hacer otra extracción? (s/n): ").strip().lower()
            
            if continuar in ['s', 'sí', 'si', 'y', 'yes']:
                print("🔄 Reiniciando proceso...")
                # Aquí podrías hacer un loop, pero por simplicidad mostramos el mensaje
                print("Ejecuta el script nuevamente para otra extracción.")
            
            input("\n⏳ Presiona ENTER para cerrar el navegador...")
            
        else:
            print("❌ No se detectó login exitoso. Verifica que hayas ingresado correctamente.")
        
    except Exception as e:
        print(f"❌ Error en la ejecución: {e}")
        import traceback
        traceback.print_exc()
    finally:
        extractor.cerrar()