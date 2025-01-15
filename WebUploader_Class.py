# -*- coding: utf-8 -*-
"""
@author: Alejandro López
Clase WebUploader

Ejecución principal del programa para cargar información de un archivo excel a una página web mediante Selenium

"""
from pandas import isnull
from time import time, sleep
from selenium import webdriver
from selenium.webdriver.edge.options import Options
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException, ElementNotInteractableException, ElementClickInterceptedException, StaleElementReferenceException, NoSuchElementException
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys

tiempo_wait_driver = 30

class WebUploader_Class:
    def __init__(self, Data_config, verbose=True):
        """
        Esta clase representa el driver usado por Selenium que permite interactuar con paginas web y hacer Web Scraping
        
        Parametros
        ----------
        webdriver_path : `text`
            Path donde se encuentra ubicado el webdriver de Chrome (Se parametriza desde 'Config.py')

        url_path : `text`
            URL del login de MX (Se parametriza desde 'Config.py')
        """
        
        # Configurar las opciones del driver de Chrome
        brwoser_options = Options()
        #chrome_options.add_argument("--start-minimized")  # Abrir el navegador minimizado
        brwoser_options.add_argument("--window-position=-1000,300")  # Coordenadas de la segunda pantalla

        self.driver = webdriver.Edge(brwoser_options)      
        #self.configurar_ventana(pantalla=0, tipo='min')
        self.driver.maximize_window()
        
        self.driver.get(Data_config["url_path"])
        
        self.logs = ""
        self.errors = ""

        if verbose == True:
            self.logs += f"Web Driver iniciado correctamente, URL:'{Data_config['url_path']}' - Lista para ser usada por el bot\n"
            self.logs += "--------------------------------------------------------\n"
        
    def cerrar_navegador(self):
        self.driver.quit()
    
    def log_in(self, Data_config, IDs, verbose=True):
        """
        Este metodo permite iniciar sesión en MX con el username y password indicado en los parametros
        
        Parametros
        ----------
        Data_config : `json`
                Información del usuario de tipo User_Config.json
            
        IDs : `json`
                IDs del login de tipo ID_Config.json, ejemplo:IDs["Login"]
            
        verbose : `Boolean`
            mostrar información del proceso por pantalla. (default ``True``)
        
        --------------------
        Parametros de Salida
        --------------------
            return validador: `Boolean`
                    Salida del metodo, indica si el buscador encontro la OT
        """
        
        self.username = Data_config["USER"]
        self.password = Data_config["PASS"]
       
        #Inciar y validar si la sesión ha sido iniciada
        try:           
            username_field = self.click_until_interactable(By.ID, IDs["username"])
            username_field.send_keys(self.username)
            password_field = self.click_until_interactable(By.ID, IDs["password"])
            password_field.send_keys(self.password)
            
            # Hacer click en la opción de lenguaje seleccionada por el usuario     
            list_button = self.click_until_interactable(By.XPATH, IDs["loginlist"]["list_name"])  # Wait a que el boton sea localizado                
            list_selected = self.click_until_interactable(By.XPATH, IDs["loginlist"][Data_config["idioma"]])
            
            # Hacer click en el botón de iniciar sesión
            submit_button = self.click_until_interactable(By.ID, IDs["loginbutton"])

            # Validación del inicio de sesion
            name_sesion = self.click_until_interactable(By.ID, IDs["validador"], click=False)
            sesion_started = name_sesion.is_displayed()
            if verbose==True:
                self.logs += f"Sesion iniciada correctamente con la cuenta: {self.username}\n"
                self.logs += "--------------------------------------------------------\n"
                
            return sesion_started
        
        # Error al no encontrar el elemento buscado por Selenium
        except TimeoutException:
            self.errors += f"Error: \n-Web_Uploader_Class \n-Metodo: log_in \nDescripción: No se ha podido iniciar sesión correctamente con la cuenta {self.username}\n"
            sesion_started = False
            return sesion_started
        
    
    #---------------------------------------------------------------------------------------------------
    def log_menu(self, menu, verbose=False):
        """
        Este metodo permite seleccionar el menu deseado por el usuario en MX
        
        Parametros
        ----------
        menu : `text`
            Listado del menu que se va a seleccionar en MX, debe ser una lista que contine en orden los ID de la pestaña a la que se debe ingresar (Se parametriza desde 'Config.py')
        
        verbose : `Boolean`
            mostrar información del proceso por pantalla. (default ``False``)
            
        --------------------
        Parametros de Salida
        --------------------
            return validador: `Boolean`
                    Indica si encontro y accedio al menu
        """
        try:
            validador = True
            for selected in menu:
                sleep(0.5)
                menu_button = self.click_until_interactable(By.ID, selected)
            if verbose == True:
                self.logs += "Menu encontrado\n"
                self.logs += "--------------------------------------------------------\n"
            sleep(1)
            return   validador  
                
        # Error al no encontrar el elemento buscado por Selenium
        except TimeoutException:
            self.errors += "Error: \n-Web_Uploader_Class \n-Metodo: log_menu \nDescripción: No se ha podido acceder al botón indicado\n"
            validador = False
            return validador
     
    #---------------------------------------------------------------------------------------------------   
    def Crear_OT(self, OT_actual, IDs, verbose_crearOT = False):
        #Validar si se requiere crear la OT
        if OT_actual.OT_existente == False:
            elemento_crearOT = self.click_until_interactable(By.ID, IDs["Buscadores"]["OTs"]["Crear_OT"]) #Click en crear OT
            #Se validan los erorres que se pueden generar al crear una nueva OT
            OT_actual.Validar_Error_cambiar_OT(IDs, False)
           
            if verbose_crearOT == True:
                self.logs += f"\n--------------------------------------------------------\n"
                self.logs += f"Se crea exitosamente la OT {OT_actual.OT} del ID {OT_actual.ID}\n"
                
        #Si el excel contiene un numero de OT, se busca dicha OT en MX
        else:
            self.buscar_valor(OT_actual, IDs, enter_key=True)
            #Cambiar a ventana orden de trabajo para poder evaluar el estado de la OT
            if OT_actual.Ventana_actual != "Orden_trabajo":
                OT_actual.Cambiar_ventana("Orden_trabajo", IDs)
            OT_actual.Evaluar_Estado_OT(IDs)
            if verbose_crearOT == True:
                self.logs += f"\n--------------------------------------------------------\n"
                self.logs += f"La OT {OT_actual.OT} con ID {OT_actual.ID} existe y fue encontrada con éxito\n"
     
    #---------------------------------------------------------------------------------------------------
    def buscar_valor(self, OT_object, IDs, enter_key = False, verbose_buscar=True):
        """
        Este metodo permite buscar un valor dentro del buscador de OT de MX y presionar enter si se requiere
        ---------------------
        Parametros de Entrada
        ---------------------     
        valor : `text`
            Número que se desea buscar (Ejmplo una OT en MX)
            
        IDs : `json`
                IDs del buscador en la pagina web de tipo ID_Config.json, ejemplo:IDs["Buscadores"]["OTs"]
                
        enter_key: `Boolean`
                Indica si desea que se haga key_Enter en el campo tras rellenarlo
        
        verbose : `Boolean`
            mostrar información del proceso por pantalla. (default ``True``)
        
        --------------------
        Parametros de Salida
        --------------------
            return validador: `Boolean`
                    Indica si el buscador encontro la OT
        """
        try:
            # Encontrar el elemento que permite buscar OT en MX
            buscador = self.click_until_interactable(By.ID, IDs["Buscadores"]["OTs"]["Buscar"])
            # Ingresar la OT en MX
            buscador.send_keys(str(OT_object.OT))   
            # Press enter key
            if enter_key == True:
                buscador.send_keys('\ue007')
                #Se validan los errores que se pueden generar al buscar una nueva OT
                OT_object.Validar_Error_cambiar_OT(IDs, False)
            # No press enter key and return      
            else:
                return
        
        # Error al no encontrar el elemento buscado por Selenium
        except TimeoutException:
            OT_object.ERROR_OT = f"'Error en: \n-WebUploader_Class \n-Metodo: buscar_valor \nDescripción: La OT '{OT_object.OT}' con ID '{OT_object.ID}' no fue posible encontrarla, error del BOT al usar el buscador, id no encontrado\n"
            raise Exception(OT_object.ERROR_OT)
                      
    #---------------------------------------------------------------------------------------------------
    def click_until_interactable(self, buscador, button_id, click = True, tiempo_wait=tiempo_wait_driver):
        sleep(0.3)
        wait = WebDriverWait(self.driver, tiempo_wait)
        element = wait.until(EC.visibility_of_element_located((buscador, button_id)))
        element_is_interactable = False
        star_time = time()
        counter = 1
        if element.is_displayed():
            while not element_is_interactable and not (time()-star_time)>tiempo_wait:
                try:
                    if click==True:
                        element.click()
                    element_is_interactable = True
                    sleep(0.1)
                except (ElementNotInteractableException, ElementClickInterceptedException, TimeoutException, StaleElementReferenceException):
                 element = self.driver.find_element(buscador, button_id)
                 counter = counter + 1
                 sleep(1)   
        return element
    
    
    def configurar_ventana(self, pantalla, tipo):
        # Obtener las dimensiones de la ventana del navegador actual
        ventana_actual = self.driver.get_window_rect()
        ventana_actual_width = ventana_actual['width']
        ventana_actual_height = ventana_actual['height']
    
        # Obtener las dimensiones de la pantalla objetivo
        pantalla_width = self.driver.execute_script("return window.screen.availWidth;")
        pantalla_height = self.driver.execute_script("return window.screen.availHeight;")
    
        # Calcular la nueva posición y tamaño de la ventana
        nueva_posicion_x = pantalla_width if pantalla == 1 else 0
        nueva_posicion_y = 0
        nueva_ventana_width = ventana_actual_width
        nueva_ventana_height = ventana_actual_height
    
        # Configurar la posición y tamaño de la ventana del navegador
        self.driver.set_window_rect(nueva_posicion_x, nueva_posicion_y, nueva_ventana_width, nueva_ventana_height)
    
        # Configurar el estado de la ventana
        if tipo == 'min':
            self.driver.minimize_window()
        elif tipo == 'max':
            self.driver.maximize_window()
    
 
#---------------------------------------------------------------------------------------------------
#---------------------------------------------------------------------------------------------------
# Metodos para operar dentro de la OT
class Orden_Trabajo_Class():
    """
    Clase para gestionar la creación y manipulación de Órdenes de Trabajo (OT) de IBM Maximo

    Parameters
    ----------
    WebUploader_object : WebUploader_Class
        Instancia de la clase WebUploader para operaciones de webscraping.
    DF_TAREAS_init : pandas.DataFrame
        DataFrame con información inicial de tareas.
    DF_MANO_DE_OBRA_init : pandas.DataFrame
        DataFrame con información inicial de mano de obra.
    DF_MATERIALES_init : pandas.DataFrame
        DataFrame con información inicial de materiales.
    DF_SERVICIOS_init : pandas.DataFrame
        DataFrame con información inicial de servicios.
    Data_OT : dict
        Diccionario con datos iniciales de la OT.
    """
    #Campos de la OT
    def __init__(self, WebUploader_object, DF_TAREAS_init, DF_MANO_DE_OBRA_init, DF_MATERIALES_init, DF_SERVICIOS_init, Data_OT):
        self.WebUploader: WebUploader_Class = WebUploader_object 
        self.ID: int = self.Convertir_a_tipo(Data_OT['ID'], int, error_print=['OT','ID'])
        self.OT: int = self.Convertir_a_tipo(Data_OT['OT'], int, error_print=['OT','OT'])
        self.DESCRIPCION: str = self.Convertir_a_tipo(Data_OT['DESCRIPCION'], str, error_print=['OT','DESCRIPCION'])
        self.UBICACION: str = self.Convertir_a_tipo(Data_OT['UBICACION'], str, error_print=['OT','UBICACION'])
        self.OT_PADRE: str = self.Convertir_a_tipo(Data_OT['OT_PADRE'], str, error_print=['OT','OT_PADRE'])
        self.CLASIFICACION: str = self.Convertir_a_tipo(Data_OT['CLASIFICACION'], str, error_print=['OT','CLASIFICACION'])
        self.TIPO_TRABAJO: str = self.Convertir_a_tipo(Data_OT['TIPO_TRABAJO'], str, error_print=['OT','TIPO_TRABAJO'])
        self.TIPO_PROYECTO: list = self.Convertir_a_tipo(Data_OT['TIPO_PROYECTO'], list, [str, str], error_print=['OT','TIPO_PROYECTO'])
        self.RESIDUOS: bool = self.Convertir_a_tipo(Data_OT['RESIDUOS'], bool, error_print=['OT','RESIDUOS'])
        self.HEREDA: bool = self.Convertir_a_tipo(Data_OT['HEREDA'], bool, error_print=['OT','HEREDA'])
        self.PRIORIDAD: str = self.Convertir_a_tipo(Data_OT['PRIORIDAD'], str, error_print=['OT','PRIORIDAD'])
        self.GROT: str = self.Convertir_a_tipo(Data_OT['GROT'], str, error_print=['OT','GROT'])
        self.FSE: str = self.Convertir_a_tipo(Data_OT['FSE'], str, error_print=['OT','FSE'])
        self.ACCION: str = self.Convertir_a_tipo(Data_OT['ACCION'], str, error_print=['OT','ACCION'])
        self.UNIDAD_NEGOCIO: str = self.Convertir_a_tipo(Data_OT['UNIDAD_NEGOCIO'], str, error_print=['OT','UNIDAD_NEGOCIO'])
        self.ACTIVIDAD_COSTEO: str = self.Convertir_a_tipo(Data_OT['ACTIVIDAD_COSTEO'], str, error_print=['OT','ACTIVIDAD_COSTEO'])
        self.INICIO_PREVISTO: str = self.Convertir_a_tipo(Data_OT['INICIO_PREVISTO'], str, error_print=['OT','INICIO_PREVISTO'])
        self.FIN_PREVISTO: str = self.Convertir_a_tipo(Data_OT['FIN_PREVISTO'], str, error_print=['OT','FIN_PREVISTO'])
        self.INICIO_PROGRAMADO: str = self.Convertir_a_tipo(Data_OT['INICIO_PROGRAMADO'], str, error_print=['OT','INICIO_PROGRAMADO'])
        self.FIN_PROGRAMADO: str = self.Convertir_a_tipo(Data_OT['FIN_PROGRAMADO'], str, error_print=['OT','FIN_PROGRAMADO'])
        self.NO_ANTERIOR: str = self.Convertir_a_tipo(Data_OT['INICIO_PREVISTO'], str, error_print=['OT','NO_ANTERIOR'])
        self.MAS_TARDAR: str = self.Convertir_a_tipo(Data_OT['FIN_PREVISTO'], str, error_print=['OT','MAS_TARDAR'])
        self.DURACION: str = self.Convertir_a_tipo(Data_OT['DURACION'], str, error_print=['OT','DURACION'])
        self.PLANIFICADOR: str = self.Convertir_a_tipo(Data_OT['PLANIFICADOR'], str, error_print=['OT','PLANIFICADOR'])
        self.INTERVENTOR: str = self.Convertir_a_tipo(Data_OT['INTERVENTOR'], str, error_print=['OT','INTERVENTOR'])
        self.RESPONSABLE: str = self.Convertir_a_tipo(Data_OT['RESPONSABLE'], str, error_print=['OT','RESPONSABLE'])
        self.CONTRATO: str = self.Convertir_a_tipo(Data_OT['CONTRATO'] , str, error_print=['OT','CONTRATO'])
        self.OT_existente: bool = False if isnull(self.OT) else True
        self.ESTADO = "ESPPLAN"  
        self.Ventana_actual = "Orden_trabajo"
        self.logs = ""
        self.ERROR_OT = ''
        #Instancias de las clases internas
        #Instancia de tipo tarea
        Listado_Tareas = []
        for row, Datos in DF_TAREAS_init.iterrows():
            tarea_actual = self.Tareas_class(self, Datos)              
            Listado_Tareas.append(tarea_actual)    
        
        self.Listado_Tareas: self.Tareas_class = Listado_Tareas
        
        #Instancia de tipo mano de obra
        Listado_Mano_de_obra = []
        for row, Datos in DF_MANO_DE_OBRA_init.iterrows():
            mano_de_obra_actual = self.Mano_de_obra_class(self, Datos)              
            Listado_Mano_de_obra.append(mano_de_obra_actual)    
        
        self.Listado_Mano_de_obra: self.Mano_de_obra_class = Listado_Mano_de_obra
        
        #Instancia de tipo materiales
        Listado_Materiales = []
        for row, Datos in DF_MATERIALES_init.iterrows():
            materiales_actual = self.Materiales_class(self, Datos)              
            Listado_Materiales.append(materiales_actual)    
        
        self.Listado_Materiales: self.Materiales_class = Listado_Materiales
        
        
        #Instancia de tipo servicios
        Listado_Servicios = []
        for row, Datos in DF_SERVICIOS_init.iterrows():
            servicios_actual = self.Servicios_class(self, Datos)              
            Listado_Servicios.append(servicios_actual)    
        
        self.Listado_Servicios: self.Servicios_class = Listado_Servicios
        
        
        #Instancia de tipo asignaciones
        Listado_Asignaciones = []
        for row, Datos in DF_MANO_DE_OBRA_init.iterrows():
            asignaciones_actual = self.Asignaciones_class(self, Datos)              
            Listado_Asignaciones.append(asignaciones_actual)    
        
        self.Listado_Asignaciones: self.Asignaciones_class = Listado_Asignaciones
        
 
# =============================================================================
    #%% Metodos de las clase orden de trabajo        
    #---------------------------------------------------------------------------------------------------
    def guardarOT(self, IDs, verbose=False):
        """
        Función para guarda la Orden de Trabajo (OT) en IBM MAXIMO.

        Parameters
        ----------
        IDs : dict
            Diccionario que contiene los IDs necesarios para interactuar con la interfaz de usuario de MAXIMO.
        verbose : bool, optional
            Indica si se deben imprimir mensajes detallados durante la ejecución (predeterminado False).

        Raises
        ------
        Exception
            Exception 1: Se genera una excepción si no se puede guardar la OT debido a un error de diligenciamiento en la OT
            Exception 2: Se genera una exepción si el BOT no puede encontrar el boton de guardar con el ID especificado

        Returns
        -------
        None
        """
        try:            
        #El boton de guardar en MX es dinamico, se requiere buscar el boton con los dos ID posibles
            try:
                elementoguardar = self.WebUploader.click_until_interactable(By.ID, IDs["Guardar"]["boton_guardar1"])
                try:   
                    #En caso de no poder guardar la OT, aparecera una ventana de error en MX
                    ventana_aceptar = self.WebUploader.click_until_interactable(By.ID, IDs["Guardar"]["ventana_aceptar"], tiempo_wait=5)
                   
                    #Exception 1: Debido a un error de diligenciamiento en la OT
                    self.ERROR_OT = f"'Error en: \n-Orden_Trabajo_Class \n-Metodo: guardarOT \nDescripción: La OT '{self.OT}' con ID '{self.ID}' no se guardo por algun error en los datos cargados en MX\n"
                    raise Exception(self.ERROR_OT)
                #Si la ventana de error no aparece, quiere decir que la OT se guardo correctamente
                except TimeoutException:
                     if verbose:
                         self.logs += "La OT con ID: {self.ID} se guarda con el número de OT: {self.OT}\n"
            #En caso de no encontrar el primer ID del boton guardar, se intenta con el id2
            except NoSuchElementException:
                elementoguardar = self.WebUploader.click_until_interactable(By.ID, IDs["Guardar"]["boton_guardar2"])
                try:   
                    #En caso de no poder guardar la OT, aparecera una ventana de error en MX
                    ventana_aceptar = self.WebUploader.click_until_interactable(By.ID, IDs["Guardar"]["ventana_aceptar"], tiempo_wait=5)
                    
                    #Exception 1: Debido a un error de diligenciamiento en la OT
                    self.ERROR_OT = f"'Error en: \n-Orden_Trabajo_Class \n-Metodo: guardarOT \nDescripción: La OT '{self.OT}' con ID '{self.ID}' no se guardo por algun error en los datos cargados en MX\n"
                    raise Exception(self.ERROR_OT)
                #Si la ventana de error no aparece, quiere decir que la OT se guardo correctamente
                except TimeoutException:
                     if verbose:
                         self.logs += "La OT con ID: {self.ID} se guarda con el número de OT: {self.OT}\n"
                       
        except:
            #Exception 2: El BOT no puede encontrar el boton de guardar con el ID especificado
            self.ERROR_OT = f"'Error en: \n-Orden_Trabajo_Class \n-Metodo: guardarOT \nDescripción: La OT {self.OT} con ID {self.ID} no fue posible guardar, el BOT no encontro el boton de guardar\n"
            raise Exception(self.ERROR_OT)
        
    #---------------------------------------------------------------------------------------------------
    def Validar_Error_cambiar_OT(self, IDs, verbose_EG_OT=True):
        """
        Validar y manejar errores al cambiar de Órdenes de Trabajo (OT) en IBM MAXIMO.

        Parameters
        ----------
        IDs : dict
            Diccionario que contiene los IDs necesarios para interactuar con la interfaz de usuario de MAXIMO.
        verbose_EG_OT : bool, optional
            Indica si se deben imprimir mensajes detallados durante la ejecución (predeterminado True).

        Raises
        ------
        Exception
            Exception 1: Se genera una excepción si la OT buscada no existe, este caso se presenta al intentar trabjar con una OT existente 
            
        Returns
        -------
        None.
        """
        try:                  
            #En caso de haber un error  en la OT anteriormente ejecutada, al buscar otra OT aparecera un ventana emergente, se debe pulsar "NO"
            ventana_no_guardar= self.WebUploader.click_until_interactable(By.ID, IDs["Orden_trabajo"]["Guardar"]["No_guardar"], tiempo_wait=5)
            try:
                #En caso de que la OT actual no exista, aparece un mensaje de error que se debe click en "Aceptar"
                ventana_no_guardar= self.WebUploader.click_until_interactable(By.ID, IDs["Buscadores"]["OTs"]["OT_No_Existe"], tiempo_wait=5)
                
                #Except 1: Error por que la OT buscada no existe
                self.ERROR = f"'Error en: \n-WebUploader_Class \n-Metodo: buscar_valor \nDescripción: La OT {self.OT} con ID {self.ID} no existe en MX\n"
                raise Exception(self.ERROR)
            
            except TimeoutException:
                # Buscar la OT, para validar si la pudo encontrar
                OT_encontrada = self.Capturar_Numero_OT(IDs=IDs["Buscadores"]["OTs"], verbose=verbose_EG_OT)
                    
        except TimeoutException:
            try:
                #En caso de que la OT actual no exista, aparece un mensaje de error que se debe click en "Aceptar"
                ventana_no_guardar= self.WebUploader.click_until_interactable(By.ID, IDs["Buscadores"]["OTs"]["OT_No_Existe"], tiempo_wait=5)
                
                #Except 1: Error por que la OT buscada no existe
                self.ERROR = f"'Error en: \n-WebUploader_Class \n-Metodo: buscar_valor \nDescripción: La OT {self.OT} con ID {self.ID} no existe en MX\n"
                raise Exception(self.ERROR)
            
            except TimeoutException:
                # Buscar la OT, para validar si la pudo encontrar
                OT_encontrada = self.Capturar_Numero_OT(IDs=IDs["Buscadores"]["OTs"], verbose=verbose_EG_OT)    
                   
    
    #---------------------------------------------------------------------------------------------------
    def Ingresar_Datos(self, columnas, IDs, Object_OT, instancia_interna = None, verbose=False):
        """
        Ingresa datos en la interfaz de usuario de IBM MAXIMO para una Orden de Trabajo (OT).

        Parameters
        ----------
        columnas : list
            Lista de columnas o atributos que se utilizarán para ingresar datos en MAXIMO.
        IDs : dict
            Diccionario que contiene los IDs necesarios para interactuar con la interfaz de usuario de MAXIMO.
        Object_OT : Orden_Trabajo_Class
            Instancia de la clase Orden_Trabajo_Class que representa la OT en la que se ingresarán los datos.
        instancia_interna : Orden_Trabajo_Class, optional
            Instancia interna para manipulación (predeterminado None).
        verbose : bool, optional
            Indica si se deben imprimir mensajes detallados durante la ejecución (predeterminado False).

        Raises
        ------
        Exception
            Exception 1: Se genera una excepción si el dato ingresado a MX genera un error dentro de la web, eso pasa porque se captura la ventana de error de MX y advierte sobre un error del dato ingresado. Por tanto, no es un error del BOT, se trata de un error de diligenciamiento, puesto que el dato que se intenta ingresar en la celda no es aceptado por MX
            Exception 2: Error de estado, solo se puede trabajar con OTs en estado ESPPLAN
        
        Returns
        -------
        None.
            
        """
        if Object_OT.ESTADO == "ESPPLAN":
            #Manipulación de la instancia interna, si aplica
            if instancia_interna is None:  
                instancia_interna = Object_OT
            # Proceso iterativo de diligenicamiento de la información 
            if verbose:
                self.logs +=f"\n--------------------------------------------------------\n"
                self.logs += f"    OT con ID {Object_OT.ID}, OT MX {Object_OT.OT}\n"
            for row in columnas:
                #Captura del parametro a diligenciar
                valor_celda = getattr(instancia_interna, row)
                    
                #Si el dato a ingresar es diferente a NA (Celda con dato en el excel), se ingresa
                if valor_celda is not None:
                    #Si el dato es tipo bool, se da click dependiendo si esta marcado o no
                    if isinstance(valor_celda, bool):
                        if valor_celda==False:
                            #Se da click para desactivar el checkbox
                            elemento = self.WebUploader.click_until_interactable(By.ID, IDs[row])
                    
                    #Si el dato es tipo texto o int, se typea el dato
                    elif isinstance(valor_celda, str) or isinstance(valor_celda, int):
                        if isnull(valor_celda)==False:
                            elemento = self.WebUploader.click_until_interactable(By.ID, IDs[row])
                            elemento.clear()
                            elemento.send_keys((str(valor_celda) + Keys.TAB))
                            
                    #Si el dato es tipo list, se recorre la lista
                    elif isinstance(valor_celda, list):
                        for i in valor_celda:
                            elemento = self.WebUploader.click_until_interactable(By.ID, IDs[row][i])
                    
                    #----------------------------------------------------------------
                    #En caso de haber un error al cargar el dato en MX, se identifica
                    try:
                        self.WebUploader.click_until_interactable(By.ID, IDs["Error_Dato"], click=False, tiempo_wait=2.5)
                        
                        #Exception 1: Debido a un error de diligenciamiento en la OT
                        self.ERROR_OT = f"Error en: \n-Orden_Trabajo_Class \n-Metodo: Ingresar_Datos \nDescripción: La OT con ID {Object_OT.ID} en la ventana {Object_OT.Ventana_actual} presenta un error en la celda {row} al ingresarlo a MX, validar el dato que desea cargar\n"
                        raise Exception(self.ERROR_OT)
                    except TimeoutException:
                        if verbose:
                            self.logs +=f"        {Object_OT.Ventana_actual}.{row}: {valor_celda}\n"
                              
                elif verbose:
                    self.logs +=f"        {Object_OT.Ventana_actual}.{row}:  ''\n"
        else:
            #Exception 2: Error de estado, solo se permiten OTs en ESPPLAN 
            self.ERROR_OT = f"'Error en: \n-Orden_Trabajo_Class \n-Metodo: Ingresar_Datos \nDescripción: La OT {Object_OT.OT} con ID {Object_OT.ID} no se encuentra en estado 'ESSPLAN', no es posible modificar\n"
            raise Exception(self.ERROR_OT)
        
    #---------------------------------------------------------------------------------------------------        
    def Cambiar_ventana(self, ventana, IDs):
        """
        Función para cambiar de ventana dentro de la orden de trabajo

        Parameters
        ----------
        ventana : str
            ventana a la que se va a accer, se puede poner lo siguiente:
                -"Orden_trabajo"
                -"Planes"
                    -"mano_de_obra"
                    -"servicios"
                -"Asiganciones"
        IDs : json
            IDs_Config.

        Returns
        -------
        None
            La función no retorna un valor explícito, pero modifica el estado interno (self.Ventana_actual).
        """
        self.WebUploader.click_until_interactable(By.ID, IDs[ventana]["VENTANA"])
        self.Ventana_actual = ventana
        sleep(0.5)


    #---------------------------------------------------------------------------------------------------
    def Capturar_Numero_OT(self, IDs, verbose=False):
        """
        Captura el número de la Orden de Trabajo (OT) en IBM MAXIMO.

        Parameters
        ----------
        IDs : dict
            Diccionario que contiene los IDs necesarios para interactuar con la interfaz de usuario de MAXIMO.
        verbose : bool, optional
            Indica si se deben imprimir mensajes detallados durante la ejecución (predeterminado False).

        Raises
        ------
        Exception
            Exception 1: Se genera una excepción si la OT buscada en MX no coincide con la parametrizada
            Exception 2: Se genera una excepción si no se encuentra el ID del boton especificado
        
        Returns
        -------
        None
            La función no retorna un valor explícito, pero modifica el estado interno (self.OT).
        """
        try:
            elemento = self.WebUploader.click_until_interactable(By.ID, IDs["Validador"][self.Ventana_actual])
            informacion = elemento.get_attribute("value")
            #Si la OT existe y el número en excel coincide con el mostrado en MX, o si la OT no existe
            if (informacion==str(self.OT) and self.OT_existente) or not (self.OT_existente):
                self.OT = int(informacion)
                if verbose:
                    self.logs +="\n--------------------------------------------------------\n"
                    self.logs +=f"*OT MX: {informacion}\n"
                  
            else:
                #Exception 1: OT buscada no coincide con la parametrizada
                self.ERROR_OT = f"'Error en: \n-WebUploader_Class \n-Metodo: Capturar_Numero_OT \nDescripción: La OT con ID {self.ID}, no coincide con la encontrada en MX. {self.OT} =! {int(informacion)}\n"
                raise Exception(self.ERROR_OT)
                 
        except:
            #Exception 2: Error de ID, no se encontro el boton
            self.ERROR_OT = f"'Error en: \n-WebUploader_Class \n-Metodo: Capturar_Numero_OT \nDescripción: En la OT con ID {self.ID} no fue posible caputar el numero de OT en MX, el BOT no encontro el id especificado\n"
            raise Exception(self.ERROR_OT)
            
    #---------------------------------------------------------------------------------------------------
    def Convertir_a_tipo(self, parametro, tipo, tipo_lista=[], error_print=[None, None], OT_object=None): 
        """
        Convierte un parámetro al tipo de dato especificado.

        Parameters
        ----------
        parametro : any
            El parámetro que se desea convertir.
        tipo : type
            El tipo de dato al que se desea convertir el parámetro (str, int, float, bool, list).
        tipo_lista : list, optional
            Lista de tipos para los elementos de una lista (predeterminado []).
        error_print : list, optional
            Lista que contiene información para imprimir en caso de error (predeterminado [None, None]).
        OT_object : Orden_Trabajo_Class, optional
            Instancia de la clase Orden_Trabajo_Class para manejar errores en el contexto de una OT (predeterminado None).

        Returns
        -------
        parametro
            El parámetro convertido al tipo de dato especificado.

        Raises
        ------
        Exception
            Exception 1: Se genera una excepción si ocurre un error al convertir el parámetro.
        """
        try:
            if isnull(parametro): #En caso de tener celda vacia como input
                return        
            if tipo == str:
                return str(parametro).replace('\xa0', ' ')
            elif tipo == int:
                return int(parametro)
            elif tipo == float:
                return float(parametro)
            elif tipo == bool:
                return bool(int(parametro))
            elif tipo == list:
                parametro = eval(parametro)
                for i, tipo_i in enumerate(tipo_lista):
                    parametro[i] = self.Convertir_a_tipo(parametro[i], tipo_i)
                return parametro
        except ValueError:
            if OT_object is not None:
                # Exception 1: Error de lectura de datos
                OT_object.ERROR_OT = f"Error en: \n-WebUploader_Class \n-Metodo: _init_Convertir_a_tipo \nDescripción: Error de lectura de datos en la OT con ID {self.ID}, en la clase {error_print[0]}, en el parametro {error_print[1]}\n"
                raise Exception(OT_object.ERROR_OT)
            else:
                # Exception 1: Error de lectura de datos
                self.ERROR_OT = f"Error en: \n-WebUploader_Class \n-Metodo: _init_Convertir_a_tipo \nDescripción: Error de lectura de datos en la OT con ID {self.ID}, en la clase {error_print[0]}, en el parametro {error_print[1]}\n"
                raise Exception(self.ERROR_OT)

            
    def Fila_Nueva (self, IDs):
        """
        Crea una nueva fila en la interfaz de usuario para agregar información a una Orden de Trabajo (OT).

        Parameters
        ----------
        IDs : dict
            Diccionario que contiene los IDs necesarios para interactuar con la interfaz de usuario de MAXIMO.

        Raises
        ------
        Exception
            Se genera una excepción si ocurre un error al no encontrar el ID
        """
        try:
            fila = self.WebUploader.click_until_interactable(By.ID, IDs["Fila_Nueva"]) 
            sleep(0.5)
            
        except:
             self.ERROR_OT = "Error en: \n-Orden_Trabajo_Class \n-Metodo: Fila_Nueva \nDescripción: Error al encontro id\n"
             raise Exception(self.ERROR_OT)
             
    def Eliminar_Filas(self, tipo, IDs, labor_bot="MODIFICAR"):
        """
        Elimina todas las filas de una sección especifica de una Orden de Trabajo (OT).

        Parameters
        ----------
        tipo : str
            Tipo de filas a eliminar (p. ej., "asignaciones", "materiales", "servicios"). Se eliminan todas las filas de dicha sección.
        IDs : dict
            Diccionario que contiene los IDs necesarios para interactuar con la interfaz de usuario de MAXIMO.
        labor_bot : str, optional
            Tipo de labor del bot, puede ser "MODIFICAR" o "CREAR" (predeterminado "MODIFICAR").

        Raises
        ------
        Exception
            Se genera una excepción si ocurre no se encuentra el ID del boton de eliminar
        """
        if (labor_bot != "CREAR") or (labor_bot == "CREAR" and tipo=="asignaciones"):
            try:
                numero=0
                while numero <= 100: 
                    try:
                        id_base = IDs[tipo]["Eliminar_Filas"]
                        id_actual = id_base.replace("[R:0]", "[R:{}]".format(numero))
                        eliminar = self.WebUploader.click_until_interactable(By.ID, id_actual, click=True, tiempo_wait=5)
                        
                    except TimeoutException:
                        #try:
                            #Encontrar el boton que permite pasar de pagina
                            pasar_pagina = self.WebUploader.click_until_interactable(By.ID, IDs[tipo]["Ver_Mas_Filas"], click=False, tiempo_wait=5) 
                            # Obtener el valor del atributo "ev"
                            informacion = pasar_pagina.get_attribute('ev')
                     
                            if informacion=='true':
                                pasar_pagina.click()
                                eliminar = self.WebUploader.click_until_interactable(By.ID, id_actual, click=True, tiempo_wait=5)
                            else:
                                try:
                                    eliminar = self.WebUploader.click_until_interactable(By.ID, id_actual, click=True, tiempo_wait=5)
                                except TimeoutException:
                                    break
                    numero += 1
                    
                #Guardar los cambios tras eliminar las filas
                self.guardarOT(IDs)
            except:
                self.ERROR_OT = "Error en: \n-Orden_Trabajo_Class \n-Metodo: Eliminar_Filas \nDescripción: Error al encontro id\n"
                raise Exception(self.ERROR_OT)
            
    def Flujo_OT (self, IDs, Estado_deseado, listado_flujo, verbose=True):

        self.Evaluar_Estado_OT(IDs)        
        if self.ESTADO == Estado_deseado:  
            if verbose==True:
                self.logs +=f"La OT {self.OT} con ID {self.ID} ya se encuentra en estado: {Estado_deseado}\n"
        
        else:
            #Verificar si el estado esta parametrizado en el bot
            if (Estado_deseado == 'ASIGNADA') or (Estado_deseado == 'ESPPROG'):
                if self.ESTADO =='ESPPLAN':     
                    conta_errores = 0                   #Variable para validar la cantidad de veces que se presenta error al darle flujo
                    tam_estados = len(listado_flujo)    #size del listado de estados
                    conta1 = 0                          #Variable de control, para asegurar que en                                                         
                    for estado in listado_flujo:
                        try:
                            conta1 = conta1 + 1
                            conta2 = 0
                            for id_i in IDs["flujo_OT"][Estado_deseado][estado]:
                                conta2 = conta2 + 1
                                elemento = self.WebUploader.click_until_interactable(By.ID, id_i, tiempo_wait=90) 
                            
                            #Evaluar el estado de la OT
                            self.Evaluar_Estado_OT(IDs)
                            if self.ESTADO == Estado_deseado:
                                break
                        except:
                            conta_errores = conta_errores + 1
                            if (conta_errores>1) | (conta1==tam_estados):
                                self.ERROR_OT = f"Error en: \n-Orden_Trabajo_Class \n-Metodo: Flujo_OT \nDescripción: La OT con ID {self.ID} presenta un error al darle flujo a la OT, no se encontro el id\n"
                                raise Exception(self.ERROR_OT)
                            else:
                                continue
                
                    if verbose==True:
                        self.logs +=f"Flujo correcto de la OT {self.OT}: con ID {self.ID}, estado final: {self.ESTADO}\n"
                else:
                    self.ERROR_OT = f"Error en: \n-Orden_Trabajo_Class \n-Metodo: Flujo_OT \nDescripción: La OT con ID {self.ID} no se encuentra en estado 'ESPPLAN'\n"
                    raise Exception(self.ERROR_OT)
            else: # Acción por defecto si el estado no coincide con ninguna condición parametrizada en el bot
                self.ERROR_OT = f"Error en: \n-Orden_Trabajo_Class \n-Metodo: Flujo_OT \nDescripción: La OT con ID {self.ID} presenta un error al darle flujo a la OT, el estado deseado no esta parametrizado en el BOT\n"
                raise Exception(self.ERROR_OT)

    def Evaluar_Estado_OT (self, IDs):
        try:
            sleep(1)
            informacion = self.WebUploader.click_until_interactable(By.ID, IDs["Orden_trabajo"]["Estado_MX"], click=True) 
            # Obtener el valor del atributo "ev"
            self.ESTADO = informacion.get_attribute('value')
        except:
             self.ERROR_OT = (f"Error en: \n-WebUploader_Class \n-Metodo: Evaluar_Estado_OT \nDescripción: No se encontro el estado en MX de la OT: {self.OT}\n")
             raise Exception(self.ERROR_OT) 
          
    def Cancelar_OT(self, IDs, Estado_deseado, listado_flujo, verbose=True):
        #Verificar si el estado esta parametrizado en el bot
        if Estado_deseado == 'CANCELADA':
            if self.ESTADO =='ESPPLAN':                                                                                   
                for estado in listado_flujo:
                    try:
                        conta = 0
                        for id_i in IDs["flujo_OT"][Estado_deseado][estado]:
                            conta = conta + 1
                            elemento = self.WebUploader.click_until_interactable(By.ID, id_i, tiempo_wait=60) 
                        
                        #Evaluar el estado de la OT
                        self.Evaluar_Estado_OT(IDs)
                        if self.ESTADO == Estado_deseado:
                            break
                    except:
                        self.ERROR_OT = f"Error en: \n-Orden_Trabajo_Class \n-Metodo: Flujo_OT \nDescripción: La OT con ID {self.ID} presenta un error al darle flujo a la OT, no se encontro el id\n"
                        raise Exception(self.ERROR_OT)
            
                if verbose==True:
                    self.logs +=f"Flujo correcto de la OT {self.OT}: con ID {self.ID}, estado final: {self.ESTADO}\n"
            else:
                self.ERROR_OT = f"Error en: \n-Orden_Trabajo_Class \n-Metodo: Flujo_OT \nDescripción: La OT con ID {self.ID} no se encuentra en estado 'ESPPLAN'\n"
                raise Exception(self.ERROR_OT)
        else: # Acción por defecto si el estado no coincide con ninguna condición parametrizada en el bot
            self.ERROR_OT = f"Error en: \n-Orden_Trabajo_Class \n-Metodo: Flujo_OT \nDescripción: La OT con ID {self.ID} presenta un error al darle flujo a la OT, el estado deseado no esta parametrizado en el BOT\n"
            raise Exception(self.ERROR_OT)
        
    
    #%% Clases internas de la clase OT
    class Tareas_class():
        def __init__(self, OT_actual, Data):
            self.ID: int = OT_actual.Convertir_a_tipo(Data['ID'], int, error_print=['TAREAS','ID'], OT_object = OT_actual)
            self.TAREA: int = OT_actual.Convertir_a_tipo(Data['TAREA'], int, error_print=['TAREAS','TAREA'], OT_object = OT_actual)
            self.RESUMEN: int = OT_actual.Convertir_a_tipo(Data['RESUMEN'], str, error_print=['TAREAS','RESUMEN'], OT_object = OT_actual) 
               
    class Servicios_class():
        def __init__(self, OT_actual, Data):
            self.ID: int = OT_actual.Convertir_a_tipo(Data['ID'], int, error_print=['SERVICIOS','ID'], OT_object = OT_actual)
            self.TAREA: int = OT_actual.Convertir_a_tipo(Data['TAREA'], int, error_print=['SERVICIOS','TAREA'], OT_object = OT_actual)
            self.TIPO_LINEA: list = OT_actual.Convertir_a_tipo(Data['TIPO_LINEA'], list, [str, str], error_print=['SERVICIOS','TIPO_LINEA'], OT_object = OT_actual)
            self.SERVICIO: str = OT_actual.Convertir_a_tipo(Data['SERVICIO'], str, error_print=['SERVICIOS','SERVICIO'], OT_object = OT_actual)
            self.CANTIDAD: int = OT_actual.Convertir_a_tipo(Data['CANTIDAD'], int, error_print=['SERVICIOS','CANTIDAD'], OT_object = OT_actual)
    
    class Mano_de_obra_class():
        def __init__(self, OT_actual, Data):
            self.ID: int = OT_actual.Convertir_a_tipo(Data['ID'], int, error_print=['MO','ID'], OT_object = OT_actual)
            self.TAREA: int = OT_actual.Convertir_a_tipo(Data['TAREA'], int, error_print=['MO','TAREA'], OT_object = OT_actual)
            self.TIPO_CUADRILLA: str = OT_actual.Convertir_a_tipo(Data['TIPO_CUADRILLA'], str, error_print=['MO','TIPO_CUADRILLA'], OT_object = OT_actual)  
            self.HORA: int = OT_actual.Convertir_a_tipo(Data['HORA'], int, error_print=['MO','HORA'], OT_object = OT_actual)  
            
    class Materiales_class():
        def __init__(self, OT_actual, Data):
            self.ID: int = OT_actual.Convertir_a_tipo(Data['ID'], int, error_print=['MATERIALES','ID'], OT_object = OT_actual)
            self.TAREA: int = OT_actual.Convertir_a_tipo(Data['TAREA'], int, error_print=['MATERIALES','TAREA'], OT_object = OT_actual)
            self.TIPO_LINEA: list = OT_actual.Convertir_a_tipo(Data['TIPO_LINEA'], list, [str, str], error_print=['MATERIALES','TIPO_LINEA'], OT_object = OT_actual)
            self.PARTE: str = OT_actual.Convertir_a_tipo(Data['PARTE'], int, error_print=['MATERIALES','PARTE'], OT_object = OT_actual)
            self.ALMACEN: int = OT_actual.Convertir_a_tipo(Data['ALMACEN'], str, error_print=['MATERIALES','ALMACEN'], OT_object = OT_actual)
            self.CANTIDAD: int = OT_actual.Convertir_a_tipo(Data['CANTIDAD'], int, error_print=['MATERIALES','CANTIDAD'], OT_object = OT_actual) 
               
    class Asignaciones_class():
        def __init__(self, OT_actual, Data):
            self.ID: int = OT_actual.Convertir_a_tipo(Data['ID'], int, error_print=['MO_ASIGNACION','ID'], OT_object = OT_actual)
            self.TAREA: str = OT_actual.Convertir_a_tipo(Data['TAREA'], int, error_print=['MO_ASIGNACION','TAREA'], OT_object = OT_actual)
            self.TIPO_CUADRILLA: str = OT_actual.Convertir_a_tipo(Data['TIPO_CUADRILLA'], str, error_print=['MO_ASIGNACION','TIPO_CUADRILLA'], OT_object = OT_actual)
            self.MANO_OBRA: str = OT_actual.Convertir_a_tipo(Data['MANO_OBRA'], str, error_print=['MO_ASIGNACION','MANO_OBRA'], OT_object = OT_actual)
            self.HORA: int = OT_actual.Convertir_a_tipo(Data['HORA'], int, error_print=['MO_ASIGNACION','HORA'], OT_object = OT_actual)
    
        
                
                
                
                
       