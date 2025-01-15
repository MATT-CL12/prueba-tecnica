
# Explicación del Sistema de Automatización de Órdenes de Trabajo

El sistema consiste en dos archivos principales que trabajan en conjunto para automatizar la carga masiva de órdenes de trabajo en el sistema IBM Maximo. Veamos cómo funciona todo el proceso:

## 1. Archivo Principal (Main)

### Inicialización y Configuración

```python
import pandas as pd
from WebUploader_Class import WebUploader_Class, Orden_Trabajo
from time import time
import datetime
import json
import sys

time_program = {
    "start_program": time(),
    "end_program":0,
    "start_OT":0, 
    "end_OT":0, 
    "list_time_OT":[]
}
```

Este código inicial establece:
- Un sistema de medición de tiempo para monitorear el rendimiento
- Las importaciones necesarias para manejar datos, tiempo y configuraciones
- Un registro del tiempo de inicio del programa

### Sistema de Registro Dual

```python
class DualWriter:
    def __init__(self, original_stdout, archivo):
        self.original_stdout = original_stdout
        self.archivo = archivo
    
    def write(self, mensaje):
        self.original_stdout.write(mensaje)
        self.archivo.write(mensaje)
```

Esta clase permite:
- Escribir simultáneamente en la consola y en un archivo de registro
- Mantener un registro detallado de toda la ejecución
- Crear archivos de registro con marca de tiempo para múltiples ejecuciones

### Configuración del Sistema

```python
with open("Input\\User_Config.json", encoding='utf-8') as file:
    Data_config = json.load(file)
with open("Input\\ID_Config.json", encoding='utf-8') as file:
    IDs = json.load(file)
with open("Input\\Campos_Diligenciar.json", encoding='utf-8') as file:
    Campos_Diligenciar = json.load(file)
```

El sistema carga tres tipos de configuración:
1. Configuración de usuario (credenciales y settings)
2. IDs de elementos web para automatización
3. Definición de campos a completar en las OTs

### Carga de Datos

```python
name = 'Datos_entrada.xlsx'
DF_OTs = pd.read_excel('Input\\' + name, sheet_name="OTs")
DF_TAREAS = pd.read_excel('Input\\' + name, sheet_name="TAREAS")
DF_MANO_DE_OBRA = pd.read_excel('Input\\' + name, sheet_name="MANO_DE_OBRA")
DF_MATERIALES = pd.read_excel('Input\\' + name, sheet_name="MATERIALES")
DF_SERVICIOS = pd.read_excel('Input\\' + name, sheet_name="SERVICIOS")
```

Se cargan los datos desde un Excel estructurado con múltiples hojas que contienen:
- Información general de OTs
- Detalles de tareas
- Asignaciones de mano de obra
- Listados de materiales
- Servicios requeridos

### Proceso Principal

```python
for row, Datos in DF_OTs.iterrows():
    if Datos['LABOR_BOT']!='NADA':
        try:
            time_program["start_OT"] = time()
            
            DF_TAREAS_ACTUAL = DF_TAREAS.loc[DF_TAREAS['ID']==Datos['ID']]
            OT_actual = Orden_Trabajo_Class(WebUploader, DF_TAREAS_ACTUAL, ...)
            
            WebUploader.Crear_OT(OT_actual, IDs, True)
```

Este bucle principal:
- Procesa cada OT en el archivo
- Filtra los datos relacionados
- Crea instancias de OT
- Maneja el proceso de creación/actualización

## 2. Clase WebUploader


---

### Constructor `__init__`
```python
def __init__(self, Data_config, verbose=True):
    brwoser_options = Options()
    brwoser_options.add_argument("--window-position=-1000,300")  # Coordenadas de la segunda pantalla

    self.driver = webdriver.Edge(brwoser_options)      
    self.driver.maximize_window()
    
    self.driver.get(Data_config["url_path"])

    if verbose == True:
        print("Web Driver iniciado correctamente, URL:'", Data_config["url_path"], "' - Lista para ser usada por el bot")
        print("--------------------------------------------------------")
```

**Propósito**: Este constructor configura e inicia el navegador web para realizar tareas automatizadas.  
**Pasos**:  
1. Crea las opciones para el navegador Edge.
2. Define la posición inicial de la ventana en la pantalla.
3. Inicializa el driver de Edge con las opciones configuradas.
4. Maximiza la ventana del navegador.
5. Navega hacia la URL especificada en `Data_config`.
6. Si `verbose` es `True`, imprime un mensaje de confirmación indicando que el navegador está listo.

---

### Función `log_in`
```python
def log_in(self, Data_config, IDs, verbose=True):
    self.username = Data_config["USER"]
    self.password = Data_config["PASS"]
   
    try:           
        username_field = self.click_until_interactable(By.ID, IDs["username"])
        username_field.send_keys(self.username)
        password_field = self.click_until_interactable(By.ID, IDs["password"])
        password_field.send_keys(self.password)
        
        list_button = self.click_until_interactable(By.XPATH, IDs["loginlist"]["list_name"])
        list_selected = self.click_until_interactable(By.XPATH, IDs["loginlist"][Data_config["idioma"]])
        
        submit_button = self.click_until_interactable(By.ID, IDs["loginbutton"])
        name_sesion = self.click_until_interactable(By.ID, IDs["validador"], click=False)
        sesion_started = name_sesion.is_displayed()
        
        if verbose==True:
            print("Sesion iniciada correctamente con la cuenta:", self.username)         
            print("--------------------------------------------------------")
            
        return sesion_started
    
    except TimeoutException:
        print('Error: \n-Web_Uploader_Class \n-Metodo: log_in \nDescripción: No se ha podido iniciar sesión correctamente con la cuenta', self.username)
        sesion_started = False
        return sesion_started
```

**Propósito**: Automatiza el proceso de inicio de sesión en la página web.  
**Pasos**:  
1. Asigna el nombre de usuario y la contraseña desde el diccionario `Data_config`.
2. Encuentra el campo de entrada del usuario y lo completa con las credenciales.
3. Encuentra el campo de entrada de contraseña y lo completa.
4. Selecciona un idioma específico de la lista desplegable.
5. Hace clic en el botón de inicio de sesión.
6. Verifica si la sesión fue iniciada exitosamente comprobando la visibilidad de un elemento validador.
7. Imprime un mensaje de éxito si `verbose` es `True`.
8. Retorna `True` si la sesión fue iniciada exitosamente, o `False` en caso de error.

---

### Función `click_until_interactable`
```python
def click_until_interactable(self, buscador, button_id, click=True, tiempo_wait=tiempo_wait_driver):
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
```

**Propósito**: Localiza un elemento en la página web y realiza clic en él cuando sea interactuable.  
**Pasos**:  
1. Espera a que el elemento sea visible en la página.
2. Verifica si el elemento es interactuable.
3. Si el clic está habilitado, intenta hacer clic en el elemento.
4. Si no es posible, lo reintenta hasta que el tiempo de espera expire.
5. Retorna el elemento interactuado o levanta excepciones si no es posible encontrarlo.

---

### Función `log_menu`
```python
def log_menu(self, menu, verbose=False):
    try:
        validador = True
        for selected in menu:
            sleep(0.5)
            menu_button = self.click_until_interactable(By.ID, selected)
        if verbose == True:
            print("Menu encontrado")
            print("--------------------------------------------------------")
        sleep(1)
        return validador  
                
    except TimeoutException:
        print('Error: \n-Web_Uploader_Class \n-Metodo: log_in \nDescripción: No se ha podido acceder al boton indicado')
        validador = False
        return validador
```

**Propósito**: Selecciona una opción de menú en la interfaz web.  
**Pasos**:  
1. Itera sobre una lista de IDs de botones del menú.
2. Hace clic en cada botón en el orden especificado.
3. Si `verbose` es `True`, imprime un mensaje indicando que el menú fue encontrado.
4. Retorna `True` si todos los clics fueron exitosos, o `False` en caso de error.


---

### Función `cerrar_navegador`
```python
def cerrar_navegador(self):
    self.driver.quit()
```

**Propósito**: Cierra el navegador y finaliza la instancia de Selenium.  
**Pasos**:  
1. Llama al método `quit` del driver para cerrar el navegador y liberar los recursos asociados.

---

### Función `Crear_OT`
```python
def Crear_OT(self, OT_actual, IDs, verbose_crearOT=False):
    if OT_actual.OT_existente == False:
        elemento_crearOT = self.click_until_interactable(By.ID, IDs["Buscadores"]["OTs"]["Crear_OT"])
        OT_actual.Validar_Error_cambiar_OT(IDs, False)
        if verbose_crearOT:
            print("--------------------------------------------------------")
            print(f"Se crea exitosamente la OT {OT_actual.OT} del ID {OT_actual.ID}")
    else:
        self.buscar_valor(OT_actual, IDs, enter_key=True)
        if OT_actual.Ventana_actual != "Orden_trabajo":
            OT_actual.Cambiar_ventana("Orden_trabajo", IDs)
        OT_actual.Evaluar_Estado_OT(IDs)
        if verbose_crearOT:
            print("--------------------------------------------------------")
            print(f"La OT {OT_actual.OT} con ID {OT_actual.ID} existe y fue encontrada con éxito")
```

**Propósito**: Crea una nueva orden de trabajo (OT) o busca una existente en IBM Maximo.  
**Pasos**:  
1. Verifica si la OT actual no existe (`OT_existente == False`).
2. Si no existe:
   - Encuentra y hace clic en el botón para crear una nueva OT.
   - Valida posibles errores al crear la OT.
   - Imprime un mensaje de éxito si `verbose_crearOT` es `True`.
3. Si ya existe:
   - Busca la OT en el sistema usando el número.
   - Cambia a la ventana correspondiente si no está ya en la de "Orden_trabajo".
   - Evalúa el estado actual de la OT.
   - Imprime un mensaje de éxito si `verbose_crearOT` es `True`.

---

### Función `buscar_valor`
```python
def buscar_valor(self, OT_object, IDs, enter_key=False, verbose_buscar=True):
    try:
        buscador = self.click_until_interactable(By.ID, IDs["Buscadores"]["OTs"]["Buscar"])
        buscador.send_keys(str(OT_object.OT))   
        if enter_key:
            buscador.send_keys('\ue007')
            OT_object.Validar_Error_cambiar_OT(IDs, False)
        else:
            return
    except TimeoutException:
        OT_object.ERROR_OT = f"'Error en: \n-WebUploader_Class \n-Metodo: buscar_valor \nDescripción: La OT '{OT_object.OT}' con ID '{OT_object.ID}' no fue posible encontrarla"
        raise Exception(OT_object.ERROR_OT)
```

**Propósito**: Busca un valor (número de OT) en un campo de búsqueda de IBM Maximo.  
**Pasos**:  
1. Localiza el campo de búsqueda por su ID.
2. Ingresa el valor del número de OT en el campo.
3. Si `enter_key` es `True`, simula presionar la tecla "Enter".
4. Valida posibles errores al buscar la OT.
5. Si ocurre un error de tiempo de espera, lanza una excepción personalizada.

---

### Función `configurar_ventana`
```python
def configurar_ventana(self, pantalla, tipo):
    ventana_actual = self.driver.get_window_rect()
    ventana_actual_width = ventana_actual['width']
    ventana_actual_height = ventana_actual['height']
    pantalla_width = self.driver.execute_script("return window.screen.availWidth;")
    pantalla_height = self.driver.execute_script("return window.screen.availHeight;")
    nueva_posicion_x = pantalla_width if pantalla == 1 else 0
    nueva_posicion_y = 0
    nueva_ventana_width = ventana_actual_width
    nueva_ventana_height = ventana_actual_height
    self.driver.set_window_rect(nueva_posicion_x, nueva_posicion_y, nueva_ventana_width, nueva_ventana_height)
    if tipo == 'min':
        self.driver.minimize_window()
    elif tipo == 'max':
        self.driver.maximize_window()
```

**Propósito**: Configura la posición y el estado de la ventana del navegador.  
**Pasos**:  
1. Obtiene las dimensiones actuales de la ventana del navegador.
2. Calcula las dimensiones y posición de la nueva ventana.
3. Ajusta la posición y dimensiones de la ventana.
4. Si `tipo` es `'min'`, minimiza la ventana. Si es `'max'`, la maximiza.

---

### Función `guardarOT`
```python
def guardarOT(self, IDs, verbose=False):
    try:            
        elementoguardar = self.WebUploader.click_until_interactable(By.ID, IDs["Guardar"]["boton_guardar1"])
        try:   
            ventana_aceptar = self.WebUploader.click_until_interactable(By.ID, IDs["Guardar"]["ventana_aceptar"], tiempo_wait=5)
            self.ERROR_OT = f"'Error en: \n-Orden_Trabajo_Class \n-Metodo: guardarOT \nDescripción: La OT '{self.OT}' con ID '{self.ID}' no se guardó por un error en los datos cargados en MX"
            raise Exception(self.ERROR_OT)
        except TimeoutException:
            if verbose:
                print("La OT con ID: ", self.ID, " se guarda con el número de OT:", self.OT)
    except:
        self.ERROR_OT = f"'Error en: \n-Orden_Trabajo_Class \n-Metodo: guardarOT \nDescripción: La OT {self.OT} con ID {self.ID} no fue posible guardar"
        raise Exception(self.ERROR_OT)
```

**Propósito**: Guarda una orden de trabajo (OT) en IBM Maximo.  
**Pasos**:  
1. Intenta encontrar y hacer clic en el botón "Guardar".
2. Si aparece una ventana de error, lanza una excepción indicando que hubo un problema con los datos.
3. Si no hay errores, imprime un mensaje de éxito si `verbose` es `True`.


---

### Función `Validar_Error_cambiar_OT`
```python
def Validar_Error_cambiar_OT(self, IDs, verbose_EG_OT=True):
    try:                  
        ventana_no_guardar = self.WebUploader.click_until_interactable(By.ID, IDs["Orden_trabajo"]["Guardar"]["No_guardar"], tiempo_wait=5)
        try:
            ventana_no_guardar = self.WebUploader.click_until_interactable(By.ID, IDs["Buscadores"]["OTs"]["OT_No_Existe"], tiempo_wait=5)
            self.ERROR = f"'Error en: \n-WebUploader_Class \n-Metodo: buscar_valor \nDescripción: La OT {self.OT} con ID {self.ID} no existe en MX"
            raise Exception(self.ERROR)
        except TimeoutException:
            OT_encontrada = self.Capturar_Numero_OT(IDs=IDs["Buscadores"]["OTs"], verbose=verbose_EG_OT)
    except TimeoutException:
        try:
            ventana_no_guardar = self.WebUploader.click_until_interactable(By.ID, IDs["Buscadores"]["OTs"]["OT_No_Existe"], tiempo_wait=5)
            self.ERROR = f"'Error en: \n-WebUploader_Class \n-Metodo: buscar_valor \nDescripción: La OT {self.OT} con ID {self.ID} no existe en MX"
            raise Exception(self.ERROR)
        except TimeoutException:
            OT_encontrada = self.Capturar_Numero_OT(IDs=IDs["Buscadores"]["OTs"], verbose=verbose_EG_OT)
```

**Propósito**: Maneja y valida errores al cambiar de órdenes de trabajo en IBM Maximo.  
**Pasos**:  
1. Intenta cerrar una ventana emergente de "No guardar" si existe.
2. Busca una ventana emergente de error indicando que la OT no existe.  
   - Si la encuentra, lanza una excepción personalizada.
3. Si no hay errores, captura el número de OT.
4. Maneja excepciones en caso de errores adicionales.

---

### Función `Ingresar_Datos`
```python
def Ingresar_Datos(self, columnas, IDs, Object_OT, instancia_interna=None, verbose=False):
    if Object_OT.ESTADO == "ESPPLAN":
        if instancia_interna is None:  
            instancia_interna = Object_OT
        if verbose:
            print('--------------------------------------------------------')
            print(f"    OT con ID {Object_OT.ID}, OT MX {Object_OT.OT}")
        for row in columnas:
            valor_celda = getattr(instancia_interna, row)
            if valor_celda is not None:
                if isinstance(valor_celda, bool):
                    if not valor_celda:
                        elemento = self.WebUploader.click_until_interactable(By.ID, IDs[row])
                elif isinstance(valor_celda, (str, int)):
                    if not isnull(valor_celda):
                        elemento = self.WebUploader.click_until_interactable(By.ID, IDs[row])
                        elemento.clear()
                        elemento.send_keys((str(valor_celda) + Keys.TAB))
                elif isinstance(valor_celda, list):
                    for i in valor_celda:
                        elemento = self.WebUploader.click_until_interactable(By.ID, IDs[row][i])
                try:
                    self.WebUploader.click_until_interactable(By.ID, IDs["Error_Dato"], click=False, tiempo_wait=2.5)
                    self.ERROR_OT = f"Error en: \n-Orden_Trabajo_Class \n-Metodo: Ingresar_Datos \nDescripción: Error en la celda {row} al ingresarla en MX"
                    raise Exception(self.ERROR_OT)
                except TimeoutException:
                    if verbose:
                        print(f"        {Object_OT.Ventana_actual}.{row}: {valor_celda}")
            elif verbose:
                print(f"        {Object_OT.Ventana_actual}.{row}:  ''")
    else:
        self.ERROR_OT = f"'Error en: \n-Orden_Trabajo_Class \n-Metodo: Ingresar_Datos \nDescripción: La OT {Object_OT.OT} con ID {Object_OT.ID} no está en estado 'ESPPLAN'"
        raise Exception(self.ERROR_OT)
```

**Propósito**: Ingresa datos en campos específicos de una orden de trabajo en la interfaz de IBM Maximo.  
**Pasos**:  
1. Verifica que el estado de la OT sea "ESPPLAN".
2. Itera sobre las columnas para ingresar los datos.
   - Maneja tipos de datos (`bool`, `str`, `int`, `list`).
3. Intenta detectar errores al ingresar datos y lanza excepciones si los encuentra.
4. Si todo es exitoso, imprime los valores ingresados si `verbose` es `True`.

---

### Función `Cambiar_ventana`
```python
def Cambiar_ventana(self, ventana, IDs):
    self.WebUploader.click_until_interactable(By.ID, IDs[ventana]["VENTANA"])
    self.Ventana_actual = ventana
    sleep(0.5)
```

**Propósito**: Cambia a una ventana específica dentro de la interfaz de la orden de trabajo.  
**Pasos**:  
1. Encuentra y hace clic en el botón correspondiente a la ventana deseada.
2. Actualiza el atributo `Ventana_actual` con la ventana seleccionada.
3. Espera un momento para asegurar la estabilidad de la interfaz.

---

### Función `Capturar_Numero_OT`
```python
def Capturar_Numero_OT(self, IDs, verbose=False):
    try:
        elemento = self.WebUploader.click_until_interactable(By.ID, IDs["Validador"][self.Ventana_actual])
        informacion = elemento.get_attribute("value")
        if (informacion == str(self.OT) and self.OT_existente) or not self.OT_existente:
            self.OT = int(informacion)
            if verbose:
                print("--------------------------------------------------------")
                print('*OT MX: ', informacion) 
        else:
            self.ERROR_OT = f"'Error en: \n-WebUploader_Class \n-Metodo: Capturar_Numero_OT \nDescripción: La OT con ID {self.ID}, no coincide con la encontrada en MX. {self.OT} =! {int(informacion)}"
            raise Exception(self.ERROR_OT)
    except:
        self.ERROR_OT = f"'Error en: \n-WebUploader_Class \n-Metodo: Capturar_Numero_OT \nDescripción: No se pudo capturar el número de OT"
        raise Exception(self.ERROR_OT)
```

**Propósito**: Captura el número de la OT desde la interfaz de usuario.  
**Pasos**:  
1. Busca el elemento validador correspondiente a la ventana actual.
2. Obtiene el valor del elemento y verifica que coincida con la OT esperada.
3. Si hay discrepancias, lanza una excepción personalizada.
4. Imprime el número capturado si `verbose` es `True`.

---

### Función `Fila_Nueva`
```python
def Fila_Nueva(self, IDs):
    try:
        fila = self.WebUploader.click_until_interactable(By.ID, IDs["Fila_Nueva"]) 
        sleep(0.5)
    except:
        self.ERROR_OT = "Error en: \n-Orden_Trabajo_Class \n-Metodo: Fila_Nueva \nDescripción: Error al encontrar el ID"
        raise Exception(self.ERROR_OT)
```

**Propósito**: Agrega una nueva fila en la interfaz de usuario para ingresar datos en una sección específica de la OT.  
**Pasos**:  
1. Busca el botón de "Agregar fila nueva" usando su ID y hace clic en él.
2. Si ocurre un error al localizar el elemento, lanza una excepción personalizada.

---

### Función `Eliminar_Filas`
```python
def Eliminar_Filas(self, tipo, IDs, labor_bot="MODIFICAR"):
    if (labor_bot != "CREAR") or (labor_bot == "CREAR" and tipo == "asignaciones"):
        try:
            numero = 0
            while numero <= 100: 
                try:
                    id_base = IDs[tipo]["Eliminar_Filas"]
                    id_actual = id_base.replace("[R:0]", f"[R:{numero}]")
                    eliminar = self.WebUploader.click_until_interactable(By.ID, id_actual, click=True, tiempo_wait=5)
                except TimeoutException:
                    pasar_pagina = self.WebUploader.click_until_interactable(By.ID, IDs[tipo]["Ver_Mas_Filas"], click=False, tiempo_wait=5)
                    informacion = pasar_pagina.get_attribute('ev')
                    if informacion == 'true':
                        pasar_pagina.click()
                        eliminar = self.WebUploader.click_until_interactable(By.ID, id_actual, click=True, tiempo_wait=5)
                    else:
                        break
                numero += 1
            self.guardarOT(IDs)
        except:
            self.ERROR_OT = "Error en: \n-Orden_Trabajo_Class \n-Metodo: Eliminar_Filas \nDescripción: Error al encontrar el ID"
            raise Exception(self.ERROR_OT)
```

**Propósito**: Elimina todas las filas de una sección específica en la interfaz de la OT.  
**Pasos**:  
1. Itera sobre las filas visibles usando un patrón de ID.
2. Si no se encuentran más filas, intenta pasar a la siguiente página de resultados.
3. Elimina las filas encontradas.
4. Guarda los cambios realizados.
5. Si hay errores al localizar un botón, lanza una excepción personalizada.

---

### Función `Flujo_OT`
```python
def Flujo_OT(self, IDs, Estado_deseado, listado_flujo, verbose=True):
    self.Evaluar_Estado_OT(IDs)
    if self.ESTADO == Estado_deseado:  
        if verbose:
            print(f"La OT {self.OT} con ID {self.ID} ya se encuentra en estado: {Estado_deseado}")
    else:
        if Estado_deseado in ['ASIGNADA', 'ESPPROG', 'ESPPO']:
            if self.ESTADO == 'ESPPLAN':     
                conta_errores = 0
                tam_estados = len(listado_flujo)
                conta1 = 0
                for estado in listado_flujo:
                    try:
                        conta1 += 1
                        for id_i in IDs["flujo_OT"][Estado_deseado][estado]:
                            elemento = self.WebUploader.click_until_interactable(By.ID, id_i, tiempo_wait=90) 
                            sleep(2)
                        self.Evaluar_Estado_OT(IDs)
                        if self.ESTADO == Estado_deseado:
                            break
                    except:
                        conta_errores += 1
                        if (conta_errores > 1) or (conta1 == tam_estados):
                            self.ERROR_OT = f"Error en: \n-Orden_Trabajo_Class \n-Metodo: Flujo_OT \nDescripción: Error al cambiar el estado de la OT"
                            raise Exception(self.ERROR_OT)
            else:
                self.ERROR_OT = f"Error en: \n-Orden_Trabajo_Class \n-Metodo: Flujo_OT \nDescripción: La OT con ID {self.ID} no se encuentra en estado 'ESPPLAN'"
                raise Exception(self.ERROR_OT)
        else:
            self.ERROR_OT = f"Error en: \n-Orden_Trabajo_Class \n-Metodo: Flujo_OT \nDescripción: El estado deseado no está parametrizado en el bot"
            raise Exception(self.ERROR_OT)
```

**Propósito**: Cambia el estado de una OT según un flujo definido.  
**Pasos**:  
1. Verifica si el estado actual ya coincide con el deseado.
2. Si no, ejecuta una serie de clics en botones definidos para cambiar el estado.
3. Evalúa nuevamente el estado tras cada acción.
4. Si no puede completar el flujo, lanza una excepción.

---

### Función `Evaluar_Estado_OT`
```python
def Evaluar_Estado_OT(self, IDs):
    try:
        sleep(1)
        informacion = self.WebUploader.click_until_interactable(By.ID, IDs["Orden_trabajo"]["Estado_MX"], click=True)
        self.ESTADO = informacion.get_attribute('value')
    except:
        self.ERROR_OT = (f"Error en: \n-WebUploader_Class \n-Metodo: Evaluar_Estado_OT \nDescripción: No se encontró el estado en MX de la OT: {self.OT}")
        raise Exception(self.ERROR_OT)
```

**Propósito**: Obtiene el estado actual de una OT desde la interfaz de IBM Maximo.  
**Pasos**:  
1. Localiza el campo que contiene el estado.
2. Extrae el valor del estado y lo guarda en la variable interna `ESTADO`.
3. Si no puede encontrar el elemento, lanza una excepción.

---

### Clases internas (`Tareas_class`, `Servicios_class`, etc.)
Estas clases representan estructuras específicas dentro de las órdenes de trabajo. A continuación, explico una de ellas como ejemplo:

#### `Tareas_class`
```python
class Tareas_class:
    def __init__(self, OT_actual, Data):
        self.ID: int = OT_actual.Convertir_a_tipo(Data['ID'], int, error_print=['TAREAS', 'ID'], OT_object=OT_actual)
        self.TAREA: int = OT_actual.Convertir_a_tipo(Data['TAREA'], int, error_print=['TAREAS', 'TAREA'], OT_object=OT_actual)
        self.RESUMEN: int = OT_actual.Convertir_a_tipo(Data['RESUMEN'], str, error_print=['TAREAS', 'RESUMEN'], OT_object=OT_actual)
```

**Propósito**: Representa una tarea dentro de una OT.  
**Pasos**:  
1. Convierte los valores del diccionario `Data` al tipo adecuado usando el método `Convertir_a_tipo`.
2. Asigna los valores a atributos internos.

---

