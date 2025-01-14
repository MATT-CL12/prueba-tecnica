
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

### Inicialización del Navegador

```python
class WebUploader_Class:
    def __init__(self, Data_config, verbose=True):
        brwoser_options = Options()
        self.driver = webdriver.Edge(brwoser_options)
        self.driver.maximize_window()
        self.driver.get(Data_config["url_path"])
```

Esta clase:
- Inicia un navegador Edge automatizado
- Configura las opciones del navegador
- Navega a la URL del sistema

### Gestión de Sesión

```python
def log_in(self, Data_config, IDs, verbose=True):
    self.username = Data_config["USER"]
    self.password = Data_config["PASS"]
    username_field = self.click_until_interactable(By.ID, IDs["username"])
    username_field.send_keys(self.username)
```

Maneja:
- El inicio de sesión en el sistema
- La validación de credenciales
- La navegación inicial

### Manipulación de OTs

```python
class Orden_Trabajo_Class():
    def __init__(self, WebUploader_object, DF_TAREAS_init, ...):
        self.WebUploader = WebUploader_object
        self.ID = self.Convertir_a_tipo(Data_OT['ID'], int)
```

Esta clase gestiona:
- La creación/modificación de OTs
- El llenado de campos
- La validación de datos
- El manejo de errores

### Funciones de Soporte

```python
def click_until_interactable(self, buscador, button_id, click=True):
    wait = WebDriverWait(self.driver, tiempo_wait)
    element = wait.until(EC.visibility_of_element_located((buscador, button_id)))
```

Incluye funciones para:
- Interactuar con elementos web
- Esperar elementos interactivos
- Manejar timeouts y errores

## Flujo Completo del Sistema

1. El sistema inicia y carga configuraciones
2. Establece conexión con el navegador
3. Inicia sesión en IBM Maximo
4. Por cada OT en el archivo:
   - Carga los datos relacionados
   - Crea o modifica la OT
   - Llena todos los campos requeridos
   - Valida y guarda los cambios
   - Registra el resultado
5. Mantiene un log detallado de todo el proceso
6. Maneja errores y reintentos cuando es necesario

## Características Clave

- Automatización robusta de procesos manuales
- Manejo de errores y recuperación
- Registro detallado de operaciones
- Sistema modular y extensible
- Validación de datos y operaciones
- Medición de rendimiento y tiempos
- Soporte para múltiples tipos de datos
- Capacidad de procesamiento masivo

El sistema está diseñado para ser:
- Confiable: Maneja errores y excepciones
- Eficiente: Procesa múltiples OTs en batch
- Flexible: Permite diferentes tipos de operaciones
- Mantenible: Código modular y bien estructurado
- Trazable: Registro detallado de operaciones