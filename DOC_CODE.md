# **Documentación del Sistema de Carga Masiva de Datos en Órdenes de Trabajo (OTs)**

Este sistema automatiza la carga de datos en una plataforma web utilizando Selenium, procesamiento multi-hilo y manipulación de archivos Excel.

## **Estructura del Sistema**

El sistema consta de tres archivos principales:

1. **`Main_Cargar_Servicios.py`**: Controlador principal del proceso de carga de datos.
2. **`multi_thread_runner.py`**: Ejecuta la carga en paralelo dividiendo los datos en fragmentos.
3. **`WebUploader_Class.py`**: Contiene la lógica de automatización con Selenium para interactuar con la plataforma web.

---

## **1. Archivo Principal: `Main_Cargar_Servicios.py`**

Este archivo es el punto de entrada del sistema. Su función es:
- Leer los archivos de configuración y datos de entrada.
- Inicializar la automatización con Selenium.
- Iterar sobre las filas del archivo de datos y ejecutar el proceso de carga en la plataforma web.

### **Estructura del Código**

### **1.1 Importación de Bibliotecas y Configuración**

```python
import pandas as pd
from WebUploader_Class import WebUploader_Class, Orden_Trabajo_Class
from time import time
import datetime
import json
import sys
import os
```

- Se importan las bibliotecas necesarias.
- `WebUploader_Class` y `Orden_Trabajo_Class` permiten interactuar con la web mediante Selenium.

### **1.2 Redirección de Salida a un Archivo de Log**

El sistema captura la salida estándar (`stdout`) para registrar la ejecución:

```python
class DualWriter:
    def __init__(self, original_stdout, archivo):
        self.original_stdout = original_stdout
        self.archivo = archivo
    
    def write(self, mensaje):
        self.original_stdout.write(mensaje)
        self.archivo.write(mensaje)
    
    def flush(self):
        self.original_stdout.flush()
        self.archivo.flush()
```

- Esta clase redirige la salida estándar (`print`) tanto a la consola como a un archivo de log.
- Se usa para registrar la ejecución del programa.

### **1.3 Carga de Configuraciones y Datos**

```python
with open("Input/User_Config.json", encoding='utf-8') as file:
    Data_config = json.load(file)
```

- Se leen las configuraciones del usuario desde un archivo JSON.
- Se carga el archivo Excel con las órdenes de trabajo y otros datos relacionados.

### **1.4 Inicialización del WebUploader y Autenticación**

```python
WebUploader = WebUploader_Class(Data_config)
WebUploader.log_in(Data_config, IDs["Login"])
```

- Se inicializa la clase `WebUploader_Class`.
- Se ejecuta el inicio de sesión en la plataforma web.

### **1.5 Iteración sobre las Órdenes de Trabajo (OTs)**

```python
for row, Datos in DF_OTs.iterrows():
    if Datos['LABOR_BOT']!='NADA':
        OT_actual = Orden_Trabajo_Class(WebUploader, DF_TAREAS_ACTUAL, DF_MANO_DE_OBRA_ACTUAL, DF_MATERIALES_ACTUAL, DF_SERVICIOS_ACTUAL, Datos)
        WebUploader.Crear_OT(OT_actual, IDs, True)
```

- Se recorre el archivo Excel.
- Se filtran las filas según el campo `LABOR_BOT`.
- Se crea una instancia de `Orden_Trabajo_Class` para manejar cada OT.

### **1.6 Manejo de Errores y Registro de Resultados**

```python
try:
    WebUploader.Crear_OT(OT_actual, IDs, True)
except Exception:
    Registro_error = pd.DataFrame({'ID': [Datos["ID"]], 'OT': [OT_actual.OT], 'Registro': [OT_actual.ERROR_OT]})
    df_registros = df_registros.append(Registro_error, ignore_index=True)
```

- Si ocurre un error durante la creación de la OT, se registra en un archivo de depuración (`Debug.xlsx`).

---

## **2. Procesamiento Multi-Hilo: `multi_thread_runner.py`**

Este archivo divide la carga de trabajo en varios hilos para mejorar el rendimiento.

### **2.1 División del Archivo en Fragmentos**

```python
def save_chunk_to_excel(data_dict, chunk_file):
    with pd.ExcelWriter(chunk_file) as writer:
        data_dict["OTs"].to_excel(writer, sheet_name="OTs", index=False)
```

- Se guarda cada fragmento de datos en un archivo Excel separado.

### **2.2 Creación de Hilos y Ejecución Paralela**

```python
with concurrent.futures.ThreadPoolExecutor(max_workers=n_threads) as executor:
    futures = []
    for thread_id, df_chunk in enumerate(chunks, start=1):
        future = executor.submit(worker, thread_id, df_chunk, full_excel)
        futures.append(future)
```

- Se usa `ThreadPoolExecutor` para ejecutar múltiples instancias de `Main_Cargar_Servicios.py` en paralelo.

---

## **3. Automatización con Selenium: `WebUploader_Class.py`**

Este archivo define la clase `WebUploader_Class`, encargada de interactuar con la plataforma web mediante Selenium.

### **3.1 Inicialización del Navegador**

```python
class WebUploader_Class:
    def __init__(self, Data_config, verbose=True):
        self.driver = webdriver.Edge(Options())
        self.driver.get(Data_config["url_path"])
```

- Se usa Selenium para iniciar un navegador Edge.
- Se navega a la URL de la plataforma web.

### **3.2 Inicio de Sesión**

```python
def log_in(self, Data_config, IDs, verbose=True):
    username_field = self.click_until_interactable(By.ID, IDs["username"])
    username_field.send_keys(self.username)
    password_field = self.click_until_interactable(By.ID, IDs["password"])
    password_field.send_keys(self.password)
```

- Se localizan los campos de usuario y contraseña.
- Se ingresan las credenciales automáticamente.

### **3.3 Creación y Manejo de Órdenes de Trabajo (OTs)**

```python
def Crear_OT(self, OT_actual, IDs, verbose_crearOT = False):
    if OT_actual.OT_existente == False:
        elemento_crearOT = self.click_until_interactable(By.ID, IDs["Buscadores"]["OTs"]["Crear_OT"])
```

- Si la OT no existe, se crea en la plataforma web.

### **3.4 Cambio de Ventanas y Flujo de OT**

```python
def Cambiar_ventana(self, ventana, IDs):
    self.WebUploader.click_until_interactable(By.ID, IDs[ventana]["VENTANA"])
    self.Ventana_actual = ventana
```

- Se cambia entre pestañas dentro de la plataforma web.

---


