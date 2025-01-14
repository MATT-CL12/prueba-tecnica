## Documentación del Código

A continuación se detalla el funcionamiento de cada una de las funciones presentes en los archivos `Main.py` y `WebUploader_Class.py`, explicando su propósito, entrada, salida y lógica interna.

---

### Archivo: `Main.py`

Este archivo actúa como el punto de entrada del programa y coordina todas las operaciones principales.

#### **`main()`**

- **Propósito**: Es el controlador principal que organiza la ejecución del programa.
- **Lógica Interna**:
  1. Carga los archivos JSON de configuración:
     - `User_Config.json`: Contiene las credenciales y configuraciones generales del usuario.
     - `ID_Config.json`: Contiene los identificadores de los elementos HTML necesarios.
     - `Campos_Diligenciar.json`: Define los campos que se deben llenar en cada OT.
  2. Lee el archivo Excel (`Datos_entrada.xlsx`) para obtener las OT a procesar.
  3. Crea una instancia de la clase `WebUploader_Class` para manejar las interacciones con el navegador.
  4. Itera sobre cada OT en el archivo Excel y llama a las funciones necesarias para procesarla.
  5. Genera reportes de ejecución y errores.

---

### Archivo: `WebUploader_Class.py`

Esta clase encapsula todas las interacciones con el navegador y las operaciones relacionadas con el procesamiento de las OT.

#### **Clase: `WebUploader_Class`**

##### **`__init__(self, Data_config, verbose=True)`**

- **Propósito**: Inicializa el navegador y configura las opciones necesarias para Selenium.
- **Entrada**:
  - `Data_config`: Diccionario con las configuraciones del usuario, como URL y credenciales.
  - `verbose`: Indica si se deben imprimir mensajes de depuración.
- **Lógica Interna**:
  1. Configura las opciones del navegador.
  2. Inicia una instancia de WebDriver para Microsoft Edge.
  3. Navega a la URL proporcionada en `Data_config`.

##### **`log_in(self, Data_config, IDs, verbose=True)`**

- **Propósito**: Inicia sesión en el sistema utilizando las credenciales proporcionadas.
- **Entrada**:
  - `Data_config`: Diccionario con el nombre de usuario y contraseña.
  - `IDs`: Diccionario con los identificadores de los campos HTML para el inicio de sesión.
- **Salida**: Devuelve `True` si el inicio de sesión fue exitoso; de lo contrario, `False`.
- **Lógica Interna**:
  1. Encuentra los elementos HTML para el nombre de usuario y contraseña.
  2. Ingresa las credenciales y selecciona el idioma si es necesario.
  3. Valida que el inicio de sesión fue exitoso mediante la presencia de un elemento específico.

##### **`log_menu(self, menu, verbose=False)`**

- **Propósito**: Navega al menú deseado dentro del sistema.
- **Entrada**:
  - `menu`: Lista de pasos del menú a seleccionar.
  - `verbose`: Indica si se deben imprimir mensajes de depuración.
- **Salida**: Devuelve `True` si la navegación fue exitosa; de lo contrario, `False`.
- **Lógica Interna**:
  1. Itera sobre los elementos del menú especificados en la lista.
  2. Hace clic en cada elemento utilizando su identificador HTML.

##### **`Crear_OT(self, OT_actual, IDs, verbose_crearOT=False)`**

- **Propósito**: Crea una nueva OT o accede a una existente en el sistema.
- **Entrada**:
  - `OT_actual`: Objeto que contiene la información de la OT a procesar.
  - `IDs`: Diccionario con los identificadores HTML necesarios.
  - `verbose_crearOT`: Indica si se deben imprimir mensajes de depuración.
- **Lógica Interna**:
  1. Valida si la OT ya existe en el sistema.
  2. Si no existe, utiliza los identificadores para crearla.
  3. Si ya existe, navega a sus detalles.

##### **`Ingresar_Datos(self, columnas, IDs, Object_OT, instancia_interna=None, verbose=False)`**

- **Propósito**: Llena los campos de la interfaz de usuario con los datos proporcionados.
- **Entrada**:
  - `columnas`: Lista de campos a llenar.
  - `IDs`: Diccionario con los identificadores HTML.
  - `Object_OT`: Objeto que representa la OT.
  - `instancia_interna`: (Opcional) Subcomponente de la OT que se está manipulando.
  - `verbose`: Indica si se deben imprimir mensajes de depuración.
- **Lógica Interna**:
  1. Itera sobre los campos especificados en `columnas`.
  2. Verifica el tipo de dato del campo (texto, booleano, lista, etc.).
  3. Realiza las acciones correspondientes para ingresar el dato en el campo HTML.
  4. Maneja errores como celdas vacías o datos no válidos.

##### **`cerrar_navegador(self)`**

- **Propósito**: Cierra el navegador y finaliza la sesión de Selenium.
- **Lógica Interna**:
  1. Llama al método `quit()` del WebDriver para cerrar todas las ventanas del navegador.

---

#### Clase: `Orden_Trabajo_Class`

##### **`__init__(self, WebUploader_object, DF_TAREAS_init, DF_MANO_DE_OBRA_init, DF_MATERIALES_init, DF_SERVICIOS_init, Data_OT)`**

- **Propósito**: Inicializa una OT con los datos proporcionados.
- **Entrada**:
  - `WebUploader_object`: Instancia de `WebUploader_Class` para manejar el navegador.
  - `DF_TAREAS_init`: DataFrame con información inicial de tareas.
  - `DF_MANO_DE_OBRA_init`: DataFrame con información inicial de mano de obra.
  - `DF_MATERIALES_init`: DataFrame con información inicial de materiales.
  - `DF_SERVICIOS_init`: DataFrame con información inicial de servicios.
  - `Data_OT`: Diccionario con los datos de la OT.

##### **`guardarOT(self, IDs, verbose=False)`**

- **Propósito**: Guarda los cambios realizados en la OT en el sistema.
- **Entrada**:
  - `IDs`: Diccionario con los identificadores HTML.
  - `verbose`: Indica si se deben imprimir mensajes de depuración.
- **Lógica Interna**:
  1. Encuentra el botón de guardar en la interfaz.
  2. Realiza la acción de guardado y verifica errores.

##### **`Eliminar_Filas(self, tipo, IDs, labor_bot="MODIFICAR")`**

- **Propósito**: Elimina todas las filas de una sección específica de la OT.
- **Entrada**:
  - `tipo`: Sección a limpiar (e.g., "tareas", "materiales").
  - `IDs`: Identificadores HTML necesarios.
  - `labor_bot`: Tipo de acción (crear o modificar).
- **Lógica Interna**:
  1. Encuentra y elimina todas las filas de la sección especificada.
  2. Guarda los cambios tras completar la eliminación.

---

### Ejemplo de Ejecución
```python
from WebUploader_Class import WebUploader_Class
from Orden_Trabajo_Class import Orden_Trabajo_Class
import json

# Configuración inicial
with open("Input/User_Config.json", encoding='utf-8') as file:
    Data_config = json.load(file)

with open("Input/ID_Config.json", encoding='utf-8') as file:
    IDs = json.load(file)

# Inicializar WebUploader
uploader = WebUploader_Class(Data_config)

# Iniciar sesión
if uploader.log_in(Data_config, IDs):
    print("Inicio de sesión exitoso")

# Procesar una OT de ejemplo
Data_OT = {
    "ID": 12345,
    "OT": 67890,
    "DESCRIPCION": "Reparación de equipo",
}

ot_actual = Orden_Trabajo_Class(uploader, DF_TAREAS_init=None, DF_MANO_DE_OBRA_init=None,
                                DF_MATERIALES_init=None, DF_SERVICIOS_init=None, Data_OT=Data_OT)

uploader.Crear_OT(ot_actual, IDs)
uploader.cerrar_navegador()
```

