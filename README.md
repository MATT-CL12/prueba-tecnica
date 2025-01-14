# prueba-tecnica
 

# Proyecto: Automatización de Carga Masiva de Órdenes de Trabajo

## Descripción General
Este proyecto está diseñado para automatizar la carga masiva de órdenes de trabajo (OT) en un sistema de gestión empresarial utilizando Python y Selenium. El programa permite la configuración, validación y manipulación de datos provenientes de archivos Excel y JSON para ser procesados en IBM Maximo u otros sistemas similares.

## Características Principales
- **Carga Masiva**: Procesamiento de múltiples OT con detalles específicos como tareas, mano de obra, materiales, servicios y asignaciones.
- **Automatización Web**: Uso de Selenium para interactuar con el sistema de gestión empresarial mediante un navegador automatizado.
- **Registro de Errores**: Manejo de excepciones y registro detallado en archivos de depuración.
- **Configurabilidad**: Personalización de los parámetros de entrada mediante archivos JSON y Excel.

## Requisitos del Sistema
- **Lenguaje**: Python 3.13.1 (o superior).
- **Bibliotecas**:
  - pandas
  - selenium
- **WebDriver**: Microsoft Edge o compatible, con el driver correspondiente instalado.
- **Archivos de Entrada**:
  - `User_Config.json`: Configuración del usuario (credenciales y URL del sistema).
  - `ID_Config.json`: Identificadores de elementos HTML necesarios para la automatización.
  - `Campos_Diligenciar.json`: Especificación de campos que se deben llenar en cada OT.
  - Archivo Excel con datos de entrada: `Datos_entrada.xlsx`.

## Estructura del Proyecto
```plaintext
Proyecto/
|-- Input/
|   |-- Campos_Diligenciar
|   |-- User_Config.json
|   |-- ID_Config.json
|   |-- Datos_L33_forestal.xlsx
|   |-- Datos_entrada.xlsx
|   |-- Recurrentes/
|       |--Datos_guaduales_urbanos.xlsx
|
|-- WebUploader_Class.py
|-- Main.py
```

### Archivos Principales
- **`Main.py`**: Archivo principal que coordina la ejecución del programa.
- **`WebUploader_Class.py`**: Clase principal que encapsula la lógica de interacción con el navegador.

## Uso del Proyecto

### 1. Configuración Inicial
1. Asegúrte de instalar las dependencias con el siguiente comando:
   ```bash
   pip install selenium
   pip install pandas
   ```
2. Coloca los archivos de configuración JSON y el archivo Excel en la carpeta `Input/`.

### 2. Ejecución
Para ejecutar el programa, utiliza el siguiente comando:
```bash
python Main.py
```

### 3. Salida
- Los resultados se guardan en:
  - **`Output/Report/`**: Reportes detallados de la ejecución.
  - **`Output/Debug/`**: Archivos de depuración.

## Documentación de Clases y Métodos
### Clase: `WebUploader_Class`
Esta clase representa el driver usado por Selenium para automatizar la interacción con el navegador.

#### Métodos Principales
- **`log_in(Data_config, IDs, verbose=True)`**:
  Inicia sesión en el sistema.
  - **Parámetros**:
    - `Data_config`: Configuración del usuario.
    - `IDs`: Identificadores HTML.
  - **Salida**: Booleano que indica el éxito.

- **`log_menu(menu, verbose=False)`**:
  Navega al menú deseado dentro del sistema.

- **`Crear_OT(OT_actual, IDs, verbose_crearOT=False)`**:
  Crea una nueva orden de trabajo o modifica una existente.

- **`Ingresar_Datos(columnas, IDs, Object_OT, instancia_interna=None, verbose=False)`**:
  Ingresa datos en los campos especificados de la interfaz.

### Clase: `Orden_Trabajo_Class`
Gestiona la creación y manipulación de órdenes de trabajo.

#### Atributos Principales
- `ID`: Identificador de la OT.
- `OT`: Número de la OT.
- `Listado_Tareas`: Lista de tareas asociadas.

#### Métodos Principales
- **`guardarOT(IDs, verbose=False)`**: Guarda los cambios realizados en la OT.
- **`Eliminar_Filas(tipo, IDs, labor_bot="MODIFICAR")`**: Elimina filas específicas de la OT.

## Manejo de Errores
El sistema implementa excepciones detalladas para los errores encontrados durante la ejecución, como:
- Elementos no encontrados en la página web.
- Datos inválidos en los archivos de entrada.

## Licencia
Este proyecto está bajo Licencia de software libre.

## Contacto
Para preguntas o sugerencias, puedes contactar a:
- **Autor**: Alejandro López



