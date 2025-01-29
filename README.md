# Sistema de Carga Masiva de Datos en Órdenes de Trabajo (OTs)

## Descripción

Este sistema automatiza la carga masiva de Órdenes de Trabajo (OTs) en una plataforma web. Utiliza Selenium para interactuar con la interfaz, procesamiento multi-hilo para mejorar la eficiencia y manipula archivos Excel como fuente de datos. Se compone de tres módulos principales:

1. **`Main_Cargar_Servicios.py`**: Controla la ejecución del proceso de carga de OTs.
2. **`multi_thread_runner.py`**: Divide la carga de trabajo en varios hilos para optimizar el rendimiento.
3. **`WebUploader_Class.py`**: Implementa la automatización con Selenium para interactuar con la web.

---

## Características

- **Carga masiva automatizada** de OTs en una plataforma web.
- **Manejo de archivos Excel** para estructurar y procesar los datos.
- **Automatización con Selenium**, interactuando con formularios web.
- **Paralelización** del proceso para mejorar la eficiencia.
- **Generación de logs y reportes** para depuración y seguimiento.

---

## Estructura del Proyecto

```
.
├── Input/                  # Archivos de entrada (configuración, datos Excel)
│   ├── User_Config.json    # Configuración del usuario y credenciales
│   ├── ID_Config.json      # IDs de los elementos en la página web
│   ├── Datos_entrada.xlsx  # Archivo con datos de las OTs
│   ├── Campos_Diligenciar.json
│
├── Output/                 # Directorio de salida con reportes y logs
│   ├── Report/             # Registros de ejecución
│   ├── Debug/              # Registros de errores
│
├── Main_Cargar_Servicios.py   # Script principal que coordina el proceso
├── multi_thread_runner.py     # Controla la ejecución multi-hilo
├── WebUploader_Class.py       # Automatiza la interacción con la web
├── README.md                  # Documentación del proyecto
```

---

## Instalación

### 1. Clonar el repositorio
```bash
git clone https://github.com/Zarcasmo/Web_Scraping_MX
cd main
```

### 2. Crear y activar un entorno virtual
```bash
python -m venv venv
source venv/bin/activate  # En Windows: venv\Scripts\activate
```

### 3. Instalar dependencias
```bash
pip install -r requirements.txt
```

---

## Uso

### Ejecución con un solo hilo
```bash
python Main_Cargar_Servicios.py
```

Este comando:
1. Carga los datos desde `Datos_entrada.xlsx`.
2. Inicia sesión en la plataforma web usando `WebUploader_Class.py`.
3. Itera sobre cada fila del archivo y automatiza la carga en la web.
4. Genera reportes y logs en `Output/`.

### Ejecución con procesamiento multi-hilo
```bash
python multi_thread_runner.py
```

Este script:
1. Divide el archivo `Datos_entrada.xlsx` en fragmentos.
2. Crea hilos para procesar cada fragmento en paralelo.
3. Ejecuta `Main_Cargar_Servicios.py` en cada hilo con datos separados.
4. Borra los archivos temporales y genera logs de ejecución.

---

## Explicación de los Códigos

### `Main_Cargar_Servicios.py` (Script Principal)

```python
with open("Input/User_Config.json", encoding='utf-8') as file:
    Data_config = json.load(file)
```
Carga las configuraciones del usuario, incluyendo credenciales y configuraciones de la plataforma.

```python
WebUploader = WebUploader_Class(Data_config)
WebUploader.log_in(Data_config, IDs["Login"])
```
Inicia sesión en la plataforma web usando Selenium.

```python
for row, Datos in DF_OTs.iterrows():
    if Datos['LABOR_BOT']!='NADA':
        OT_actual = Orden_Trabajo_Class(WebUploader, DF_TAREAS_ACTUAL, DF_MANO_DE_OBRA_ACTUAL, DF_MATERIALES_ACTUAL, DF_SERVICIOS_ACTUAL, Datos)
        WebUploader.Crear_OT(OT_actual, IDs, True)
```
Recorre el archivo Excel, filtra las filas según `LABOR_BOT` y automatiza la carga de datos.

### `multi_thread_runner.py` (Procesamiento Multi-Hilo)

```python
with concurrent.futures.ThreadPoolExecutor(max_workers=n_threads) as executor:
    futures = []
    for thread_id, df_chunk in enumerate(chunks, start=1):
        future = executor.submit(worker, thread_id, df_chunk, full_excel)
        futures.append(future)
```
Divide la carga en varios hilos y ejecuta `Main_Cargar_Servicios.py` en paralelo.

### `WebUploader_Class.py` (Automatización con Selenium)

```python
class WebUploader_Class:
    def __init__(self, Data_config, verbose=True):
        self.driver = webdriver.Edge(Options())
        self.driver.get(Data_config["url_path"])
```
Inicia el navegador Edge con Selenium y accede a la URL configurada.

```python
def log_in(self, Data_config, IDs, verbose=True):
    username_field = self.click_until_interactable(By.ID, IDs["username"])
    username_field.send_keys(self.username)
    password_field = self.click_until_interactable(By.ID, IDs["password"])
    password_field.send_keys(self.password)
```
Automatiza el inicio de sesión en la plataforma.

```python
def Crear_OT(self, OT_actual, IDs, verbose_crearOT = False):
    if OT_actual.OT_existente == False:
        elemento_crearOT = self.click_until_interactable(By.ID, IDs["Buscadores"]["OTs"]["Crear_OT"])
```
Si la OT no existe, la crea en la plataforma web.

---

## Registro y Depuración
Los logs y resultados se almacenan en la carpeta `Output/`. En caso de errores, se generan registros detallados en `Output/Debug/Debug.xlsx`.

---

## Contribuciones

1. Realiza un fork del proyecto.
2. Crea una nueva rama (`git checkout -b feature/nueva-funcionalidad`).
3. Realiza tus cambios y haz commit (`git commit -m 'Añadir nueva funcionalidad'`).
4. Haz push a la rama (`git push origin feature/nueva-funcionalidad`).
5. Abre un Pull Request.

---

## Licencia

Este proyecto está bajo la Licencia Software libre.
