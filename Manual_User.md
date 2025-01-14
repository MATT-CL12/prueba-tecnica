# Manual de Usuario para el Sistema de Ingreso de Datos

## Introducción
Este sistema permite automatizar el ingreso de datos en una plataforma web a partir de un archivo Excel. El código está diseñado para procesar la información contenida en un archivo de entrada, estructurado con varias hojas que representan diferentes categorías de datos (como OTs, Tareas, Materiales, etc.). Este manual describe cómo preparar los datos y ejecutar el sistema correctamente.

---

## Requisitos Previos
1. **Entorno de Ejecución**:
   - Python 3.x instalado.
   - Bibliotecas necesarias: `pandas`, `selenium`.

2. **Archivos Necesarios**:
   - Archivo principal de código Python (`main.py`).
   - Archivo de configuración JSON para IDs y campos (`User_Config.json`, `ID_Config.json`, `Campos_Diligenciar.json`).
   - Archivo Excel de entrada: `Datos_entrada.xlsx`.

---

## Estructura del Archivo Excel
El archivo `Datos_entrada.xlsx` debe contener los siguientes datos:

|                   |                    |                    |                   |
|-------------------|--------------------|--------------------|-------------------|
| LABOR_BOT         | OT                 | OT_PADRE           | LLENAR_OT         |
| LLENAR_TAREA      | LLENAR_SERVICIO    | LLENAR_MO          | LLENAR_MATERIAL   |
| PLAN_TRABAJO      | NOMBRE_PLANIFICADOR| NOMBRE_INTERVENTOR | DESCRIPCION       |
| INICIO_PREVISTO   | FIN_PREVISTO       | INICIO_PROGRAMADO  | FIN_PROGRAMADO    |
| CIRCUITO          | TRAMO/TRANSFORMADOR| NOMBRE_CLASIFICACION| ESTADO_DESEADO   |
| UBICACION         | CLASIFICACION      | TIPO_TRABAJO       | TIPO_PROYECTO     |
| PRIORIDAD         | RESIDUOS           | HEREDA             | GROT              |
| FSE               | ACCION             | DURACION           | PLANIFICADOR      |
| INTERVENTOR       | RESPONSABLE        | UNIDAD_NEGOCIO     | ACTIVIDAD_COSTEO  |
| CONTRATO          |                    |                    |                   |



---

## Uso del Sistema

### Paso 1: Preparar el Archivo Excel
1. **Actualizar Datos**:
   - Abra el archivo `Datos_entrada.xlsx` con un editor compatible (como Microsoft Excel o LibreOffice Calc).
   - Modifique o actualice las hojas según sea necesario.

2. **Validar Estructura**:
   - Asegúrese de que todas las columnas obligatorias estén presentes.
   - Verifique que las celdas requeridas no estén vacías.

3. **Guardar Cambios**:
   - Guarde el archivo como `Datos_entrada.xlsx` en la carpeta `Input`.

### Paso 2: diligenciar datos del usuario
1. Abra el archivo `User_Config` 
2. digita datos de usuario ubicados en:
```bash
  "USER": "usuario",
  "PASS": "contraseña",
   ```
3. Guarda el documento y cierralo.

### Paso 3: Ejecutar el Sistema
1. Abra una terminal o consola.
2. Navegue al directorio donde se encuentra el archivo `main.py`.
3. Ejecute el siguiente comando:
   ```bash
   python main.py
   ```

### Paso 4: Revisar Resultados
1. **Archivos de Salida**:
   - El sistema generará un archivo Excel con los registros procesados y errores en la carpeta `Output/Debug`.

2. **Registro de Logs**:
   - Los logs detallados se almacenan en `Output/Report`.

---

## Reemplazar el Archivo Excel
Si desea procesar nuevos datos:
1. Prepare un nuevo archivo Excel con la misma estructura que `Datos_entrada.xlsx`.
2. Reemplace el archivo anterior ubicado en `Input` con el nuevo archivo, usando el mismo nombre (`Datos_entrada.xlsx`).
3. Repita los pasos descritos en "Uso del Sistema".

---

## Resolución de Problemas
- **Error: "Archivo no encontrado"**:
  - Verifique que el archivo `Datos_entrada.xlsx` esté en la carpeta `Input`.

- **Campos Vacíos**:
  - Asegúrese de que todas las celdas requeridas estén completas en el Excel.

- **Errores en el Navegador**:
  - Verifique que el driver de Selenium esté correctamente configurado.

---

## Contribuir y Actualizar el Código
1. **Repositorio en GitHub**:
   - Suba el archivo `main.py` y este manual al repositorio.
   - Incluya un archivo `README.md` para documentación adicional.

2. **Estructura del Repositorio**:
   ```
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

