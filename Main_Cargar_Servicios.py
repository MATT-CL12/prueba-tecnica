import pandas as pd
from concurrent.futures import ThreadPoolExecutor, as_completed
from WebUploader_Class import WebUploader_Class, Orden_Trabajo_Class
from time import time
import json
import sys
import datetime
from threading import current_thread

# Clase DualWriter para redireccionar la salida estándar
time_program = {"start_program": time(), "end_program": 0}

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

# Nombre del archivo donde se almacenarán los prints
nombre_archivo = "Output/Report/Report_" + datetime.datetime.now().strftime('D%Y_%m_%d_H%H_%M_%S') + ".txt"

# Función para procesar una fila del Excel
def procesar_fila(fila, Data_config, IDs, campos, DF_TAREAS, DF_MANO_DE_OBRA, DF_MATERIALES, DF_SERVICIOS):
    thread_name = current_thread().name
    try:
        start_time = time()
        print(f"[Hilo {thread_name}] Inicio procesamiento fila ID: {fila['ID']}")

        # Crear instancia independiente de WebUploader_Class
        web_uploader = WebUploader_Class(Data_config)
        web_uploader.log_in(Data_config, IDs["Login"])

        # Crear instancias internas según los datos de la fila
        DF_TAREAS_ACTUAL = DF_TAREAS.loc[DF_TAREAS['ID'] == fila['ID']].reset_index(drop=True)
        DF_MANO_DE_OBRA_ACTUAL = DF_MANO_DE_OBRA.loc[DF_MANO_DE_OBRA['ID'] == fila['ID']].reset_index(drop=True)
        DF_MATERIALES_ACTUAL = DF_MATERIALES.loc[DF_MATERIALES['ID'] == fila['ID']].reset_index(drop=True)
        DF_SERVICIOS_ACTUAL = DF_SERVICIOS.loc[DF_SERVICIOS['ID'] == fila['ID']].reset_index(drop=True)

        # Crear instancia de Orden_Trabajo
        OT_actual = Orden_Trabajo_Class(web_uploader, DF_TAREAS_ACTUAL, DF_MANO_DE_OBRA_ACTUAL, DF_MATERIALES_ACTUAL, DF_SERVICIOS_ACTUAL, fila)

        # Ejecutar el proceso de la OT
        web_uploader.Crear_OT(OT_actual, IDs, True)
        OT_actual.Cambiar_ventana("Orden_trabajo", IDs)

        # Diligenciar datos de la OT
        if fila['LABOR_BOT'] == 'CREAR' or (fila['LABOR_BOT'] == 'MODIFICAR' and int(fila['LLENAR_OT']) == True):
            OT_actual.Ingresar_Datos(campos["Campos_OT"], IDs["Orden_trabajo"], OT_actual, verbose=True)
            OT_actual.guardarOT(IDs["Orden_trabajo"])

        # Registrar éxito
        registro = pd.DataFrame({
            'ID': [fila['ID']],
            'OT': [OT_actual.OT],
            'Registro': ['Éxito']
        })

        web_uploader.cerrar_navegador()

        end_time = time()
        print(f"[Hilo {thread_name}] Fin procesamiento fila ID: {fila['ID']} - Duración: {round(end_time - start_time, 2)} segundos")

        return registro

    except Exception as e:
        web_uploader.cerrar_navegador()
        print(f"[Hilo {thread_name}] Error en procesamiento fila ID: {fila['ID']} - {str(e)}")

        registro_error = pd.DataFrame({
            'ID': [fila['ID']],
            'OT': [fila.get('OT', 'N/A')],
            'Registro': [str(e)]
        })
        return registro_error

# Abre el archivo en modo de anexado ('a')
with open(nombre_archivo, 'a') as archivo:
    stdout_original = sys.stdout
    dual_writer = DualWriter(stdout_original, archivo)
    sys.stdout = dual_writer

    try:
        # Cargar configuraciones y datos
        with open("Input/User_Config.json", encoding='utf-8') as file:
            Data_config = json.load(file)

        with open("Input/ID_Config.json", encoding='utf-8') as file:
            IDs = json.load(file)

        with open("Input/Campos_Diligenciar.json", encoding='utf-8') as file:
            Campos_Diligenciar = json.load(file)

        campos = Campos_Diligenciar

        # Cargar datos del Excel
        DF_OTs = pd.read_excel('Input/Datos_entrada.xlsx', sheet_name="OTs", dtype=str)
        DF_TAREAS = pd.read_excel('Input/Datos_entrada.xlsx', sheet_name="TAREAS", dtype=str)
        DF_MANO_DE_OBRA = pd.read_excel('Input/Datos_entrada.xlsx', sheet_name="MO", dtype=str)
        DF_MATERIALES = pd.read_excel('Input/Datos_entrada.xlsx', sheet_name="MATERIALES", dtype=str)
        DF_SERVICIOS = pd.read_excel('Input/Datos_entrada.xlsx', sheet_name="SERVICIOS", dtype=str)

        # Procesar en paralelo con un límite de 3 procesos
        resultados = []
        with ThreadPoolExecutor(max_workers=3) as executor:
            future_to_id = {
                executor.submit(procesar_fila, fila, Data_config, IDs, campos, DF_TAREAS, DF_MANO_DE_OBRA, DF_MATERIALES, DF_SERVICIOS): fila['ID']
                for fila in DF_OTs.to_dict('records')
            }

            for future in as_completed(future_to_id):
                fila_id = future_to_id[future]
                try:
                    result = future.result()
                    resultados.append(result)
                except Exception as exc:
                    print(f"[Main] Error en hilo para fila ID {fila_id}: {exc}")

        # Combinar resultados
        registros_finales = pd.concat(resultados, ignore_index=True)

        # Guardar resultados
        name_output = "Output/Debug/Debug_" + datetime.datetime.now().strftime('D%Y_%m_%d_H%H_%M_%S') + ".xlsx"
        registros_finales.to_excel(name_output, index=False)

        time_program["end_program"] = time()
        print("--------------------------------------------------------")
        print("Tiempo ejecución total: ", round(time_program["end_program"] - time_program["start_program"], 1))

    finally:
        sys.stdout = stdout_original
