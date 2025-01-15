# -*- coding: utf-8 -*-
"""
@author: Alejandro López
Principal

Código para cargar datos de manera masiva en un listado de OT, los datos cargados se encuentran definidos en Config.py

---------------------
Parametros de Entrada
---------------------     
Config : `archivo.py`
    Archivo .py con la configuración y parametros del código
    
1_OTs_Cargar : `xlsx`
     Archivo excel con las OT y los parametros a llenar   
"""

import pandas as pd
import threading
import queue
from concurrent.futures import ThreadPoolExecutor
from WebUploader_Class import WebUploader_Class, Orden_Trabajo_Class
from time import time
import datetime
import json
import pdb

# Bloqueo para sincronización
lock = threading.Lock()

# Inicialización del tiempo de ejecución global
start_program_time = time()

# Lectura de archivos de configuración
with open("Input\\User_Config.json", encoding='utf-8') as file:
    Data_config = json.load(file)
with open("Input\\ID_Config.json", encoding='utf-8') as file:
    IDs = json.load(file)
with open("Input\\Campos_Diligenciar.json", encoding='utf-8') as file:
    Campos_Diligenciar = json.load(file)

# Campos y flujo
campos_OT = Campos_Diligenciar["Campos_OT"]
Campos_Tarea = Campos_Diligenciar["Campos_Tarea"]
Campos_Mano_de_obra = Campos_Diligenciar["Campos_Mano_de_obra"]
Campos_Servicio = Campos_Diligenciar["Campos_Servicio"]
Campos_Asignacion = Campos_Diligenciar["Campos_Asignacion"]
listado_flujo = Campos_Diligenciar["Flujo_OT"]

# Lectura de Excel
name = 'Datos_entrada.xlsx'
DF_OTs = pd.read_excel('Input\\' + name, sheet_name="OTs", dtype=str)
DF_TAREAS = pd.read_excel('Input\\' + name, sheet_name="TAREAS", dtype=str)
DF_MANO_DE_OBRA = pd.read_excel('Input\\' + name, sheet_name="MO", dtype=str)
DF_MATERIALES = pd.read_excel('Input\\' + name, sheet_name="MATERIALES", dtype=str)
DF_SERVICIOS = pd.read_excel('Input\\' + name, sheet_name="SERVICIOS", dtype=str)

# DataFrame de resultados
df_registros = pd.DataFrame({'ID': [], 'OT': [], 'OT_PADRE': [], 'UBICACION': [], 'Registro': []})

# Crear la cola de tareas
task_queue = queue.Queue()
for _, row in DF_OTs.iterrows():
    if row['LABOR_BOT'] != 'NADA':
        task_queue.put(row)
        

# Función para procesar cada tarea
def process_task():
    global df_registros
    
    # Inicializar listas locales para consolidar logs y errores
    thread_logs = []
    thread_errors = []

    # Crear una instancia de WebUploader para este hilo
    WebUploader = WebUploader_Class(Data_config)

    # Iniciar sesión y navegar al menú OT
    WebUploader.log_in(Data_config, IDs["Login"])
    WebUploader.log_menu(IDs["Menus"]["Menu_OT"])
    
    while not task_queue.empty():
        try:
            task = task_queue.get_nowait()
            time_start = time()
            
            try:
                task_id = task["ID"]
                pdb.set_trace()

                # Filtrar dataframes relevantes
                DF_TAREAS_ACTUAL = DF_TAREAS[DF_TAREAS['ID'] == task_id].reset_index(drop=True)
                DF_MANO_DE_OBRA_ACTUAL = DF_MANO_DE_OBRA[DF_MANO_DE_OBRA['ID'] == task_id].reset_index(drop=True)
                DF_MATERIALES_ACTUAL = DF_MATERIALES[DF_MATERIALES['ID'] == task_id].reset_index(drop=True)
                DF_SERVICIOS_ACTUAL = DF_SERVICIOS[DF_SERVICIOS['ID'] == task_id].reset_index(drop=True)

                # Crear instancia de Orden_Trabajo_Class
                OT_actual = Orden_Trabajo_Class(WebUploader, DF_TAREAS_ACTUAL, DF_MANO_DE_OBRA_ACTUAL, DF_MATERIALES_ACTUAL, DF_SERVICIOS_ACTUAL, task)

                # Crear y procesar OT
                WebUploader.Crear_OT(OT_actual, IDs, True)
                
                # Consolidar logs de WebUploader
                thread_logs.append(WebUploader.logs)
                thread_errors.append(WebUploader.errors)
                
                OT_actual.Cambiar_ventana("Orden_trabajo", IDs)

                # Procesar datos en pestaña "Orden_trabajo"
                if task['LABOR_BOT'] in ['CREAR', 'MODIFICAR'] and int(task['LLENAR_OT']):
                    OT_actual.Ingresar_Datos(campos_OT, IDs["Orden_trabajo"], OT_actual, verbose=True)
                    OT_actual.guardarOT(IDs["Orden_trabajo"])

                # Procesar datos en pestaña "Planes"
                if any([int(task[col]) for col in ['LLENAR_TAREA', 'LLENAR_SERVICIO', 'LLENAR_MO', 'LLENAR_MATERIAL']]):
                    OT_actual.Cambiar_ventana("Planes", IDs)

                    # Procesar tareas
                    if int(task['LLENAR_TAREA']):
                        OT_actual.Eliminar_Filas("tareas", IDs["Planes"], task['LABOR_BOT'])
                        for row_tarea, _ in DF_TAREAS_ACTUAL.iterrows():
                            OT_actual.Fila_Nueva(IDs["Planes"]["tareas"])
                            OT_actual.Ingresar_Datos(Campos_Tarea, IDs["Planes"]["tareas"], OT_actual, OT_actual.Listado_Tareas[row_tarea], verbose=True)
                        OT_actual.guardarOT(IDs["Orden_trabajo"])

                    # Procesar mano de obra
                    if int(task['LLENAR_MO']):
                        OT_actual.Cambiar_ventana("mano_de_obra", IDs["Planes"])
                        OT_actual.Eliminar_Filas("mano_de_obra", IDs["Planes"], task['LABOR_BOT'])
                        for row_mano_de_obra, _ in DF_MANO_DE_OBRA_ACTUAL.iterrows():
                            OT_actual.Fila_Nueva(IDs["Planes"]["mano_de_obra"])
                            OT_actual.Ingresar_Datos(Campos_Mano_de_obra, IDs["Planes"]["mano_de_obra"], OT_actual, OT_actual.Listado_Mano_de_obra[row_mano_de_obra], verbose=True)
                        OT_actual.guardarOT(IDs["Orden_trabajo"])

                    # Procesar servicios
                    if int(task['LLENAR_SERVICIO']):
                        OT_actual.Cambiar_ventana("servicios", IDs["Planes"])
                        OT_actual.Eliminar_Filas("servicios", IDs["Planes"], task['LABOR_BOT'])
                        for row_servicio, _ in DF_SERVICIOS_ACTUAL.iterrows():
                            OT_actual.Fila_Nueva(IDs["Planes"]["servicios"])
                            OT_actual.Ingresar_Datos(Campos_Servicio, IDs["Planes"]["servicios"], OT_actual, OT_actual.Listado_Servicios[row_servicio], verbose=True)
                        OT_actual.guardarOT(IDs["Orden_trabajo"])
                        
                    # Procesar asignaciones
                    if int(task['LLENAR_MO']):
                        OT_actual.Cambiar_ventana("Asignaciones", IDs)
                        OT_actual.Eliminar_Filas("asignaciones", IDs["Asignaciones"], task['LABOR_BOT'])
                        for row_mano_de_obra, Datos_Asignaciones  in DF_MANO_DE_OBRA_ACTUAL.iterrows():
                            if not pd.isnull(Datos_Asignaciones["MANO_OBRA"]):
                                OT_actual.Fila_Nueva(IDs["Asignaciones"]["asignaciones"])
                                OT_actual.Ingresar_Datos(Campos_Asignacion, IDs["Asignaciones"]["asignaciones"], OT_actual, OT_actual.Listado_Asignaciones[row_mano_de_obra], verbose=True)
                        OT_actual.guardarOT(IDs["Orden_trabajo"])

                # Procesar flujo de la OT
                OT_actual.Cambiar_ventana("Orden_trabajo", IDs)
                OT_actual.Flujo_OT(IDs, task["ESTADO_DESEADO"], listado_flujo[task["ESTADO_DESEADO"]])

                # Registro de resultado
                registro = pd.DataFrame({'ID': [task_id], 'OT': [OT_actual.OT], 'OT_PADRE': [OT_actual.OT_PADRE],
                                         'UBICACION': [task['TRAMO/TRANSFORMADOR']], 'Registro': [OT_actual.ESTADO]})

                # Consolidar logs de OT_actual nuevamente
                thread_logs.append(OT_actual.logs)
                thread_errors.append(OT_actual.ERROR_OT)
                thread_logs.append(f"Tiempo ejecución OT {task_id}: {round(time() - time_start, 1)} segundos.")
                
                # Al final, imprimir el log acumulado de manera controlada
                with lock:
                    print("==== LOGS ====")
                    print("\n".join(thread_logs))
                    print("==== ERRORES ====")
                    print("\n".join(thread_errors))
        
                with lock:
                    df_registros = pd.concat([df_registros, registro], ignore_index=True)
                

            except Exception as e:
                # Manejar excepciones específicas y registrar el error
                error_registro = pd.DataFrame({'ID': [task["ID"]], 'OT': ["ERROR"], 'OT_PADRE': ["ERROR"],
                                               'UBICACION': [task['TRAMO/TRANSFORMADOR']], 'Registro': thread_errors})

                with lock:
                    df_registros = pd.concat([df_registros, error_registro], ignore_index=True)
                WebUploader.cerrar_navegador()
                
                # Consolidar logs de OT_actual nuevamente
                thread_logs.append(OT_actual.logs)
                thread_errors.append(OT_actual.ERROR_OT)
                thread_logs.append(f"Tiempo ejecución OT {task_id}: {round(time() - time_start, 1)} segundos.")
                
                # Al final, imprimir el log acumulado de manera controlada
                with lock:
                    print("==== LOGS ====")
                    print("\n".join(thread_logs))
                    print("==== ERRORES ====")
                    print("\n".join(thread_errors))
                
                WebUploader = WebUploader_Class(Data_config)
                WebUploader.log_in(Data_config, IDs["Login"])
                WebUploader.log_menu(IDs["Menus"]["Menu_OT"])
                        
        except queue.Empty:
            break
    
  
# Iniciar multithreading con ThreadPoolExecutor
num_threads = 2  # Ajustar según los recursos disponibles
with ThreadPoolExecutor(max_workers=num_threads) as executor:
    for _ in range(num_threads):
        executor.submit(process_task)

# Guardar resultados finales
output_file = "Output/Debug/Debug_" + datetime.datetime.now().strftime('D%Y_%m_%d_H%H_%M_%S') + ".xlsx"
df_registros.to_excel(output_file, index=False)

# Calcular tiempo total del programa
end_program_time = time()
print(f"Tiempo total del programa: {round(end_program_time - start_program_time, 1)} segundos")
