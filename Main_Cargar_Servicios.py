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
from WebUploader_Class import WebUploader_Class, Orden_Trabajo_Class
from time import time
import datetime
import json
import sys
import os
#Funcion para determinar el tiempo de ejcución
time_program = {"start_program": time(),"end_program":0,
                "start_OT":0, "end_OT":0, "list_time_OT":[]}  

#Clase para poder escribir los print en un archivo .txt
#La clase permite escribir tanto en stdout como en un archivo
class DualWriter:
    def __init__(self, original_stdout, archivo):
        self.original_stdout = original_stdout  # Referencia al stdout original
        self.archivo = archivo  # Archivo donde se almacenarán los prints
    
    def write(self, mensaje):
        self.original_stdout.write(mensaje)  # Escribe el mensaje en la consola
        self.archivo.write(mensaje)  # Escribe el mensaje en el archivo
    
    def flush(self):
        self.original_stdout.flush()  # Asegura que se vacíen los buffers de la consola
        self.archivo.flush()  # Asegura que se vacíen los buffers del archivo

# Nombre del archivo donde se almacenarán los prints
nombre_archivo = "Output/Report/Report_" + datetime.datetime.now().strftime('D%Y_%m_%d_H%H_%M_%S') +".txt"

# Abre el archivo en modo de anexado ('a')
with open(nombre_archivo, 'a') as archivo:
    # Guarda una referencia al stdout original
    stdout_original = sys.stdout
    
    # Crea una instancia de DualWriter
    dual_writer = DualWriter(stdout_original, archivo)
    
    # Redirige stdout a la instancia de DualWriter
    sys.stdout = dual_writer

    try:
    
        #%%Lectura de archivos de parametrizacion
        with open("Input\\User_Config.json", encoding='utf-8') as file:
            Data_config = json.load(file)
            
        #Lectrua del archivo JSON con los IDs que requiere Selenium
        with open("Input\\ID_Config.json", encoding='utf-8') as file:
            IDs = json.load(file)
            
        #Lectura del archivo JSON con los IDs que requiere Selenium
        with open("Input\\Campos_Diligenciar.json", encoding='utf-8') as file:
            Campos_Diligenciar = json.load(file)
        
        #%% Campos que se van a diligenciar
        #Campos a diligenciar en la OT
        campos_OT = Campos_Diligenciar["Campos_OT"]
        Campos_Tarea = Campos_Diligenciar["Campos_Tarea"]
        Campos_Mano_de_obra = Campos_Diligenciar["Campos_Mano_de_obra"]
        #Campos_Material = Campos_Diligenciar["Campos_Material"]
        Campos_Servicio = Campos_Diligenciar["Campos_Servicio"]
        Campos_Asignacion = Campos_Diligenciar["Campos_Asignacion"]
        Campos_Cancelar = Campos_Diligenciar["Campos_Cancelar"]
        #Orden del flujo de las ordenes de trabajo
        listado_flujo = Campos_Diligenciar["Flujo_OT"]
        
        #%% Lectura de Excel
       
        if len(sys.argv) > 1:
            excel_name = sys.argv[1]  # Nombre del Excel que viene por argumento
        else:
            excel_name = "Datos_entrada.xlsx"  # Valor por defecto si no pasas nada

        # IMPORTANTE: Aquí ajustas la ruta según como le pases el arg:
        # Si le pasas solo el nombre (ej. "Datos_entrada_chunk_1.xlsx"),
        # lo concatenas con la carpeta Input\
        excel_path = os.path.join("Input", excel_name)

        # Luego lees:
        DF_OTs = pd.read_excel(excel_path, sheet_name="OTs", dtype=str)
        DF_TAREAS = pd.read_excel(excel_path, sheet_name="TAREAS", dtype=str)
        DF_MANO_DE_OBRA = pd.read_excel(excel_path, sheet_name="MO", dtype=str)
        DF_MATERIALES = pd.read_excel(excel_path, sheet_name="MATERIALES", dtype=str)
        DF_SERVICIOS = pd.read_excel(excel_path, sheet_name="SERVICIOS", dtype=str)
        
        #%% Creación de WebUploader
        WebUploader = WebUploader_Class(Data_config)
        
        #%%% Método log_in para iniciar sesión
        WebUploader.log_in(Data_config, IDs["Login"])
        
        #%%% Ingreso a menu OT en MX
        # Se ejecuta el metodo log_menu para ingresar al menu de OT de MX
        WebUploader.log_menu(IDs["Menus"]["Menu_OT"])
        
        #%% Proceso iterativo por OT
        
        # Lista para almacenar los objetos de la clase Orden_Trabajo
        orden_list_objects = []
        #Registro de las OTs 
        df_registros = pd.DataFrame({'ID': [], 'OT':[], 'OT_PADRE':[],'UBICACION':[],'Registro': []})
        
        for row, Datos in DF_OTs.iterrows():
            if Datos['LABOR_BOT']!='NADA':
                try:
                    time_program["start_OT"] = time()
                    #%%%%Creacion de instancias
                    #Parametros para las instancias internas
                    DF_TAREAS_ACTUAL = DF_TAREAS.loc[DF_TAREAS['ID']==Datos['ID']].reset_index(drop=True)
                    DF_MANO_DE_OBRA_ACTUAL = DF_MANO_DE_OBRA.loc[DF_MANO_DE_OBRA['ID']==Datos['ID']].reset_index(drop=True)
                    DF_MATERIALES_ACTUAL = DF_MATERIALES.loc[DF_MATERIALES['ID']==Datos['ID']].reset_index(drop=True)
                    DF_SERVICIOS_ACTUAL = DF_SERVICIOS.loc[DF_SERVICIOS['ID']==Datos['ID']].reset_index(drop=True)
                    
                    #Instancia de tipo OT y sus instancias internas
                    OT_actual = Orden_Trabajo_Class(WebUploader, DF_TAREAS_ACTUAL, DF_MANO_DE_OBRA_ACTUAL, DF_MATERIALES_ACTUAL, DF_SERVICIOS_ACTUAL,  Datos)
                            
                    #%%% Diligenciamiento OT en MX
                    WebUploader.Crear_OT(OT_actual, IDs, True)     # Creación de la OT
                    
                    #%%%% Pestaña Orden de trabajo
                    OT_actual.Cambiar_ventana("Orden_trabajo", IDs)
                    #Se ingresan los datos en cada campo de la OT
                    if (Datos['LABOR_BOT']=='CREAR') or ((Datos['LABOR_BOT']=='MODIFICAR') and (int(Datos['LLENAR_OT'])==True)):
                        OT_actual.Ingresar_Datos(campos_OT, IDs["Orden_trabajo"], OT_actual, verbose=True)
                        OT_actual.guardarOT(IDs["Orden_trabajo"])
                        
                    #%%%% Pestaña planes
                    if (int(Datos['LLENAR_TAREA'])==True) or (int(Datos['LLENAR_SERVICIO'])==True) or (int(Datos['LLENAR_MO'])==True) or (int(Datos['LLENAR_MATERIAL'])==True):
                        OT_actual.Cambiar_ventana("Planes", IDs)
                        #%%%%% Tareas
                        if (int(Datos['LLENAR_TAREA'])==True):
                            #Eliminar todas las tareas de la OT
                            OT_actual.Eliminar_Filas("tareas", IDs["Planes"], Datos['LABOR_BOT'])
                            #Agregar tareas y sus parametros
                            for row_tarea,_ in DF_TAREAS_ACTUAL.iterrows():
                                #Se agrega la fila nueva para cada tarea
                                OT_actual.Fila_Nueva(IDs["Planes"]["tareas"])
                                #Se ingresan los datos en cada campo de la pestaña tareas
                                OT_actual.Ingresar_Datos(Campos_Tarea, IDs["Planes"]["tareas"], OT_actual, OT_actual.Listado_Tareas[row_tarea], verbose=True)
                            #Guardar OT
                            OT_actual.guardarOT(IDs["Orden_trabajo"])
                        
                        #%%%%% Mano de obra
                        if (int(Datos['LLENAR_MO'])==True):
                            OT_actual.Cambiar_ventana("mano_de_obra", IDs["Planes"])
                            #Eliminar todas las MO de la OT
                            OT_actual.Eliminar_Filas("mano_de_obra", IDs["Planes"], Datos['LABOR_BOT'])
                            #Agregar MO y sus parametros
                            for row_mano_de_obra,_ in DF_MANO_DE_OBRA_ACTUAL.iterrows():
                                 #Se agrega la fila nueva para cada servicio
                                 OT_actual.Fila_Nueva(IDs["Planes"]["mano_de_obra"])
                                 #Se ingresan los datos en cada campo de la pestaña servicio
                                 OT_actual.Ingresar_Datos(Campos_Mano_de_obra, IDs["Planes"]["mano_de_obra"], OT_actual, OT_actual.Listado_Mano_de_obra[row_mano_de_obra], verbose=True)
                            #Guardar OT
                            OT_actual.guardarOT(IDs["Orden_trabajo"])
                        
                        #%%%%% Servicios
                        if (int(Datos['LLENAR_SERVICIO'])==True):
                            OT_actual.Cambiar_ventana("servicios", IDs["Planes"])
                            #Eliminar todas las servicios de la OT
                            OT_actual.Eliminar_Filas("servicios", IDs["Planes"], Datos['LABOR_BOT'])
                            #Agregar servicios y sus parametros
                            for row_servicio,_ in DF_SERVICIOS_ACTUAL.iterrows():
                                #Se agrega la fila nueva para cada servicio
                                OT_actual.Fila_Nueva(IDs["Planes"]["servicios"])
                                #Se ingresan los datos en cada campo de la pestaña servicio
                                OT_actual.Ingresar_Datos(Campos_Servicio, IDs["Planes"]["servicios"], OT_actual, OT_actual.Listado_Servicios[row_servicio], verbose=True)
                            #Guardar OT
                            OT_actual.guardarOT(IDs["Orden_trabajo"])     
                        
                        #%%%% Pestaña de asignaciones
                        if (int(Datos['LLENAR_MO'])==True):
                            OT_actual.Cambiar_ventana("Asignaciones", IDs)
                            # #Eliminar todas las asignaciones de la OT
                            OT_actual.Eliminar_Filas("asignaciones", IDs["Asignaciones"], Datos['LABOR_BOT'])
                            #Agregar MO y sus parametros
                            for row_mano_de_obra, Datos_Asignaciones in DF_MANO_DE_OBRA_ACTUAL.iterrows():
                                if not pd.isnull(Datos_Asignaciones["MANO_OBRA"]):
                                    #Se agrega la fila nueva para cada servicio
                                    OT_actual.Fila_Nueva(IDs["Asignaciones"]["asignaciones"])
                                    OT_actual.Ingresar_Datos(Campos_Asignacion, IDs["Asignaciones"]["asignaciones"], OT_actual, OT_actual.Listado_Asignaciones[row_mano_de_obra], verbose=True)   
                            #Guardar OT
                            OT_actual.guardarOT(IDs["Orden_trabajo"])
                        
                    
                    #%%% Flujo de la OT
                    OT_actual.Cambiar_ventana("Orden_trabajo", IDs)
                    OT_actual.Flujo_OT(IDs, Datos["ESTADO_DESEADO"], listado_flujo[Datos["ESTADO_DESEADO"]])
                    
                    #%%% Almacenar en registro 
                    # Se almacena las OT creada/modificada correctamente
                    orden_list_objects.append(OT_actual)
                    Registro = pd.DataFrame({'ID': [Datos["ID"]], 'OT': [OT_actual.OT],'OT_PADRE':[OT_actual.OT_PADRE],'UBICACION':[Datos['TRAMO/TRANSFORMADOR']], 'Registro':[OT_actual.ESTADO]})
                    df_registros = df_registros.append(Registro, ignore_index=True)
                    time_program["end_OT"] = time()
                    time_program["list_time_OT"].append(time_program["end_OT"] - time_program["start_OT"])
                    print("Tiempo ejecución OT: ", round(time_program["list_time_OT"][-1],1))
                    
                #%%% Errores en el ciclo de la OT
                except Exception:
                    Registro_error = pd.DataFrame({'ID': [Datos["ID"]], 'OT': [OT_actual.OT], 'OT_PADRE':[OT_actual.OT_PADRE],'UBICACION':[Datos['TRAMO/TRANSFORMADOR']],'Registro': [OT_actual.ERROR_OT]})
                    df_registros = df_registros.append(Registro_error, ignore_index=True)
                    print(Registro_error['Registro'][0])
                    time_program["end_OT"] = time()
                    time_program["list_time_OT"].append(time_program["end_OT"] - time_program["start_OT"])
                    print("Tiempo ejecución OT: ", round(time_program["list_time_OT"][-1],1))
                    
                    #Reiniciar el navegador en caso de presentarse un error
                    WebUploader.cerrar_navegador()
                    #Recreación de WebUploader
                    WebUploader = WebUploader_Class(Data_config, verbose=False)
                    #Método log_in para iniciar sesión
                    WebUploader.log_in(Data_config, IDs["Login"])
                    #Ingreso a menu OT en MX
                    WebUploader.log_menu(IDs["Menus"]["Menu_OT"])
         
                     
        time_program["end_program"] = time()  
        print("--------------------------------------------------------")
        print("Tiempo ejecucion total: ", round(time_program["end_program"] - time_program["start_program"],1))
        
        #Impresion output    
        name_output = "Output/Debug/Debug_" + datetime.datetime.now().strftime('D%Y_%m_%d_H%H_%M_%S') +".xlsx"
        df_registros.to_excel(name_output)
    
    #Retornar la escritura por consola a la ruta por defecto    
    finally:
            # Restaurar el stdout original
            sys.stdout = stdout_original