import pandas as pd
import math
import os
import subprocess
import concurrent.futures

def save_chunk_to_excel(data_dict, chunk_file):
    """
    Guarda en 'chunk_file' un Excel con:
      - OTs = data_dict["OTs"] (filtrado para este hilo)
      - TAREAS, MO, MATERIALES, SERVICIOS = las hojas completas
    """
    with pd.ExcelWriter(chunk_file) as writer:
        data_dict["OTs"].to_excel(writer, sheet_name="OTs", index=False)
        data_dict["TAREAS"].to_excel(writer, sheet_name="TAREAS", index=False)
        data_dict["MO"].to_excel(writer, sheet_name="MO", index=False)
        data_dict["MATERIALES"].to_excel(writer, sheet_name="MATERIALES", index=False)
        data_dict["SERVICIOS"].to_excel(writer, sheet_name="SERVICIOS", index=False)

def worker(thread_id, df_chunk, full_excel):
    """
    Función que ejecuta cada hilo:
      1) Crea un sub-Excel (chunk) con las filas de df_chunk en la hoja 'OTs'.
      2) Llama a Principal.py pasándole el nombre del sub-Excel.
      3) Redirige la salida a un archivo 'Output_thread_X.txt'.
      4) Retorna (return) los nombres de los archivos creados para borrarlos luego.
    """
    # 1) Crear sub-Excel temporal
    chunk_name = f"Datos_entrada_chunk_{thread_id}.xlsx"
    chunk_path = os.path.join("Input", chunk_name)

    # Construye el Excel parcial para este hilo
    chunk_data = {
        "OTs": df_chunk,
        "TAREAS":     full_excel["TAREAS"],
        "MO":         full_excel["MO"],
        "MATERIALES": full_excel["MATERIALES"],
        "SERVICIOS":  full_excel["SERVICIOS"]
    }
    save_chunk_to_excel(chunk_data, chunk_path)

    # 2) Invocar Principal.py con ese sub-Excel
    log_file = f"Output_thread_{thread_id}.txt"
    with open(log_file, "w", encoding="utf-8") as lf:
        subprocess.run(
            ["python", "Main_Cargar_Servicios.py", chunk_name],
            stdout=lf,    # Redirige STDOUT al archivo
            stderr=lf,    # Redirige STDERR al mismo archivo
            text=True
        )

    # 3) Devolvemos los nombres de
    #    - Excel temporal
    #    - Log temporal
    return chunk_path, log_file

def main():
    # 1) Parámetros
    input_excel = "Input/Datos_entrada.xlsx"  # Excel original
    n_threads   = 3

    # 2) Leer todas las hojas del Excel original
    df_ots     = pd.read_excel(input_excel, sheet_name="OTs", dtype=str)
    df_tareas  = pd.read_excel(input_excel, sheet_name="TAREAS", dtype=str)
    df_mo      = pd.read_excel(input_excel, sheet_name="MO", dtype=str)
    df_mat     = pd.read_excel(input_excel, sheet_name="MATERIALES", dtype=str)
    df_serv    = pd.read_excel(input_excel, sheet_name="SERVICIOS", dtype=str)

    # 3) Filtrar las filas donde LABOR_BOT != 'NADA'
    df_ots_valid = df_ots[df_ots["LABOR_BOT"] != "NADA"].copy()
    total_filas  = len(df_ots_valid)
    print(f"Total de filas con LABOR_BOT != 'NADA': {total_filas}")

    # 4) Dividir en n_chunks
    chunk_size = math.ceil(total_filas / n_threads)
    chunks = []
    for i in range(0, total_filas, chunk_size):
        chunk_ots = df_ots_valid.iloc[i : i + chunk_size]
        chunks.append(chunk_ots)

    # 5) Diccionario con las otras hojas (completas)
    full_excel = {
        "TAREAS":     df_tareas,
        "MO":         df_mo,
        "MATERIALES": df_mat,
        "SERVICIOS":  df_serv
    }

    # 6) Lanzar hilos en paralelo
    #    Cada hilo devolverá (chunk_path, log_file) que creó
    with concurrent.futures.ThreadPoolExecutor(max_workers=n_threads) as executor:
        futures = []
        for thread_id, df_chunk in enumerate(chunks, start=1):
            # Envía cada chunk al worker
            future = executor.submit(worker, thread_id, df_chunk, full_excel)
            futures.append(future)

        # 7) Esperamos a que todos terminen y recogemos lo devuelto
        #    (nombres de archivos a borrar)
        archivos_para_borrar = []
        for f in concurrent.futures.as_completed(futures):
            chunk_path, log_file = f.result()  # lo que el worker "return"
            archivos_para_borrar.append(chunk_path)
            archivos_para_borrar.append(log_file)

    # 8) Borrar todos los archivos temporales (Excel y logs),
    #    pero *solo* al final, así te permitieron ver el log en tiempo real.
    for filepath in archivos_para_borrar:
        if os.path.exists(filepath):
            os.remove(filepath)

    print("Proceso multi-hilo finalizado. Se han borrado Excel temporales y logs.")

if __name__ == "__main__":
    main()
