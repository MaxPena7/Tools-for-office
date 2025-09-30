import os
import pandas as pd
from sqlalchemy import create_engine
import pymysql
import multiprocessing

# Función que procesa un solo archivo (la tarea de cada "trabajador")
def procesar_archivo(ruta_archivo_completa, db_params):
    """Sube un solo archivo de Excel a una tabla de MySQL."""
    try:
        # Conexión a la base de datos dentro de cada proceso
        connection_string = f'mysql+pymysql://{db_params["user"]}:{db_params["password"]}@{db_params["host"]}/{db_params["database"]}'
        engine = create_engine(connection_string)

        nombre_archivo = os.path.basename(ruta_archivo_completa)
        nombre_tabla = os.path.splitext(nombre_archivo)[0]
        nombre_tabla = nombre_tabla.replace('-', '_').replace(' ', '_').replace('.', '_')
        nombre_tabla = nombre_tabla.lower()

        df = pd.read_excel(ruta_archivo_completa)

        print(f"Procesando archivo: {nombre_archivo} en un proceso paralelo...")
        
        # Subir el DataFrame a MySQL de forma masiva
        df.to_sql(name=nombre_tabla, con=engine, index=False, if_exists='replace', chunksize=1000)
        
        print(f"¡Archivo '{nombre_archivo}' subido a la tabla '{nombre_tabla}'!")
        
        engine.dispose()
        
    except Exception as e:
        print(f"Error al procesar el archivo '{nombre_archivo}': {e}")

def subir_archivos_a_mysql_paralelo(ruta_carpeta, db_params):
    """
    Coordina el procesamiento de múltiples archivos en paralelo.
    """
    # Recorrer la carpeta y obtener la lista de archivos a procesar
    archivos_a_procesar = []
    for nombre_archivo in os.listdir(ruta_carpeta):
        if nombre_archivo.endswith('.xlsx') or nombre_archivo.endswith('.xls'):
            ruta_completa = os.path.join(ruta_carpeta, nombre_archivo)
            archivos_a_procesar.append((ruta_completa, db_params))
    
    # Crea un pool de procesos. El número de procesos se basa en los núcleos del CPU.
    # Puedes ajustar este número si es necesario.
    num_procesos = os.cpu_count()
    print(f"Iniciando el procesamiento en paralelo con {num_procesos} procesos...")

    # Usa un Pool de procesos para ejecutar la función en paralelo
    with multiprocessing.Pool(processes=num_procesos) as pool:
        pool.starmap(procesar_archivo, archivos_a_procesar)

    print("Proceso de subida de archivos completado.")

# --- Configura tus datos aquí ---
carpeta_a_procesar = r"C:\Users\NominaAdmin\Desktop\ANEXOS\ANEXOS EXTRAIDOS (2023, 2024, 20205(Q16))\PRUEBA"
db_parametros = {
    'user': "root", 
    'password': "Max20135200",
    'host': "localhost",
    'database': "nomina_db"
}

# Llama a la función principal para iniciar el proceso paralelo
if __name__ == '__main__':
    subir_archivos_a_mysql_paralelo(carpeta_a_procesar, db_parametros)