import pandas as pd
import os
from multiprocessing import Pool, cpu_count
from tqdm import tqdm  # Para mostrar el progreso en la terminal

# Ruta a la carpeta con los Excel
carpeta = r"C:\Users\NominaAdmin\Desktop\DBF_TO_EXCEL\PE_EXCEL"
archivo_salida_csv = r"C:\Users\NominaAdmin\Desktop\DBF_TO_EXCEL\PE_EXCEL\PE_combinado.csv"
print("Iniciando combinación de archivos en formato CSV...")

if not os.path.exists(carpeta):
    print(f"ERROR: La carpeta no existe:\n{carpeta}")
    exit()

# Buscar archivos Excel válidos
extensiones_validas = ('.xlsx', '.xls', '.xlsm', '.xlsb')
archivos = sorted([f for f in os.listdir(carpeta) if f.lower().endswith(extensiones_validas)])

if not archivos:
    print("No se encontraron archivos Excel.")
    exit()

print(f"Archivos encontrados: {len(archivos)}")
for a in archivos:
    print(f"- {a}")

# Lista de columnas que deben ser tratadas como cadenas
columnas_como_cadenas = ['NUMCHEQUE', 'PDA_1', 'PDA_2', 'PDA_3', 'PDA_4', 'PDA_5', 'PDA_6', 'PDA_7', 'PDA_8', 'PDA_9', 'PDA_10',
                          'PDA_11', 'PDA_12', 'PDA_13', 'PDA_14', 'PDA_15', 'PDA_16', 'PDA_17', 'PDA_18', 'PDA_19', 'PDA_20',
                          'DDA_1', 'DDA_2', 'DDA_3', 'DDA_4', 'DDA_5', 'DDA_6', 'DDA_7', 'DDA_8', 'DDA_9', 'DDA_10', 'DDA_11',
                          'DDA_12', 'DDA_13', 'DDA_14', 'DDA_15', 'DDA_16', 'DDA_17', 'DDA_18', 'DDA_19', 'DDA_20']


def leer_y_convertir(ruta):
    try:
        # Leer el archivo Excel
        ext = os.path.splitext(ruta)[1].lower()
        if ext == '.xlsx':
            df = pd.read_excel(ruta, engine='openpyxl')
        elif ext == '.xls':
            df = pd.read_excel(ruta, engine='xlrd')
        else:
            df = pd.read_excel(ruta, engine='openpyxl')

        if df.empty:
            print(f"[ADVERTENCIA] Archivo vacío: {os.path.basename(ruta)}")
            return None

        # Convertir todas las columnas necesarias a cadenas
        for columna in columnas_como_cadenas:
            if columna in df.columns:
                df[columna] = df[columna].astype(str)

        print(f"Archivo procesado: {os.path.basename(ruta)}")
        return df

    except Exception as e:
        print(f"[ERROR] Error al procesar el archivo {os.path.basename(ruta)}: {str(e)[:60]}")
        return None


def combinar_archivos(archivos):
    # Usar multiprocessing para procesar los archivos en paralelo
    with Pool(cpu_count() // 2) as pool:  # Usamos la mitad de los núcleos
        resultados = pool.map(leer_y_convertir, archivos)

    # Filtrar los resultados que no sean None (es decir, los que no tuvieron errores)
    return [df for df in resultados if df is not None]


def guardar_csv(datos_combinados, archivo_salida):
    # Si hay datos, escribirlos en partes para evitar sobrecargar la memoria
    if datos_combinados:
        first_chunk = True
        for df in datos_combinados:
            # Si es el primer fragmento, crear el archivo CSV con cabecera
            df.to_csv(archivo_salida, index=False, encoding='utf-8-sig', header=first_chunk, mode='a')
            first_chunk = False  # Después de la primera vez, no se escribe la cabecera
        print(f"\nArchivo CSV guardado en:\n{archivo_salida}")
    else:
        print("No se encontraron datos para guardar.")


def procesar_archivos():
    # Ruta completa de los archivos a procesar
    rutas_archivos = [os.path.join(carpeta, archivo) for archivo in archivos]

    # Combinar los archivos en paralelo
    print("Procesando archivos en paralelo...")
    datos_combinados = combinar_archivos(rutas_archivos)

    # Guardar el CSV final
    guardar_csv(datos_combinados, archivo_salida_csv)


if __name__ == "__main__":
    # Monitoreo del progreso con tqdm
    print("Iniciando procesamiento...")
    with tqdm(total=len(archivos), desc="Procesando archivos") as pbar:
        # Procesar los archivos
        procesar_archivos()
        pbar.update(len(archivos))  # Actualizar barra de progreso una vez que termine

