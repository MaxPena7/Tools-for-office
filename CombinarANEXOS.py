import pandas as pd
import os

# 1. Ruta de la carpeta con tus archivos Excel
carpeta = r'C:\Users\NÓMINA\Desktop\ANEXOS\ANEXOS EXTRAIDOS (2023, 2024, 20205(Q16))\ANEXO VI'
archivo_salida = 'Acumulado_Anexo_VI_(202301-202516).xlsx'

print("Iniciando combinación de archivos Excel...")

# 2. Verificar existencia de la carpeta
if not os.path.exists(carpeta):
    print(f"ERROR: La carpeta no existe:\n{carpeta}")
    exit()

# 3. Obtener y ordenar archivos Excel
extensiones_validas = ('.xlsx', '.xls', '.xlsm', '.xlsb')
archivos = sorted([f for f in os.listdir(carpeta) if f.lower().endswith(extensiones_validas)])

if not archivos:
    print("No se encontraron archivos Excel en la carpeta.")
    exit()

print(f"Archivos encontrados: {len(archivos)}")
for a in archivos:
    print(f"- {a}")

# 4. Función para leer archivos según la extensión
def leer_excel(ruta):
    ext = os.path.splitext(ruta)[1].lower()
    if ext == '.xlsx':
        return pd.read_excel(ruta, engine='openpyxl')
    elif ext == '.xls':
        return pd.read_excel(ruta, engine='xlrd')
    else:  # .xlsm, .xlsb
        return pd.read_excel(ruta, engine='openpyxl')

# 5. Combinar los archivos
datos_combinados = pd.DataFrame()
errores = []

for i, archivo in enumerate(archivos):
    ruta = os.path.join(carpeta, archivo)
    try:
        df = leer_excel(ruta)

        if df.empty:
            print(f"ADVERTENCIA: {archivo:<40} está vacío. Se omite.")
            continue

        if i == 0:
            datos_combinados = df.copy()
        else:
            datos_combinados = pd.concat([datos_combinados, df], ignore_index=True)

        print(f"[OK]   {archivo:<40} Registros: {len(df):>5}")

    except Exception as e:
        errores.append((archivo, str(e)))
        print(f"[ERROR] {archivo:<40} Error: {str(e)[:60]}")

# 6. Guardar el archivo combinado
if not datos_combinados.empty:
    ruta_salida = os.path.join(carpeta, archivo_salida)
    datos_combinados.to_excel(ruta_salida, index=False, engine='openpyxl')

    print("\nResumen del proceso:")
    print(f"- Archivos procesados: {len(archivos)}")
    print(f"- Archivos con errores: {len(errores)}")
    print(f"- Registros combinados: {len(datos_combinados)}")
    print(f"- Archivo guardado en: {ruta_salida}")

    if errores:
        print("\nArchivos que no se pudieron procesar:")
        for archivo, error in errores:
            print(f"- {archivo}: {error}")
else:
    print("No se pudo combinar ningún archivo. Todos vacíos o con errores.")
