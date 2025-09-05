import pandas as pd
import os

# Ruta a la carpeta con los Excel
carpeta = r'C:\ANEXO VI'
archivo_salida_csv = 'Acumulado_Anexo_VI_completo.csv'

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

# Función para leer cada archivo
def leer_excel(ruta):
    ext = os.path.splitext(ruta)[1].lower()
    if ext == '.xlsx':
        return pd.read_excel(ruta, engine='openpyxl')
    elif ext == '.xls':
        return pd.read_excel(ruta, engine='xlrd')
    else:
        return pd.read_excel(ruta, engine='openpyxl')

# Combinar todos en un solo DataFrame
datos_combinados = pd.DataFrame()
errores = []

for archivo in archivos:
    ruta = os.path.join(carpeta, archivo)
    try:
        df = leer_excel(ruta)

        if df.empty:
            print(f"[ADVERTENCIA] Archivo vacío: {archivo}")
            continue

        datos_combinados = pd.concat([datos_combinados, df], ignore_index=True)
        print(f"[OK] {archivo:<40} Registros: {len(df)}")

    except Exception as e:
        errores.append((archivo, str(e)))
        print(f"[ERROR] {archivo:<40} {str(e)[:60]}")

# Guardar como CSV
if not datos_combinados.empty:
    ruta_salida = os.path.join(carpeta, archivo_salida_csv)
    datos_combinados.to_csv(ruta_salida, index=False, encoding='utf-8-sig')
    print(f"\nArchivo CSV guardado en:\n{ruta_salida}")
    print(f"Total de registros combinados: {len(datos_combinados)}")

    if errores:
        print("\nArchivos con errores:")
        for archivo, error in errores:
            print(f"- {archivo}: {error}")
else:
    print("No se pudo combinar ningún archivo.")
