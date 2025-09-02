import pandas as pd
import os

# 1. Configuración
carpeta = r'C:\Users\Maxruso7\Desktop\REPORTE DE FUPS ACUMULADO'
archivo_salida = 'combinado.xlsx'

# 2. Verificar que la carpeta existe
if not os.path.exists(carpeta):
    raise FileNotFoundError(f"La carpeta no existe: {carpeta}")

# 3. Encontrar archivos Excel (incluyendo varias extensiones)
extensiones_validas = ('.xlsx', '.xls', '.xlsm', '.xlsb')
archivos = [f for f in os.listdir(carpeta) 
           if f.lower().endswith(extensiones_validas)]

if not archivos:
    raise FileNotFoundError("No se encontraron archivos Excel en la carpeta")

# 4. Función para leer archivos con el motor adecuado
def leer_excel(ruta):
    ext = os.path.splitext(ruta)[1].lower()
    if ext == '.xlsx':
        return pd.read_excel(ruta, engine='openpyxl')
    elif ext == '.xls':
        return pd.read_excel(ruta, engine='xlrd')
    else:  # Para .xlsm y .xlsb
        return pd.read_excel(ruta, engine='openpyxl')

# 5. Procesar archivos
datos_combinados = pd.DataFrame()
errores = []

for archivo in archivos:
    try:
        ruta = os.path.join(carpeta, archivo)
        df = leer_excel(ruta)
        datos_combinados = pd.concat([datos_combinados, df], ignore_index=True)
        print(f"✔ {archivo:.<50} [Registros: {len(df):>5}]")
    except Exception as e:
        errores.append((archivo, str(e)))
        print(f"✖ {archivo:.<50} [Error: {str(e)[:30]}...]")

# 6. Guardar resultados
if not datos_combinados.empty:
    ruta_salida = os.path.join(carpeta, archivo_salida)
    datos_combinados.to_excel(ruta_salida, index=False, engine='openpyxl')
    print(f"\n✅ Resultado:")
    print(f"- Archivos procesados: {len(archivos)}")
    print(f"- Archivos con errores: {len(errores)}")
    print(f"- Registros combinados: {len(datos_combinados)}")
    print(f"- Guardado en: {ruta_salida}")
    
    if errores:
        print("\n⚠ Archivos con errores:")
        for archivo, error in errores:
            print(f"- {archivo}: {error}")
else:
    print("No se pudo combinar ningún archivo")