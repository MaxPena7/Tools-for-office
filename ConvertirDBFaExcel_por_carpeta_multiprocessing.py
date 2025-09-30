import os
import pandas as pd
from dbfread import DBF, FieldParser
from multiprocessing import Pool

# =================================================
# RUTAS ACTUALIZADAS
# =================================================
carpeta_dbf = r"C:\Users\NominaAdmin\Desktop\DBF_TO_EXCEL\PE"  # Ruta de los archivos DBF
carpeta_excel = r"C:\Users\NominaAdmin\Desktop\DBF_TO_EXCEL\PE_EXCEL"  # Ruta donde se guardarán los archivos Excel

# =================================================
# NO MODIFIQUES DE AQUÍ PARA ABAJO
# =================================================

# Parser personalizado para evitar errores con valores numéricos inválidos
class TextFieldParser(FieldParser):
    def parseN(self, field, data):
        try:
            return data.strip().decode(errors='ignore') if data else ''
        except:
            return str(data)

def convertir_dbf_a_excel(archivo_dbf):
    """ Convierte un archivo DBF en un archivo Excel. """
    ruta_dbf = os.path.join(carpeta_dbf, archivo_dbf)
    tabla = DBF(ruta_dbf, encoding="latin-1", parserclass=TextFieldParser, load=True)
    df = pd.DataFrame(iter(tabla))

    if df.empty:
        print(f"Advertencia: {archivo_dbf} está vacío, se omite.")
        return

    # Nombre del archivo Excel basado en el nombre del archivo DBF
    nombre_excel = archivo_dbf.replace('.dbf', '.xlsx').replace('.DBF', '.xlsx')
    ruta_excel = os.path.join(carpeta_excel, nombre_excel)

    # Guardar como Excel
    df.to_excel(ruta_excel, index=False, engine='openpyxl')
    print(f"Archivo Excel guardado: {ruta_excel}")

def procesar_archivos_dbf():
    """ Lee todos los archivos DBF en la carpeta y los convierte en Excel. """
    archivos_dbf = [f for f in os.listdir(carpeta_dbf) if f.lower().endswith(".dbf")]
    if not archivos_dbf:
        print("No se encontraron archivos DBF en la carpeta.")
        return

    # Crear la carpeta de salida si no existe
    os.makedirs(carpeta_excel, exist_ok=True)

    # Usar multiprocessing para convertir los archivos en paralelo
    with Pool() as pool:
        pool.map(convertir_dbf_a_excel, archivos_dbf)

    print("\n✅ Conversión completa.")

if __name__ == "__main__":
    procesar_archivos_dbf()
