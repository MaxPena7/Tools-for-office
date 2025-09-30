import os
import pandas as pd
from dbfread import DBF, FieldParser

# =================================================
# RUTAS ACTUALIZADAS
# =================================================
carpeta_dbf = r"C:\Users\NominaAdmin\Desktop\DBF_TO_EXCEL\NO"  # Ruta de los archivos DBF
carpeta_excel = r"C:\Users\NominaAdmin\Desktop\DBF_TO_EXCEL\NO_EXCEL"  # Ruta donde se guardarán los archivos Excel

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

def convertir_dbf_a_excel(carpeta_dbf, carpeta_excel, encoding="latin-1"):
    archivos_dbf = [f for f in os.listdir(carpeta_dbf) if f.lower().endswith(".dbf")]

    if not archivos_dbf:
        print("No se encontraron archivos DBF en la carpeta.")
        return

    print(f"Se encontraron {len(archivos_dbf)} archivos DBF.")
    
    # Crear carpeta de salida si no existe
    os.makedirs(carpeta_excel, exist_ok=True)

    for archivo in archivos_dbf:
        ruta = os.path.join(carpeta_dbf, archivo)
        print(f"Leyendo: {ruta}")
        tabla = DBF(ruta, encoding=encoding, parserclass=TextFieldParser, load=True)
        df = pd.DataFrame(iter(tabla))

        if df.empty:
            print(f"Advertencia: {archivo} está vacío, se omite.")
            continue

        # Nombre del archivo Excel basado en el nombre del archivo DBF
        nombre_excel = archivo.replace('.dbf', '.xlsx').replace('.DBF', '.xlsx')
        ruta_excel = os.path.join(carpeta_excel, nombre_excel)

        # Guardar como Excel
        df.to_excel(ruta_excel, index=False, engine='openpyxl')
        print(f"Archivo Excel guardado: {ruta_excel}")

    print("\n✅ Conversión completa.")

if __name__ == "__main__":
    convertir_dbf_a_excel(carpeta_dbf, carpeta_excel)
