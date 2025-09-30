import os
import pandas as pd
from dbfread import DBF, FieldParser

# =================================================
# EDITA ESTA VARIABLE CON LA RUTA DE TU CARPETA
# =================================================
carpeta_dbf = r"C:\Users\NominaAdmin\Desktop\DBF_TO_EXCEL\PE"
archivo_csv_salida = r"C:\Users\NominaAdmin\Desktop\DBF_TO_EXCEL\PE_JUNTOS.csv"

# =================================================
# NO MODIFIQUES DE AQUÍ PARA ABAJO
# =================================================

# Parser personalizado para evitar errores con valores numéricos inválidos
class TextFieldParser(FieldParser):
    def parseN(self, field, data):
        try:
            return data.strip().decode(errors='ignore')
        except:
            return str(data)

def convertir_todos_dbf_a_csv(carpeta_dbf, archivo_csv_salida, encoding="latin-1"):
    archivos_dbf = [f for f in os.listdir(carpeta_dbf) if f.lower().endswith(".dbf")]

    if not archivos_dbf:
        print("No se encontraron archivos DBF en la carpeta.")
        return

    print(f"Se encontraron {len(archivos_dbf)} archivos DBF.")
    primera = True  # Para escribir encabezados solo una vez

    with open(archivo_csv_salida, "w", encoding="utf-8", newline="") as salida:
        for archivo in archivos_dbf:
            ruta = os.path.join(carpeta_dbf, archivo)
            print(f"Leyendo: {ruta}")
            tabla = DBF(ruta, encoding=encoding, parserclass=TextFieldParser, load=True)
            df = pd.DataFrame(iter(tabla))

            if df.empty:
                print(f"Advertencia: {archivo} está vacío, se omite.")
                continue

            df.to_csv(salida, mode="a", header=primera, index=False)
            primera = False  # Después del primer archivo ya no escribe encabezados

    print(f"\n✅ Conversión completa. Archivo combinado guardado en:\n{archivo_csv_salida}")

if __name__ == "__main__":
    convertir_todos_dbf_a_csv(carpeta_dbf, archivo_csv_salida)
