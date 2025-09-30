import os
import pandas as pd
from dbfread import DBF, FieldParser

# =================================================
# EDITA ESTAS DOS VARIABLES CON TUS RUTAS
# =================================================

# Ruta del archivo DBF que quieres convertir
ruta_dbf = r"C:\Users\NominaAdmin\Desktop\NOMINA\CHAVEZ\NOMINA\PE00010.DBF"

# Ruta donde quieres guardar el archivo Excel (.xlsx)
ruta_excel = r"C:\Users\NominaAdmin\Desktop\DBF_TO_EXCEL"

# =================================================
# NO MODIFIQUES DE AQUÍ PARA ABAJO
# =================================================

# Parser personalizado: evita que los campos numéricos se conviertan a float
class TextFieldParser(FieldParser):
    def parseN(self, field, data):  
        try:
            return data.strip().decode(errors='ignore')
        except:
            return str(data)

def convertir_dbf_a_excel(ruta_dbf, ruta_excel, encoding='latin-1'):
    try:
        if not os.path.exists(ruta_dbf):
            print(f"Error: El archivo DBF no existe:\n{ruta_dbf}")
            return

        print(f"Leyendo archivo DBF: {ruta_dbf}")
        tabla = DBF(ruta_dbf, encoding=encoding, parserclass=TextFieldParser, load=True)
        
        registros = list(tabla)
        df = pd.DataFrame(registros)

        # Crear carpeta si no existe
        os.makedirs(os.path.dirname(ruta_excel), exist_ok=True)

        # Guardar como Excel
        df.to_excel(ruta_excel, index=False, engine='openpyxl')

        print(f"Conversión exitosa. Archivo guardado en: {ruta_excel}")

    except Exception as e:
        print(f"Error durante la conversión:\n{e}")

if __name__ == "__main__":
    # Construir el nombre completo del archivo Excel
    nombre_archivo = os.path.basename(ruta_dbf).replace('.dbf', '.xlsx').replace('.DBF', '.xlsx')
    archivo_excel_completo = os.path.join(ruta_excel, nombre_archivo)

    convertir_dbf_a_excel(ruta_dbf, archivo_excel_completo)
    input("Presiona ENTER para cerrar...")
