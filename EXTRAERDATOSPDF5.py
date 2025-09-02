import fitz  # PyMuPDF
import pandas as pd
import re

# Función para extraer texto de un PDF y organizarlo en columnas
def extraer_texto_por_columnas(pdf_path):
    doc = fitz.open(pdf_path)
    datos_totales = []  # Lista para almacenar los datos de todas las páginas
    folios_fiscales = []  # Lista para almacenar los folios fiscales
    
    # Procesar cada página del PDF
    for num in range(len(doc)):
        pagina = doc.load_page(num)
        texto = pagina.get_text("text")
        lineas = texto.split("\n")  # Dividir el texto en líneas
        
        # Extraer el folio fiscal usando la expresión regular
        folio_fiscal = extraer_folio_fiscal(texto)
        folios_fiscales.append(folio_fiscal)
        
        # Agregar las líneas a la lista de datos
        datos_totales.append(lineas)
        
        # Mostrar progreso cada 100 páginas
        if (num + 1) % 100 == 0:
            print(f"Procesadas {num + 1} páginas...")
    
    doc.close()
    return datos_totales, folios_fiscales

# Función para extraer el folio fiscal usando una expresión regular
def extraer_folio_fiscal(texto):
    patron = r"[A-F0-9]{8}-[A-F0-9]{4}-[A-F0-9]{4}-[A-F0-9]{4}-[A-F0-9]{12}"
    match = re.search(patron, texto)
    return match.group(0) if match else "NO ENCONTRADO"

# Función para guardar solo las columnas específicas en un archivo Excel
def guardar_en_excel(datos_totales, folios_fiscales, output_path):
    # Columnas específicas que deseas guardar (1, 2, 9, 17, 21, 26, 34)
    columnas_a_guardar = [1, 2, 9, 17, 21, 26, 34]
    
    # Crear un diccionario para almacenar los datos en columnas
    datos_organizados = {}
    for i, col in enumerate(columnas_a_guardar):
        datos_organizados[f"Columna {col}"] = [
            pagina[col - 1] if col - 1 < len(pagina) else "" for pagina in datos_totales
        ]
    
    # Agregar la columna de folios fiscales
    datos_organizados["Folio Fiscal"] = folios_fiscales
    
    # Crear un DataFrame con los datos organizados
    df = pd.DataFrame(datos_organizados)
    
    # Guardar en Excel
    df.to_excel(output_path, index=False)

# Ruta del archivo PDF y salida de Excel
pdf_file = r"C:\Users\Maxruso7\Downloads\Recibos_R06_202401_O_1.pdf"  # Cambia esto por la ruta de tu PDF
output_file = "datos_extraidos.xlsx"

# Extraer el texto del PDF y organizarlo en columnas
print("Extrayendo texto del PDF...")
datos_extraidos, folios_fiscales = extraer_texto_por_columnas(pdf_file)

# Guardar solo las columnas específicas en un archivo Excel
print("Guardando datos en Excel...")
guardar_en_excel(datos_extraidos, folios_fiscales, output_file)

print(f"Proceso completado. Datos guardados en {output_file}")