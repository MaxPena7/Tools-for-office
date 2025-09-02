import pdfplumber
import re

# Función para extraer el RFC de un archivo PDF
def extraer_rfc(pdf_path):
    # Expresión regular para identificar un RFC de 13 caracteres
    rfc_pattern = r'\b[A-Z&Ñ]{4}\d{6}[A-Z0-9]{3}\b'

    # Abre el archivo PDF
    with pdfplumber.open(pdf_path) as pdf:
        # Itera por cada página del PDF
        for page in pdf.pages:
            # Extrae el texto de la página
            text = page.extract_text()

            if text:
                # Busca todas las coincidencias de RFC en el texto
                rfc_encontrados = re.findall(rfc_pattern, text)
                
                # Filtra los RFCs que tienen exactamente 13 caracteres
                rfc_13_caracteres = [rfc for rfc in rfc_encontrados if len(rfc) == 13]
                
                # Si se encuentran RFCs válidos, los retornamos
                if rfc_13_caracteres:
                    return rfc_13_caracteres

    return "No se encontró un RFC de 13 caracteres en el documento."

# Ruta del archivo PDF
pdf_path = r"C:\Users\Maxruso7\Downloads\AAAA571001HMN_10217775.pdf"  # Cambia esta ruta por la del archivo PDF

# Llamamos a la función y mostramos el resultado
rfc = extraer_rfc(pdf_path)
print("RFC encontrado:", rfc)