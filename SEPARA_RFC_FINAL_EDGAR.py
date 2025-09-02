import os
import re
import fitz  # PyMuPDF

def extraer_rfc_y_comprobante_de_pagina(page_text):
    rfc_regex = r'\b([A-Z&]{3,4}\d{6}[A-Z0-9]{3})\b'
    comprobante_regex = r'\b(\d{1,})\b'  

    rfc_match = re.search(rfc_regex, page_text)
    if rfc_match:
        rfc = rfc_match.group(0)
    else:
        rfc = None

    comprobante_match = re.search(comprobante_regex, page_text)
    if comprobante_match:
        num_comprobante = comprobante_match.group(0)
    else:
        num_comprobante = None

    if not rfc or not num_comprobante:
        raise ValueError("No se pudo encontrar el RFC o el número de comprobante en esta página.")

    return rfc, num_comprobante

def separar_pdf(pdf_path, ruta_guardado):
    # Abrir el PDF
    doc = fitz.open(pdf_path)
    num_paginas = len(doc)

    # Iterar sobre todas las páginas del PDF
    for i in range(num_paginas):
        page = doc.load_page(i)
        page_text = page.get_text("text")

        # Extraer RFC y número de comprobante de la página
        rfc, num_comprobante = extraer_rfc_y_comprobante_de_pagina(page_text)
        
        # Crear la carpeta con el nombre RFC_NumeroDeComprobante en la ruta de destino
        carpeta_destino = os.path.join(ruta_guardado, f"{rfc}_{num_comprobante}")
        if not os.path.exists(carpeta_destino):
            os.makedirs(carpeta_destino)
        
        # Guardar la página en la carpeta correspondiente
        nuevo_pdf_path = os.path.join(carpeta_destino, f"{rfc}_{num_comprobante}.pdf")

        # Crear un nuevo documento PDF y añadir la página
        new_doc = fitz.open()
        new_doc.insert_pdf(doc, from_page=i, to_page=i)
        new_doc.save(nuevo_pdf_path)
        new_doc.close()

        print(f"Página {i+1} guardada en: {nuevo_pdf_path}")
    
    print(f"PDF procesado y separado en carpetas correspondientes.")

# Ruta del archivo PDF original
pdf_path = r"C:\Users\pcc\Documents\TALONES DE PAGO\2024\202418\Recibos_R06_202418_O_1.pdf"

# Ruta de destino para guardar los archivos separados
ruta_guardado = r"C:\Users\pcc\Desktop\Escaner\P2"

separar_pdf(pdf_path, ruta_guardado)
