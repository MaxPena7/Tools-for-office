import fitz  # PyMuPDF
import re
import os
from concurrent.futures import ProcessPoolExecutor

def extract_text_from_page(page):
    text = page.get_text("text")
    return text

def split_and_rename_pdf(pdf_path, output_dir):
    os.makedirs(output_dir, exist_ok=True)

    processed_rfc_comprobante = set()  # Usamos un conjunto para almacenar los RFC y comprobante únicos

    # Usamos multiprocessing para procesar páginas en paralelo
    with ProcessPoolExecutor() as executor:
        futures = []
        with fitz.open(pdf_path) as doc:
            for page_num in range(len(doc)):
                futures.append(executor.submit(process_page, pdf_path, page_num, output_dir, processed_rfc_comprobante))
            
            for future in futures:
                future.result()  # Esperar a que cada tarea termine

def process_page(pdf_path, page_num, output_dir, processed_rfc_comprobante):
    with fitz.open(pdf_path) as doc:
        page = doc.load_page(page_num)
        page_text = extract_text_from_page(page)

        # Obtener el RFC y el número de comprobante para cada página
        rfc = extraer_rfc(page_text)  # Usamos solo el texto de esa página
        comprobante_num = comprobante(page_text)

        # Si no se encuentra RFC o comprobante, usar valores por defecto
        if isinstance(rfc, list) and rfc:  # Si se encuentran RFCs válidos
            rfc_comprobante = f"{rfc[0]}_{comprobante_num}"
        else:
            rfc_comprobante = f"pagina_{page_num+1}_{comprobante_num}"

        # Verificar si ya se ha procesado este RFC_comprobante
        if rfc_comprobante in processed_rfc_comprobante:
            print(f"Página {page_num+1} con RFC {rfc_comprobante} ya ha sido procesada. Omitiendo.")
            return  # Si ya se procesó, omitimos esta página

        # Agregar la combinación RFC_comprobante al conjunto para evitar duplicados
        processed_rfc_comprobante.add(rfc_comprobante)

        # Crear carpeta con el nombre basado en RFC y comprobante
        folder_name = rfc_comprobante
        folder_path = os.path.join(output_dir, folder_name)
        os.makedirs(folder_path, exist_ok=True)

        
        new_pdf_path = os.path.join(folder_path, f"{rfc_comprobante}.pdf")
        
        
        pdf_writer = fitz.open()  
        pdf_writer.insert_pdf(doc, from_page=page_num, to_page=page_num)

        pdf_writer.save(new_pdf_path)
        print(f"Página {page_num+1} guardada en carpeta '{folder_path}' como: {new_pdf_path}")

def extraer_rfc(page_text):
    
    rfc_pattern = r'\b[A-Z&Ñ]{4}\d{6}[A-Z0-9]{3}\b'
    
    rfc_encontrados = re.findall(rfc_pattern, page_text)
    rfc_13_caracteres = [rfc for rfc in rfc_encontrados if len(rfc) == 13]
    
    if rfc_13_caracteres:
        return rfc_13_caracteres  

    return "No se encontró un RFC de 13 caracteres en el documento."  

def comprobante(page_text):
    # Buscar el número de comprobante en el texto (por ejemplo, una cadena numérica de 8 dígitos)
    match = re.search(r'\b[0-9]{8}\b', page_text)
    if match:
        return match.group(0)  # Retorna el número de comprobante encontrado
    else:
        return "sin_comprobante"  # Nombre por defecto si no se encuentra

# Bloque protegido para Windows
if __name__ == "__main__":
    # Ruta de prueba
    pdf_path = r"C:\Users\Maxruso7\Downloads\Recibos_R06_202401_O_1.pdf"
    output_dir = r"C:\Users\Maxruso7\Downloads\TALONES SEPARADOS"
    split_and_rename_pdf(pdf_path, output_dir)