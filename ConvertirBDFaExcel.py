#!/usr/bin/env python3
"""
Script para convertir archivos .DBF a Excel (.xlsx)
Requiere: pip install dbfread openpyxl pandas
"""

import os
import sys
from pathlib import Path
import pandas as pd
from dbfread import DBF
import argparse

def dbf_to_excel(dbf_path, excel_path=None, encoding='latin-1'):
    """
    Convierte un archivo DBF a Excel
    
    Args:
        dbf_path (str): Ruta del archivo DBF
        excel_path (str): Ruta del archivo Excel de salida (opcional)
        encoding (str): Codificación del archivo DBF (por defecto latin-1)
    
    Returns:
        str: Ruta del archivo Excel generado
    """
    try:
        # Verificar que el archivo DBF existe
        if not os.path.exists(dbf_path):
            raise FileNotFoundError(f"El archivo {dbf_path} no existe")
        
        # Generar nombre del archivo Excel si no se proporciona
        if excel_path is None:
            excel_path = str(Path(dbf_path).with_suffix('.xlsx'))
        
        print(f"Convirtiendo: {dbf_path} -> {excel_path}")
        
        # Leer el archivo DBF
        table = DBF(dbf_path, encoding=encoding)
        
        # Convertir a DataFrame de pandas
        records = []
        for record in table:
            records.append(dict(record))
        
        if not records:
            print("Advertencia: El archivo DBF está vacío")
            df = pd.DataFrame()
        else:
            df = pd.DataFrame(records)
        
        # Guardar como Excel
        df.to_excel(excel_path, index=False, engine='openpyxl')
        
        print(f"✓ Conversión exitosa: {len(records)} registros convertidos")
        print(f"✓ Archivo guardado en: {excel_path}")
        
        return excel_path
        
    except Exception as e:
        print(f"Error al convertir {dbf_path}: {str(e)}")
        return None

def convert_multiple_dbf(directory, output_dir=None, encoding='latin-1'):
    """
    Convierte todos los archivos DBF en un directorio
    
    Args:
        directory (str): Directorio con archivos DBF
        output_dir (str): Directorio de salida (opcional)
        encoding (str): Codificación de los archivos DBF
    """
    directory = Path(directory)
    
    if not directory.exists():
        print(f"Error: El directorio {directory} no existe")
        return
    
    # Buscar archivos DBF
    dbf_files = list(directory.glob('*.dbf')) + list(directory.glob('*.DBF'))
    
    if not dbf_files:
        print(f"No se encontraron archivos DBF en {directory}")
        return
    
    print(f"Encontrados {len(dbf_files)} archivos DBF")
    
    # Configurar directorio de salida
    if output_dir:
        output_path = Path(output_dir)
        output_path.mkdir(parents=True, exist_ok=True)
    else:
        output_path = directory
    
    # Convertir cada archivo
    successful = 0
    failed = 0
    
    for dbf_file in dbf_files:
        excel_file = output_path / f"{dbf_file.stem}.xlsx"
        
        if dbf_to_excel(str(dbf_file), str(excel_file), encoding):
            successful += 1
        else:
            failed += 1
    
    print(f"\n=== Resumen ===")
    print(f"Archivos convertidos exitosamente: {successful}")
    print(f"Archivos con errores: {failed}")

def main():
    parser = argparse.ArgumentParser(
        description="Convertir archivos DBF a Excel (.xlsx)",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Ejemplos de uso:
    python dbf_to_excel.py archivo.dbf
    python dbf_to_excel.py archivo.dbf -o salida.xlsx
    python dbf_to_excel.py -d carpeta_con_dbf/
    python dbf_to_excel.py -d carpeta_con_dbf/ -o carpeta_salida/
    python dbf_to_excel.py archivo.dbf -e utf-8
        """
    )
    
    # Argumentos mutuamente excluyentes
    group = parser.add_mutually_exclusive_group(required=True)
    group.add_argument('archivo', nargs='?', help='Archivo DBF individual a convertir')
    group.add_argument('-d', '--directorio', help='Directorio con archivos DBF')
    
    parser.add_argument('-o', '--output', help='Archivo o directorio de salida')
    parser.add_argument('-e', '--encoding', default='latin-1', 
                       help='Codificación del archivo DBF (default: latin-1)')
    
    args = parser.parse_args()
    
    try:
        if args.archivo:
            # Convertir archivo individual
            dbf_to_excel(args.archivo, args.output, args.encoding)
        elif args.directorio:
            # Convertir directorio completo
            convert_multiple_dbf(args.directorio, args.output, args.encoding)
            
    except KeyboardInterrupt:
        print("\nProceso cancelado por el usuario")
    except Exception as e:
        print(f"Error inesperado: {e}")

if __name__ == "__main__":
    # =================================================================
    #  AQUÍ ES DONDE CAMBIAS LAS RUTAS - MODO FÁCIL
    # =================================================================
    
    # PASO 1: Pon aquí la ruta COMPLETA de tu archivo .DBF
    archivo_dbf = r"C:\DBF\NO00010.DBF"
    # Agregar esta línea para verificar:
    print(f"¿Existe el archivo? {os.path.exists(archivo_dbf)}")
    # PASO 2: Pon aquí donde quieres guardar el Excel (opcional)
    # Si lo dejas vacío (""), se guardará en la misma carpeta del DBF
    carpeta_destino = r"C:\DBF"
    
    # =================================================================
    # NO TOQUES NADA DE AQUÍ PARA ABAJO
    # =================================================================
    
    # Si especificaste carpeta de destino, crear el nombre completo
    if carpeta_destino:
        import os
        os.makedirs(carpeta_destino, exist_ok=True)  # Crear carpeta si no existe
        nombre_archivo = os.path.basename(archivo_dbf).replace('.dbf', '.xlsx').replace('.DBF', '.xlsx')
        archivo_excel = os.path.join(carpeta_destino, nombre_archivo)
    else:
        archivo_excel = None  # Se guardará junto al DBF original
    
    # Ejecutar la conversión
    print(" Iniciando conversión...")
    print(f" Archivo DBF: {archivo_dbf}")
    if archivo_excel:
        print(f" Se guardará en: {archivo_excel}")
    else:
        print(f" Se guardará junto al archivo original")
    
    dbf_to_excel(archivo_dbf, archivo_excel)
    
    print(" ¡Proceso completado!")
    input("Presiona ENTER para cerrar...")