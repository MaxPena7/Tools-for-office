import pandas as pd

try:
    # Intenta con el motor xlrd para archivos .xls antiguos
    print("Intentando abrir con motor xlrd...")
    df = pd.read_excel(r'C:\Users\Maxruso7\Downloads\PERDEDsaycop.xls', engine='xlrd')
    print(f"Archivo leído correctamente con xlrd. Dimensiones: {df.shape}")
    
except Exception as e:
    print(f"Error con motor xlrd: {e}")
    try:
        # Intenta con el motor openpyxl
        print("Intentando abrir con motor openpyxl...")
        df = pd.read_excel(r'C:\Users\Maxruso7\Downloads\PERDEDsaycop.xls', engine='openpyxl')
        print(f"Archivo leído correctamente con openpyxl. Dimensiones: {df.shape}")
    except Exception as e:
        print(f"Error con motor openpyxl: {e}")
        
        # Si ambos fallan, verifica si realmente es un archivo CSV
        try:
            print("Intentando abrir como CSV...")
            df = pd.read_csv(r'C:\Users\Maxruso7\Downloads\PERDEDsaycop.xls')
            print(f"Archivo leído correctamente como CSV. Dimensiones: {df.shape}")
        except Exception as e:
            print(f"Error al abrir como CSV: {e}")
            print("No fue posible abrir el archivo con ningún método disponible.")

# Si llegamos aquí y df existe, mostrar primeras filas
try:
    print("\nPrimeras 2 filas del archivo:")
    print(df.head(2))
except:
    print("No se pudo mostrar el contenido del archivo.")