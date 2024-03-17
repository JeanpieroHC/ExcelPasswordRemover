import os
import sys
import zipfile


# Verificamos si se proporcionó la ruta
if len(sys.argv) != 2 :
    print("Por favor, proporciona la ruta hacia el archivo excel")
    sys.exit(1)


path = sys.argv[1]
directorioArchivo=os.path.dirname(path)

def descomprimir_archivo(archivo):
    # Ruta del directorio donde se encuentra el archivo

    
    # Abre el archivo 
    with zipfile.ZipFile(archivo, 'r') as zip_ref:
        # Extrae todos los archivos al directorio
        zip_ref.extractall(directorioArchivo+"/excelArchives")
    
    print(f"Se ha descomprimido correctamente {archivo}")
    
descomprimir_archivo(path)