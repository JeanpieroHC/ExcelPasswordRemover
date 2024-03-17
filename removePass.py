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
  
    # Abre el archivo 
    with zipfile.ZipFile(archivo, 'r') as zip_ref:
        # Extrae todos los archivos al directorio
        zip_ref.extractall(directorioArchivo+"/excelArchives")
    
    print(f"Se ha descomprimido correctamente {archivo}")
    
descomprimir_archivo(path)

def desprotegerHojaExcel(path):
    dirHpjas=path+"/excelArchives/xl/worksheets"
    
    archivos_xml=encontrar_archivos_xml(dirHpjas)
    print("Archivos XML encontrados:")
    for archivo in archivos_xml:
        print(archivo)
    
def encontrar_archivos_xml(directorio):
    # Obtener una lista de todos los archivos en el directorio
    archivos = os.listdir(directorio)
    
    # Filtrar solo los archivos XML
    archivos_xml = [archivo for archivo in archivos if archivo.endswith('.xml')]
    
    return archivos_xml

desprotegerHojaExcel(directorioArchivo)