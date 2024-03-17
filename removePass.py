import os
import sys
import zipfile
import xml.etree.ElementTree as ET

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
    ver_contenido_xml(dirHpjas+"/"+archivos_xml[0])
    nuevo_contenido = {
        './/{http://schemas.openxmlformats.org/spreadsheetml/2006/main}sheetProtection': '0'
    }
    
    for archivo in archivos_xml:
        cambiar_contenido_xml(dirHpjas+"/"+archivo,nuevo_contenido)
        
def desprotegerLibroExcel(path):
    pathLibro=path+"/excelArchives/xl/workbook.xml"
    
    nuevo_contenido = {
        './/{http://schemas.openxmlformats.org/spreadsheetml/2006/main}workbookProtection': '0'
    }

    cambiar_contenido_xml(pathLibro,nuevo_contenido)
        
def encontrar_archivos_xml(directorio):
    # Obtener una lista de todos los archivos en el directorio
    archivos = os.listdir(directorio)
    
    # Filtrar solo los archivos XML
    archivos_xml = [archivo for archivo in archivos if archivo.endswith('.xml')]
    
    return archivos_xml
def ver_contenido_xml(archivo):
    tree = ET.parse(archivo)
    root = tree.getroot()
    # Imprimir el contenido del archivo XML
    ET.dump(root)

# Función para cambiar el contenido de un archivo XML
def cambiar_contenido_xml(archivo, nuevo_contenido):
    primer_elemento = next(iter(nuevo_contenido.items()))
    print(primer_elemento)
    tree = ET.parse(archivo)
    root = tree.getroot()
    elemento_modificar=root.find(primer_elemento[0])
    try:
        # Manejo de cualquier otra excepción no manejada anteriormente
        # ...
        atributos = elemento_modificar.attrib

        # Imprimir los atributos
        for nombre_atributo, valor_atributo in atributos.items():
            if nombre_atributo!='password' and nombre_atributo!='workbookPassword':
                elemento_modificar.set(nombre_atributo, primer_elemento[1])
        # Guardar el árbol XML modificado en el archivo
    except Exception as e:
        print(e)
    
    tree.write(archivo)

desprotegerHojaExcel(directorioArchivo)
desprotegerLibroExcel(directorioArchivo)

def recomprimir_xlsx_xlsm(directorio):
    # Nombre del archivo ZIP resultante
    nuevo_zip = os.path.join(directorio, 'nuevo_archivo.xlsm')

    # Abre un nuevo archivo ZIP en modo de escritura
    with zipfile.ZipFile(nuevo_zip, 'w') as zip_ref:
        # Recorre todos los archivos en el directorio
        zip_ref.write(directorio+'/nuevo_archivo.xlsm', 'nuevo_archivo.xlsm')
        for carpeta_actual, _, archivos in os.walk(directorio+"/excelArchives"):
            for archivo in archivos:
                # Ruta completa del archivo a agregar al ZIP
                ruta_archivo = os.path.join(carpeta_actual, archivo)
                # Ruta relativa del archivo dentro del ZIP
                ruta_zip = os.path.relpath(ruta_archivo, directorio+"/excelArchives")
                # Agrega el archivo al ZIP
                zip_ref.write(ruta_archivo, ruta_zip)

    print(f"Se ha recomprimido correctamente en {nuevo_zip}")



recomprimir_xlsx_xlsm(directorioArchivo)