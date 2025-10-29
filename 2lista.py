import os
from datetime import datetime

def listar_archivos(directorio):
    # Verificar que el directorio existe
    if not os.path.exists(directorio):
        raise FileNotFoundError(f"El directorio {directorio} no existe.")
    
    # Listar los archivos en el directorio
    archivos = os.listdir(directorio)
    # Generar las rutas completas de los archivos
    rutas_completas = [os.path.join(directorio, archivo) for archivo in archivos]
    
    return rutas_completas

def guardar_lista_en_txt(lista, ruta_txt):
    # Guardar la lista de rutas completas en el archivo de texto
    with open(ruta_txt, 'w') as archivo_txt:
        for item in lista:
            archivo_txt.write(f"{item}\n")

def main():
    # Obtener la fecha de hoy en formato AAAAMMDD
    hoy = datetime.today().strftime('%Y%m%d')
    
    # Directorio base donde buscar la carpeta con el nombre de la fecha
    directorio_base = r''
    
    # Ruta completa del directorio con el nombre de la fecha
    directorio = os.path.join(directorio_base, hoy)  # Combina el directorio base y la fecha
    
    try:
        # Listar archivos y obtener rutas completas
        lista_archivos = listar_archivos(directorio)
        
        # Ruta del archivo de salida en el mismo directorio donde se ejecuta el script
        nombre_archivo = f"lista_{hoy}.txt"
        ruta_txt = os.path.join(os.getcwd(), nombre_archivo)  # Guardar en el directorio actual
        
        # Guardar la lista en un archivo de texto
        guardar_lista_en_txt(lista_archivos, ruta_txt)
        
        print(f"Listado de archivos guardado en: {ruta_txt}")
    except FileNotFoundError as e:
        print(e)

if __name__ == "__main__":
    main()
