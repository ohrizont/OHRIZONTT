import os
from datetime import datetime

def listar_archivos(directorio):
    """
    Lista los archivos en el directorio especificado, filtrando para incluir
    solo aquellos que contienen 'indicadores' en su nombre.

    Args:
        directorio (str): El directorio a listar.

    Returns:
        list: Una lista de rutas completas de archivos que contienen 'indicadores' en su nombre.
    """
    # Verificar que el directorio existe
    if not os.path.exists(directorio):
        raise FileNotFoundError(f"El directorio {directorio} no existe.")
    
    # Listar los archivos en el directorio
    archivos = os.listdir(directorio)
    
    # Filtrar los archivos que contienen "indicadores" en su nombre
    archivos_filtrados = [archivo for archivo in archivos if "indicadores" in archivo]
    
    # Generar las rutas completas de los archivos filtrados
    rutas_completas = [os.path.join(directorio, archivo) for archivo in archivos_filtrados]
    
    return rutas_completas

def guardar_lista_en_txt(lista, ruta_txt):
    """
    Guarda la lista de rutas completas en un archivo de texto.

    Args:
        lista (list): La lista de rutas completas a guardar.
        ruta_txt (str): La ruta al archivo de texto donde guardar la lista.
    """
    # Guardar la lista de rutas completas en el archivo de texto
    with open(ruta_txt, 'w') as archivo_txt:
        for item in lista:
            archivo_txt.write(f"{item}\n")

def main():
    """
    Función principal que obtiene la fecha actual, construye la ruta al directorio,
    lista los archivos que contienen 'indicadores' en su nombre, y guarda la lista
    en un archivo de texto.
    """
    # Obtener la fecha de hoy en formato AAAAMMDD
    hoy = datetime.today().strftime('%Y%m%d')
    
    # Directorio base donde buscar la carpeta con el nombre de la fecha
    directorio_base = r''  # Reemplaza esto con tu directorio base si es necesario
    
    # Ruta completa del directorio con el nombre de la fecha
    directorio = os.path.join(directorio_base, hoy)  # Combina el directorio base y la fecha
    
    try:
        # Listar archivos y obtener rutas completas
        lista_archivos = listar_archivos(directorio)
        
        # Ruta del archivo de salida en el mismo directorio donde se ejecuta el script
        nombre_archivo = f"lista_indicadores_{hoy}.txt"
        ruta_txt = os.path.join(os.getcwd(), nombre_archivo)  # Guardar en el directorio actual
        
        # Guardar la lista en un archivo de texto
        guardar_lista_en_txt(lista_archivos, ruta_txt)
        
        print(f"Listado de archivos guardado en: {ruta_txt}")
    except FileNotFoundError as e:
        print(e)

if __name__ == "__main__":
    main()
