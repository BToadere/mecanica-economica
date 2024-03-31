import shutil

def copiar_archivo(origen, destino='/data'):
    """
    Copia un archivo desde el path de origen al path de destino.

    Argumentos:
    origen (str): La ruta del archivo de origen.

    Retorna:
    bool: True si la copia se realizó con éxito, False de lo contrario.
    """
    try:
        shutil.copy(origen, destino)
        print(f'Archivo copiado de {origen} a {destino}')
        return True
    except FileNotFoundError:
        print(f'Error: No se pudo encontrar el archivo {origen}')
        return False
    except PermissionError:
        print(f'Error: Permiso denegado para copiar el archivo {origen}')
        return False
    except Exception as e:
        print(f'Error desconocido: {e}')
        return False