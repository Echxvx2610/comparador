import os
import getpass
import datetime

def obtener_informacion_archivo(ruta_archivo):
    # Información básica del archivo
    informacion = {
        'ruta': ruta_archivo,
        'fecha_creacion': datetime.datetime.fromtimestamp(os.path.getctime(ruta_archivo)),
        'ultima_modificacion': datetime.datetime.fromtimestamp(os.path.getmtime(ruta_archivo)),
        'tamaño': os.path.getsize(ruta_archivo),
    }

    # Intentar obtener el nombre del usuario actual en sistemas Windows
    try:
        informacion['usuario_modificacion'] = getpass.getuser()
    except Exception as e:
        informacion['usuario_modificacion'] = "No se pudo obtener la información del usuario"

    return informacion

# Ruta del archivo que quieres analizar
ruta_archivo = r'H:\Temporal\Echevarria\32194-1G\SV9_007 (1).xlsx'
# Obtener información del archivo
informacion_archivo = obtener_informacion_archivo(ruta_archivo)

# Imprimir la información
print("Información del archivo:")
print("Ruta:", informacion_archivo['ruta'])
print("Fecha de creación:", informacion_archivo['fecha_creacion'])
print("Última modificación:", informacion_archivo['ultima_modificacion'])
print("Tamaño:", informacion_archivo['tamaño'], "bytes")
print("Usuario que lo modificó:", informacion_archivo['usuario_modificacion'])
