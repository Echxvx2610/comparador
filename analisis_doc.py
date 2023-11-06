import os
from datetime import datetime

# Nombre del archivo
filename = 'H:\Temporal\Echevarria\SV9_007 (1).xlsx'

# Obtener el tamaño del archivo
size = os.path.getsize(filename)
print(f"Tamaño del archivo: {size} bytes")

# Obtener la marca de tiempo de modificación y convertirla a formato legible
modification_time = os.path.getmtime(filename)
modification_time_str = datetime.fromtimestamp(modification_time).strftime('%Y-%m-%d %H:%M:%S')
print(f"Marca de tiempo de modificación: {modification_time_str}")

# Obtener la marca de tiempo de creación y convertirla a formato legible
creation_time = os.path.getctime(filename)
creation_time_str = datetime.fromtimestamp(creation_time).strftime('%Y-%m-%d %H:%M:%S')
print(f"Marca de tiempo de creación: {creation_time_str}")

# Obtener información detallada del archivo
file_stat = os.stat(filename)
print(f"Información detallada del archivo: {file_stat}")
