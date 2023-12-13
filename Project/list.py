from getPath import obtener_rutas

import os
import glob

ruta_carpeta, ruta_destino = obtener_rutas()
# Nombres de carpetas a ignorar
nombres_a_ignorar = ['8W', 'SILSAD', 'TYMSA', 'DHL', 'ESTAFETA', 'REDPACK', 'AEROMEXICO', 'TUM']
ruta_destino = ruta_destino
# Ruta de la carpeta principal
ruta_principal = ruta_carpeta

# Lista para almacenar los nombres de los archivos PDF
excel_files = []

# Recorre todas las carpetas y subcarpetas en la ruta principal
for root, dirs, files in os.walk(ruta_principal):
    # Verifica si los nombres de las carpetas contienen palabras a ignorar
    dirs[:] = [d for d in dirs if all(nombre.lower() not in d.lower() for nombre in nombres_a_ignorar)]

    # Busca los archivos excel en cada carpeta
    excel_files.extend(glob.glob(os.path.join(root, '*.xlsx')))


# Imprime los nombres de los archivos PDF encontrados
#excel_files_with_r = [r"r'" + ruta + r"'" for ruta in excel_files]
#excel_files=excel_files_with_r
#for excel_file in excel_files:
#    print(excel_file)

