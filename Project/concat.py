import PySimpleGUI as sg
import os
import pandas as pd

def obtener_ruta_carpeta():
    sg.theme('TealMono')

    # Definir el diseño de la ventana para seleccionar la carpeta a procesar
    layout = [
        [sg.Text('Selecciona carpeta a procesar')],
        [sg.InputText(key='folder_path'), sg.FolderBrowse()],
        [sg.Button('Aceptar')]
    ]

    # Crear la ventana
    window = sg.Window('Seleccionar carpeta', layout)

    # Leer eventos y datos ingresados en la ventana
    while True:
        event, values = window.read()
        if event == sg.WIN_CLOSED or event == 'Aceptar':
            ruta_carpeta = values['folder_path']
            break

    window.close()  # Cerrar la ventana
    return ruta_carpeta

def obtener_ruta_destino():
    sg.theme('LightBlue1')

    # Definir el diseño de la ventana para seleccionar la carpeta de destino
    layout = [
        [sg.Text('Selecciona carpeta de destino')],
        [sg.InputText(key='destination_path'), sg.FolderBrowse()],
        [sg.Button('Aceptar')]
    ]

    # Crear la ventana
    window = sg.Window('Seleccionar carpeta de destino', layout)

    # Leer eventos y datos ingresados en la ventana
    while True:
        event, values = window.read()
        if event == sg.WIN_CLOSED or event == 'Aceptar':
            ruta_destino = values['destination_path']
            break

    window.close()  # Cerrar la ventana
    return ruta_destino

def combine_excel_files(folder_path, output_file):
    file_list = os.listdir(folder_path)
    excel_files = [f for f in file_list if f.endswith('.xlsx') or f.endswith('.xls')]
    
    combined_data = pd.DataFrame()
    
    for file in excel_files:
        # Leer cada archivo de Excel en un DataFrame separado
        file_path = os.path.join(folder_path, file)
        df = pd.read_excel(file_path)
        
        # Combinar el DataFrame 
        combined_data = pd.concat([combined_data, df], ignore_index=True)
    
    output_path = os.path.join(folder_path, output_file)
    combined_data.to_excel(output_path, index=False)
    print("Completado. El resultado se ha guardado en:", output_path)

def procesar_archivos():
    # Obtener la ruta de la carpeta a procesar
    ruta_carpeta = obtener_ruta_carpeta()

    # Obtener la ruta de la carpeta de destino
    ruta_destino = obtener_ruta_destino()

    # Combinar archivos Excel en la carpeta seleccionada
    output_file = 'resultado.xlsx'  # Nombre del archivo de salida
    combine_excel_files(ruta_carpeta, output_file)
    
    print(f"Carpeta de destino seleccionada: {ruta_destino}")

if __name__ == "__main__":
    procesar_archivos()
