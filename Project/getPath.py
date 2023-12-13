import PySimpleGUI as sg

def obtener_rutas():
    sg.theme('BrownBlue')  

    layout1 = [
        [sg.Text('Selecciona carpeta con archivos a procesar')],
        [sg.InputText(key='folder_path'), sg.FolderBrowse()],
        [sg.Button('Aceptar')]
    ]

    window1 = sg.Window('Seleccionar carpeta', layout1)

    # Leer eventos y datos ingresados en la ventana 1
    while True:
        event, values = window1.read()
        if event == sg.WIN_CLOSED or event == 'Aceptar':
            ruta_carpeta = values['folder_path']
            break

    window1.close()  

    layout2 = [
        [sg.Text('Selecciona carpeta de destino')],
        [sg.InputText(key='destination_path'), sg.FolderBrowse()],
        [sg.Button('Aceptar')]
    ]

    window2 = sg.Window('Seleccionar carpeta de destino', layout2)

    # Leer eventos y datos ingresados en la ventana 2
    while True:
        event, values = window2.read()
        if event == sg.WIN_CLOSED or event == 'Aceptar':
            ruta_destino = values['destination_path']
            break

    window2.close()  

    return ruta_carpeta, ruta_destino

if __name__ == "__main__":
    r_carpeta, r_destino = obtener_rutas()
    print("Ruta de carpeta con archivos a procesar:", r_carpeta)
    print("Ruta de carpeta de destino:", r_destino)
