import PySimpleGUI as sg
from bom_check import comparador
import os
import subprocess

def main():
    #mostrar todos los temas
    #sg.theme_previewer()
    sg.theme('LightGreen3')
    layout = [
        [sg.Input(default_text='Ruta archivo Syteline',key='-BOM-',enable_events=True,size=(65,10),readonly=True),sg.FileBrowse(file_types=(("Excel files", "*.xlsx"), ("All files", "*.*")),button_text="Cargar BOM",)],
        [sg.Input(default_text='Ruta archivo Placement Flexa ',key='-FLEXA-',enable_events=True,size=(60,10),readonly=True),sg.FileBrowse(file_types=(("Excel files", "*.xlsx"), ("All files", "*.*")),button_text="Cargar Placement",)],
        [sg.Button('Comparar'),sg.Button('Salir')],        
    ]
    
     
    window = sg.Window('....::: Flexa vs BOM :::.....',layout,finalize=True,no_titlebar=False,element_justification='center',icon='comparador\img\compare_4222.ico',keep_on_top=False)
    def flexa_vs_bom():
        comparador(values['-BOM-'],values['-FLEXA-'])
        sg.popup('Comparacion completada con exito!')
    
    def reset():
        window['-BOM-'].update('')
        window['-FLEXA-'].update('')
        
    def open_folder_in_explorer(folder_path):
        if os.path.exists(folder_path):
            try:
                subprocess.Popen(f'explorer "{folder_path}"')
            except Exception as e:
                sg.popup_error(f"No se pudo abrir la carpeta:\n\n{e}")
        else:
            sg.popup_error("La carpeta no existe.")   
    
    while True:
        event,values = window.read()
        if event == 'Salir' or event == sg.WIN_CLOSED:
            break
        if event == '-SBOM-':
            ruta_bom = values['-BOM-']
        if event == '-FLEXA-':
            ruta_flexa = values['-FLEXA-'] 
            
        if event == 'Comparar':
            try:
                flexa_vs_bom()
                reset()
                csv_folder = r"C:\Users\CECHEVARRIAMENDOZA\OneDrive - Brunswick Corporation\Documents\Proyectos_Python\PysimpleGUI\Proyectos\comparador\csv"
                # Abre el explorador de archivos en la ruta específica
                open_folder_in_explorer(csv_folder)
                
            except:
                sg.popup_error('No se pudo realizar la comparacion\nIntente de nuevo')
                  
    window.close()


if __name__ == '__main__':
    main()