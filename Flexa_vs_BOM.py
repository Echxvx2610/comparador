import PySimpleGUI as sg
from bom_check import comparador
#from comparador_pruebas import comparador
import os
import subprocess
import logger

def main():
    #mostrar todos los temas
    #sg.theme_previewer()
    #sg.theme('LightGreen5')
    sg.theme("DefaultNoMoreNagging")
    layout = [
        [sg.Image(r'comparador\img\LOGO_NAVICO_1_90-black.png',expand_x=False,expand_y=False,enable_events=True,key='-LOGO-'),sg.Push()],
        [sg.Input(default_text='Ruta archivo Syteline',key='-BOM-',enable_events=True,size=(65,10),readonly=True,justification='center',font=('Arial',10,'italic')),sg.FileBrowse(file_types=(("Excel files", "*.xlsx"), ("All files", "*.*")),button_text="Cargar BOM",)],
        [sg.Input(default_text='Ruta archivo Placement Flexa ',key='-FLEXA-',enable_events=True,size=(60,10),readonly=True,justification='center',font=('Arial',10,'italic')),sg.FileBrowse(file_types=(("Excel files", "*.xlsx"), ("All files", "*.*")),button_text="Cargar Placement",)],
        [sg.Button("Abrir y Editar",key='-OPEN-'),sg.Button('Comparar'),sg.Button('Salir')],
        [sg.Text("Created by: Cristian Echevarría",font=('Arial',6,'italic'))],        
    ]
    
     
    window = sg.Window('......::: Flexa vs BOM :::......',
                       layout,finalize=True,
                       no_titlebar=False,
                       element_justification='center',
                       icon='comparador\img\document.ico',
                       keep_on_top=False,
                       resizable=False
                       )
    def flexa_vs_bom():
        comparador(values['-BOM-'],values['-FLEXA-'])
        sg.popup('Comparacion completada con exito!')
    
    def reset():
        window['-BOM-'].update('Ruta archivo Syteline')
        window['-FLEXA-'].update('Ruta archivo Placement Flexa ')
    
    def open_excel_and_get_path(file_path):
        try:
            # Abrir el archivo con Excel
            os.startfile(file_path)
            
            # Esperar a que el usuario cierre Excel y luego obtener la ruta de acceso del archivo guardado
            sg.popup_quick_message("Por favor, edita el archivo en Excel y guárdalo.", auto_close_duration=5)
            
            return file_path
        except Exception as e:
            sg.popup_error(f"Error: {str(e)}")
            return None
        
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
        
        if event == 'Comparar':            
             try:
                flexa_vs_bom()
                reset()
                csv_folder = r"H:\Ingenieria\SMT\Flexa_vs_BOM"
                # Abre el explorador de archivos en la ruta específica
                open_folder_in_explorer(csv_folder)
            #  except Exception as e:
            #     sg.popup_error(f"Error: {str(e)}")
             except:
                sg.popup('No se pudo realizar la comparacion,\nIntentelo de nuevo')
        if event == '-OPEN-':
            file_path = values['-FLEXA-']
            if file_path:
                edited_file_path = open_excel_and_get_path(file_path)
                if edited_file_path:
                    window["-FLEXA-"].update(edited_file_path)
                    
    window.close()


if __name__ == '__main__':
    main() 