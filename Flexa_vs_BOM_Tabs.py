import PySimpleGUI as sg    
from bom_check import comparador,logger,comparacion_nexim,comparacion_bom
#from comparador_pruebas import comparador,logger,comparacion_nexim,comparacion_bom
import os
import subprocess

#Implementacion de un layout Tab para duplicar la aplicacion y comparar con BOM para Nexim
def main():
    sg.theme("DefaultNoMoreNagging")
    tab1_layout =  [
            [sg.Image(r'comparador\img\LOGO_NAVICO_1_90-black.png',expand_x=False,expand_y=False,enable_events=True,key='-LOGO-'),sg.Push()],
            [sg.Input(default_text='Ruta archivo Syteline',key='-BOM-',enable_events=True,size=(65,10),readonly=True,justification='center',font=('Arial',10,'italic')),sg.FileBrowse(file_types=(("Excel files", "*.xlsx"), ("All files", "*.*")),button_text="Cargar BOM",)],
            [sg.Input(default_text='Ruta archivo Placement Flexa ',key='-FLEXA-',enable_events=True,size=(60,10),readonly=True,justification='center',font=('Arial',10,'italic')),sg.FileBrowse(file_types=(("Excel files", "*.xlsx"), ("All files", "*.*")),button_text="Cargar Placement",)],
            [sg.Button("Abrir y Editar",key='-OPEN-'),sg.Button('Comparar',key='-COMPARE-'),sg.Button('Salir',key='-SALIR-')],
            [sg.Text("Created by: Cristian Echevarría",font=('Arial',7,'italic'))],        
        ]

    tab2_layout = [
            [sg.Image(r'comparador\img\LOGO_NAVICO_1_90-black.png',expand_x=False,expand_y=False,enable_events=True,key='-LOGO2-'),sg.Push()],
            [sg.Input(default_text='Ruta archivo Syteline',key='-BOM2-',enable_events=True,size=(65,10),readonly=True,justification='center',font=('Arial',10,'italic')),sg.FileBrowse(file_types=(("Excel files", "*.xlsx"), ("All files", "*.*")),button_text="Cargar BOM",)],
            [sg.Input(default_text='Ruta archivo Placement Nexim ',key='-NEXIM-',enable_events=True,size=(60,10),readonly=True,justification='center',font=('Arial',10,'italic')),sg.FileBrowse(file_types=(("Excel files", "*.xlsx"), ("All files", "*.*")),button_text="Cargar Placement",)],
            [sg.Button("Abrir y Editar",key='-OPEN2-'),sg.Button('Comparar',key='-COMPARE2-'),sg.Button('Salir',key='-SALIR2-')],
            [sg.Text("Created by: Cristian Echevarría",font=('Arial',7,'italic'))],        
        ]
    tab3_layout = [
            [sg.Image(r'comparador\img\LOGO_NAVICO_1_90-black.png',expand_x=False,expand_y=False,enable_events=True,key='-LOGO2-'),sg.Push()],
            [sg.Input(default_text='Ruta archivo Syteline (BOM 1)',key='-BOM3-',enable_events=True,size=(65,10),readonly=True,justification='center',font=('Arial',10,'italic')),sg.FileBrowse(file_types=(("Excel files", "*.xlsx"), ("All files", "*.*")),button_text="Cargar BOM",)],
            [sg.Input(default_text='Ruta archivo Syteline (BOM 2) ',key='-BOM4-',enable_events=True,size=(65,10),readonly=True,justification='center',font=('Arial',10,'italic')),sg.FileBrowse(file_types=(("Excel files", "*.xlsx"), ("All files", "*.*")),button_text="Cargar BOM",)],
            [sg.Button('Comparar',key='-COMPARE3-'),sg.Button('Salir',key='-SALIR3-')],
            [sg.Text("Created by: Cristian Echevarría",font=('Arial',7,'italic'))],        
        ]   

    layout = [
        [sg.TabGroup([
            [sg.Tab("Flexa", tab1_layout,element_justification= 'center'),
            sg.Tab("Nexim", tab2_layout,element_justification= 'center'),
            sg.Tab("BOM", tab3_layout,element_justification= 'center')]],title_color='White',selected_background_color='red',tab_background_color='black',),
        ]]    
    
    window = sg.Window('..:: Flexa vs BOM ::..',
                        layout,finalize=True,
                        no_titlebar=False,
                        element_justification='center',
                        icon='comparador\img\document.ico',
                        keep_on_top=False,
                        resizable=False,
                        )
    
    #............................::::::: Comparacion Flexa vs BOM :::::::...............................
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
            sg.popup_quick_message("Por favor, edita el archivo en Excel y guárdalo.", auto_close_duration=3)
            
            return file_path
        except Exception as e:
            sg.popup_error(f"Error: {str(e)}")
            logger.error(str(e))
            return None
        
    def open_folder_in_explorer(folder_path):
        if os.path.exists(folder_path):
            try:
                subprocess.Popen(f'explorer "{folder_path}"')
            except Exception as e:
                sg.popup_error(f"No se pudo abrir la carpeta:\n\n{e}")
                logger.error(f"No se pudo abrir la carpeta:\n\n{e}")
        else:
            sg.popup_error("La carpeta no existe.")   
    
    while True:
        event,values = window.read()
        
        if event == '-SALIR-' or event == '-SALIR2-' or event == '-SALIR3-' or event == sg.WIN_CLOSED:
            break
        if event == '-COMPARE-':
            try:
                flexa_vs_bom()
                reset()
                csv_folder = r"H:\Ingenieria\SMT\Flexa_vs_BOM"
                # Abre el explorador de archivos en la ruta específica
                open_folder_in_explorer(csv_folder)
            except Exception as e:
                sg.popup('No se pudo realizar la comparacion,\nIntentelo de nuevo')
                logger.error(str(e))
                
        if event == '-OPEN-':
            file_path = values['-FLEXA-']
            if file_path:
                edited_file_path = open_excel_and_get_path(file_path)
                if edited_file_path:
                    window["-FLEXA-"].update(edited_file_path)
    
        #......................................:::::     Comparacion Nexim vs BOM ::::::...............................
        def nexim_vs_bom():
            comparacion_nexim(values['-BOM2-'],values['-NEXIM-'])
            sg.popup('Comparacion completada con exito!')
        
        def reset_nexim():
            window['-BOM2-'].update('Ruta archivo Syteline')
            window['-NEXIM-'].update('Ruta archivo Placement Nexim ')
            
        if event == '-COMPARE2-':
            try:
                nexim_vs_bom()
                reset_nexim()
                csv_folder_nexim = r"H:\Ingenieria\SMT\Flexa_vs_BOM\Nexim"
                # Abre el explorador de archivos en la ruta unica
                open_folder_in_explorer(csv_folder_nexim)
            except Exception as e:
                sg.popup('No se pudo realizar la comparacion,\nIntentelo de nuevo')
                logger.error(str(e))
        if event == '-OPEN2-':
            file_path = values['-NEXIM-']
            if file_path:
                edited_file_path = open_excel_and_get_path(file_path)
                if edited_file_path:
                    window["-NEXIM-"].update(edited_file_path)
        
         #............................::::::: Comparacion BOM vs BOM :::::::...............................
        def bom_vs_bom():
            comparacion_bom(values['-BOM3-'],values['-BOM4-'])
            sg.popup('Comparacion completada con exito!')
        
        def reset_bom():
            window['-BOM3-'].update('Ruta archivo Syteline (BOM Izq)')
            window['-BOM4-'].update('Ruta archivo Syteline (BOM Der)')
            
        if event == '-COMPARE3-':
            try:
                bom_vs_bom()
                reset_bom()
                csv_folder = r"H:\Ingenieria\SMT\Flexa_vs_BOM\BOM"
                # Abre el explorador de archivos en la ruta específica
                open_folder_in_explorer(csv_folder)
            except Exception as e:
                sg.popup('No se pudo realizar la comparacion,\nIntentelo de nuevo')
                logger.error(str(e))
    
                 
    window.close()
    
if __name__ == '__main__':
    main()