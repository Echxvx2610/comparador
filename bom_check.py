import pandas as pd
import openpyxl
from openpyxl import workbook,load_workbook
import csv
import PySimpleGUI as sg
import os
import logger
import functools as ft

#configuracion de logger
#logger = logger.setup_logger(r'H:\Ingenieria\SMT\Flexa_vs_BOM\comp.log')
logger = logger.setup_logger(r'comparador\comp.log')

#......................:::: CONFIGURACION DEL DATAFRAME ::: :..................
pd.set_option('display.max_columns', None)
pd.set_option('display.max_rows', None)
pd.set_option('display.width', None)
#pd.set_option('expand_frame_repr', False)

#...........................:::: Variales Globales ::::....................
skipeados = ""
data_to_display = ""

def comparador(ruta_bom,ruta_flexa):    
    #************************************************************** SYTELINE ******************************************************************
    #Carga y conversion de excel syteline a dataframe
    #nombre_excel = r'C:\Users\CECHEVARRIAMENDOZA\OneDrive - Brunswick Corporation\Documents\Proyectos_Python\PysimpleGUI\Proyectos\comparador\SV9_007.xlsx'
    #print(len(nombre_excel)) #133 caracteres(nombre de archivo [125:])
    syteline = pd.read_excel(ruta_bom, engine='openpyxl')                                               # leemos el archivo excel de syteline 
    bom = pd.DataFrame(syteline)                                                                        # convertimos a dataframe
    bom.rename(columns={'Designators ':'Reference'},inplace=True)                                       # Renombramos columna(Designators a Reference) para que conicida con Placement
    bom.rename(columns={'Item':'Part Number'},inplace=True)                                             # Renombramos columna(Item a Part Number) para que conicida con Placement
    bom = bom[['Level','Operation','Part Number','Description','Reference']]                                    # seleccionamos las columnas deseadas(Operation,Part Number,Description,Reference)
    bom_op20 = bom[bom['Operation']==20.0]                                                              # creamos un dataframe filtrando por operacion 20
    bom_op10 = bom[bom['Operation']==10.0]                                                              # creamos un dataframe filtrando por operacion 10
    bom_filter = bom_op20.merge(bom_op10,how='outer')                                                   # creamos un dataframe combinando con un join(outer)
    #bom_filter.to_csv(r'C:\Users\CECHEVARRIAMENDOZA\OneDrive - Brunswick Corporation\Documents\Proyectos_Python\PysimpleGUI\Proyectos\comparador\csv\bom.csv',index=False)  # guardamos el dataframe
    bom_filter['Reference'] = bom_filter['Reference'].str.split()                                       # de la columna reference desglamos los elementos en elementos unicos es decir
    bom_filter = bom_filter.explode('Reference')                                                        # desempaquetamos la lista de referencias
    bom_filter.reset_index(drop=True,inplace=True)                                                      # reseta el indice,debido que al desempaquetar agregamos mas elemenetos al dataframe
    
    # Si hay elementos con un valor 2 en el level se descarta el elemento
    bom_filter = bom_filter[bom_filter['Level'] != 2]
    bom_filter = bom_filter[['Operation','Part Number','Description','Reference']]
    #print(bom_filter)
    #bom_filter.to_csv(r'C:\Users\CECHEVARRIAMENDOZA\OneDrive - Brunswick Corporation\Documents\Proyectos_Python\PysimpleGUI\Proyectos\comparador\csv\bom_filter.csv',index=False)

    
    #******************************************************************* PLACEMENT ******************************************************************
    #Carga y conversion de placement flexa a dataframe
    flexa = pd.read_excel(ruta_flexa, engine='openpyxl')                                                # leemos el archivo excel de flexa
    
    # implementar funcion validar_panel
    # if validar_panel(flexa) == True:
    #     sg.popup("Placement validado")
    #     pass
    # else:
    #     respuesta = sg.popup_yes_no("Desea continuar?")
    #     if respuesta == "No":
    #         return
        
    placement = pd.DataFrame(flexa)                                                                     # convertimos a dataframe
    placement.rename(columns={'Ref.':'Reference'},inplace=True)                                         # Renombramos columna(Ref. a Reference)
    logger.info(f"Comienza la comparacion con el archivo {ruta_flexa} vs {ruta_bom}")
    placement = placement[['Board','Part Number','Reference','Skip','Assign']]                          # Seleccionamos las columnas deseadas
    
    # validar si hay componentes con skip
    if "Yes" in placement['Skip'].values:                                                               # Si se encuentra skip en el archivo
        #print(placement['Skip'].values == "Yes")                                                       # alertamos al usuario 
        logger.info(f'Se reviso archivo placement {ruta_flexa} y se encontraron componentes con skip')
        title = "! Alerta !"
        message = """Se encontraron componentes con skip en el archivo!"""
        sg.popup(message, title=title)
        skipeados = placement[placement['Skip']=='Yes']
        # mostrar el logger solo no.partes y referencia
        logger.info(f"Se encontraron {len(skipeados)} componentes con Skip, no.parte {skipeados['Part Number'].values} y referencia {skipeados['Reference'].values}")
        skipeados = skipeados[['Board','Part Number','Reference']]
        data_to_display = skipeados.values.tolist()                                                    # convertimos a lista todos los elementos con skip
        table(data_to_display,skipeados)                                                               # creamos la tabla y desplegamos la tabla
        respuesta = sg.popup_yes_no("Desea continuar?",title=title)
        if respuesta == "Yes":
            logger.info(f"Se decidio continuar con la comparacion del archivo {ruta_flexa}")
            #adaptar comparacion + componentes skipeados
            pass
        else:                                                                                
            logger.info(f"No se realizo la comparacion del archivo {ruta_flexa}")
            # salimos del programa si no quiere continuar
            return 
    
    
    # validar si hay componentes sin asignar ( valores nan)
    if placement['Assign'].isna().any():
        # Si se encuentra vacio el assign para un no.part en el archivo
        logger.info(f'Se reviso archivo placement {ruta_flexa} y se encontraron componentes sin asignar')
        title = "! Alerta !"
        message = """Se encontraron componentes sin asignar en el archivo!"""
        sg.popup(message, title=title)
        sin_asignar = placement[placement['Assign'].isna()]
        logger.info(f"Se encontraron {len(sin_asignar)} componentes sin asignar, no.parte {sin_asignar['Part Number'].values} y referencia {sin_asignar['Reference'].values}")
        sin_asignar = sin_asignar[['Board','Part Number','Reference']]
        data_to_display = sin_asignar.values.tolist()                                                   # convertimos a lista todos los elementos sin asignar
        table(data_to_display,sin_asignar)                                                             # creamos la tabla y desplegamos la tabla
        respuesta = sg.popup_yes_no("Desea continuar?",title=title)
        if respuesta == "Yes":
            logger.info(f"Se decidio continuar con la comparacion del archivo {ruta_flexa}")
            #adaptar comparacion + componentes sin asignar
            pass
        else:
            logger.info(f"No se realizo la comparacion del archivo {ruta_flexa}")
            # salimos del programa si no quiere continuar
            return 
    
    #******************************************************************* COMPARACION ******************************************************************
    comparacion = bom_filter.merge(placement, on = ['Part Number','Reference'], how='outer',suffixes=('_izq', '_der'), indicator=True)          # Juntamos los dataframes
    comparacion.rename(columns={'_merge':'Comparacion'},inplace=True)                                                                           # renombramos la columna merge por Comparacion
    comparacion['Comparacion'] = comparacion['Comparacion'].replace({                                                                           # personalizamos la columna comparacion
    'left_only': 'Solo en BOM',
    'right_only': 'Solo en Placement',
    'both': 'En ambos archivos'
    })
    
    only_bom = comparacion[comparacion['Comparacion'] == 'left_only']
    only_placement = comparacion[comparacion['Comparacion'] == 'right_only']
    # comparacion final sera las diferencias de ambos archivos y ademas los componentes que lleven yes en skip
    comparacion_final = comparacion[(comparacion['Comparacion'] != 'En ambos archivos') | ((comparacion['Skip'] == 'Yes') | (comparacion['Skip'].isna()) | (comparacion['Assign'] == "") | (comparacion['Assign'].isna()))]
    
    # retiramos la fila completa si se encuentra un No.Parte en la columna No.Parte
    comparacion_final = comparacion_final[~comparacion_final['Part Number'].str.startswith('017-')]
    #comparacion_final = comparacion_final[~comparacion_final['Part Number'].str.startswith('014-')]  // esto es un numero de parte!!
    comparacion_final = comparacion_final[~comparacion_final['Part Number'].str.startswith('051-')]
    comparacion_final = comparacion_final[~comparacion_final['Part Number'].str.startswith('140-')]
    comparacion_final = comparacion_final[~comparacion_final['Part Number'].str.startswith('124-')]
    
    
    print("comparacion final es vacia: ",comparacion_final.empty) #hasta ese punto si no hay diferencias el dataframe es vacio
    print("comparacion final: ",comparacion_final)  
    
    if comparacion_final.empty:
        logger.info(f"No se encontraron diferencias en comparacion con el archivo {ruta_bom} y {ruta_flexa}")
        sg.popup('No se encontraron diferencias :)')
        return None,None
    else:
        sg.popup('Se encontraron diferencias :O')
        logger.info(f"Se encontraron diferencias entre {ruta_flexa} y {ruta_bom}")
        # creamos la carpeta y el csv de la comparacion
        nombre_excel_sin_extension = os.path.splitext(os.path.basename(ruta_flexa))[0]                                                             # creamos el nombre del archivo sin extension
        logger.info(f'Se realizo la comparacion entre {ruta_flexa} y {ruta_bom}')
        carpeta_nombre_archivo = r"H:\Ingenieria\SMT\Flexa_vs_BOM\{nombre_excel_sin_extension}".format(nombre_excel_sin_extension=nombre_excel_sin_extension) # creamos la carpeta donde se guardara el archivo
        os.makedirs(carpeta_nombre_archivo, exist_ok=True)
        ruta_csv = os.path.join(carpeta_nombre_archivo,f"{nombre_excel_sin_extension}.csv")
        # creamos un dataframe que contenga las diferencias y lo guardamos en CSV
        #funcional --> comparacion_final = comparacion[comparacion['Comparacion'] != 'En ambos archivos']
        
        # retiramos la fila completa si se encuentra un No.Parte en la columna No.Parte
        comparacion_final = comparacion_final[~comparacion_final['Part Number'].str.startswith('017-')]
        #comparacion_final = comparacion_final[~comparacion_final['Part Number'].str.startswith('014-')]  // esto es un numero de parte!!
        comparacion_final = comparacion_final[~comparacion_final['Part Number'].str.startswith('051-')]
        comparacion_final = comparacion_final[~comparacion_final['Part Number'].str.startswith('140-')]
        comparacion_final = comparacion_final[['Operation','Board','Reference','Part Number','Skip','Assign','Description','Comparacion']]
        # generamos el CSV
        comparacion_final.to_csv(ruta_csv,index=False)
        logger.info(f'Se genero el CSV {ruta_csv} con las diferencias de la comparacion')
        logger.info("--------------------------------------------------------------\n")
        return True,ruta_csv
        
def comparacion_nexim(ruta_bom,ruta_nexim):
    #************************************************************** SYTELINE ******************************************************************
    syteline = pd.read_excel(ruta_bom, engine='openpyxl')                                              
    bom = pd.DataFrame(syteline)                                                                       
    bom.rename(columns={'Designators ':'Reference'},inplace=True)                                       
    bom.rename(columns={'Item':'Part Number'},inplace=True)                                            
    bom = bom[['Level','Operation','Part Number','Description','Reference']]                                    
    bom_op20 = bom[bom['Operation']==20.0]                                                             
    bom_op10 = bom[bom['Operation']==10.0]                                                             
    bom_filter = bom_op20.merge(bom_op10,how='outer')                                                  
    #bom_filter.to_csv(r'C:\Users\CECHEVARRIAMENDOZA\OneDrive - Brunswick Corporation\Documents\Proyectos_Python\PysimpleGUI\Proyectos\comparador\csv\bom.csv',index=False)  
    bom_filter['Reference'] = bom_filter['Reference'].str.split()                                      
    bom_filter = bom_filter.explode('Reference')                                                      
    bom_filter.reset_index(drop=True,inplace=True)
                                                        
    # Si hay elementos con un valor 2 en el level se descarta el elemento
    bom_filter = bom_filter[bom_filter['Level'] != 2]
    bom_filter = bom_filter[['Operation','Part Number','Description','Reference']]
    #print(bom_filter)
    #bom_filter.to_csv(r'comparador\bom_filter_nexim.csv',index=False)
    
    
    #******************************************************************* PLACEMENT ******************************************************************
    #Carga y conversion de placement nexim a dataframe
    nexim = pd.read_excel(ruta_nexim, engine='openpyxl')                                              
    placement = pd.DataFrame(nexim)                                                                    
    placement.rename(columns={'Ref.':'Reference'},inplace=True)
    placement = placement[['Board','Part Number','Reference','Skip','Assign']]
    logger.info(f"Comienza la comparacion con el archivo {ruta_nexim} vs {ruta_bom}")
    
    # Validar que no existan componentes con Skip
    placement = placement[~placement['Part Number'].str.startswith("NOT")]
    
    if "Yes" in placement['Skip'].values:
        title = "! Alerta !"
        message = """Se encontraron componentes con skip en el archivo!"""
        sg.popup(message,title = title)
        skipeados = placement[placement['Skip']=='Yes']
        logger.info(f"Se encontraron {len(skipeados)} componentes con Skip, no.parte {skipeados['Part Number'].values} y referencia {skipeados['Reference'].values}")
        skipeados = skipeados[['Board','Part Number','Reference']]
        data_to_display = skipeados.values.tolist()
        table(data_to_display,skipeados)
        respuesta = sg.popup_yes_no("¿Desea continuar?",title = title)
        if respuesta == 'Yes':
            logger.info(f"Se decidio continuar con la comparacion del archivo {ruta_nexim}")
            # placement = placement
            pass
        else:
            logger.info(f"No se realizo la comparacion del archivo {ruta_nexim}")
            return
    
    # Validar si hay comoponentes sin asignar
    if placement['Assign'].isna().any():
        logger.info(f"Se reviso archivo placement {ruta_nexim} y se encontraron componentes sin asignar")
        title = "! Alerta !"
        message = """Se encontraron componentes sin asignar en el archivo!"""
        sg.popup(message,title = title)
        sin_asignar = placement[placement['Assign'].isna()]
        logger.info(f"Se encontraron {len(sin_asignar)} componentes sin asignar, no.parte {sin_asignar['Part Number'].values} y referencia {sin_asignar['Reference'].values}")
        sin_asignar = sin_asignar[['Board','Part Number','Reference']]
        data_to_display = sin_asignar.values.tolist()
        table(data_to_display,sin_asignar)
        respuesta = sg.popup_yes_no("¿Desea continuar?",title = title)
        if respuesta == 'Yes':
            logger.info(f"Se decidio continuar con la comparacion del archivo {ruta_nexim}")
            # placement = placement
            pass
        else:
            logger.info(f"No se realizo la comparacion del archivo {ruta_nexim}")
            return
    
    
    #****************************************************************** COMPARACION ******************************************************************
    # Unimos ambos archivos y creamos el csv de comparacion
    comparacion = bom_filter.merge(placement,how='outer',on=['Part Number','Reference'],suffixes=('_bom','_placement'),indicator=True)
    #print(comparacion[comparacion['_merge']!='both'])
    comparacion.rename(columns={'_merge':'Comparacion'},inplace=True)
    comparacion["Comparacion"] = comparacion["Comparacion"].replace({
        "left_only": "Solo en Bom",
        "right_only": "Solo en Nexim",
        "both": "En ambos archivos"
    })
    # comparacion final sera las diferencias de ambos archivos y ademas los componentes que lleven yes en skip
    comparacion_final = comparacion[(comparacion['Comparacion'] != 'En ambos archivos') | ((comparacion['Skip'] == 'Yes') | (comparacion['Skip'].isna()) | (comparacion['Assign']=="") | (comparacion['Assign'].isna()))]
    #comparacion_final = comparacion[comparacion['Comparacion']!='En ambos archivos']
    
    # retiramos la fila completa si se encuentra un No.Parte en la columna No.Parte
    comparacion_final = comparacion_final[~comparacion_final['Part Number'].str.startswith('017-')]
    #comparacion_final = comparacion_final[~comparacion_final['Part Number'].str.startswith('014-')]  // esto es un numero de parte!!
    comparacion_final = comparacion_final[~comparacion_final['Part Number'].str.startswith('051-')]
    comparacion_final = comparacion_final[~comparacion_final['Part Number'].str.startswith('140-')] 
    
    print("Comparacion final es vacia: ",comparacion_final.empty)
    print("Comparacion final: ",comparacion_final)
    
    if comparacion_final.empty:
        logger.info(f"No se encontraron diferencias en comparacion con el archivo {ruta_bom} y {ruta_nexim}")
        sg.popup('No se encontraron diferencias :)')
        return None,None
    else:
        logger.info(f"Se encontraron diferencias entre {ruta_bom} y {ruta_nexim}")
        sg.popup('Se encontraron diferencias :O')
        # creamos la carpeta y el csv de la comparacion
        nombre_excel_sin_extension = os.path.splitext(os.path.basename(ruta_nexim))[0]
        logger.info(f"Se realizo la comparacion entre {ruta_nexim} y {ruta_bom}")
        carpeta_nombre_archivo = r"H:\Ingenieria\SMT\Flexa_vs_BOM\Nexim\{nombre_excel_sin_extension}".format(nombre_excel_sin_extension=nombre_excel_sin_extension)
        os.makedirs(carpeta_nombre_archivo, exist_ok=True)
        ruta_csv = os.path.join(carpeta_nombre_archivo,f'{nombre_excel_sin_extension}.csv')
        
        # retiramos la fila completa si se encuentra un No.Parte en la columna No.Parte
        comparacion_final = comparacion_final[~comparacion_final['Part Number'].str.startswith('017-')]
        #comparacion_final = comparacion_final[~comparacion_final['Part Number'].str.startswith('014-')]  // esto es un numero de parte!!
        comparacion_final = comparacion_final[~comparacion_final['Part Number'].str.startswith('051-')]
        comparacion_final = comparacion_final[~comparacion_final['Part Number'].str.startswith('140-')]
        comparacion_final = comparacion_final[['Operation','Board','Reference','Part Number','Skip','Assign','Description','Comparacion']]          
        # guardamos el csv
        comparacion_final.to_csv(ruta_csv,index=False)
        logger.info(f"Se genero el CSV {ruta_csv} con las diferencias entre {ruta_bom} y {ruta_nexim}")
        logger.info("----------------------------------------------------------------------------------")
        return True,ruta_csv
           
#Comparacion entre bom y placement 
def comparacion_bom(ruta_bom,ruta_bom2):
    #************************************************************** BOM 1 ******************************************************************
    #Carga y conversion de excel syteline a dataframe
    syteline = pd.read_excel(ruta_bom, engine='openpyxl')                                               # leemos el archivo excel de syteline 
    bom = pd.DataFrame(syteline)                                                                        # convertimos a dataframe
    bom.rename(columns={'Designators ':'Reference'},inplace=True)                                       # Renombramos columna(Designators a Reference) para que conicida con Placement
    bom.rename(columns={'Item':'Part Number'},inplace=True)                                             # Renombramos columna(Item a Part Number) para que conicida con Placement
    bom = bom[['Operation','Part Number','Description','Reference']]                                    # seleccionamos las columnas deseadas(Operation,Part Number,Description,Reference)
    bom_op20 = bom[bom['Operation']==20.0]                                                              # creamos un dataframe filtrando por operacion 20
    bom_op10 = bom[bom['Operation']==10.0]                                                              # creamos un dataframe filtrando por operacion 10
    bom_filter = bom_op20.merge(bom_op10,how='outer')                                                   # creamos un dataframe combinando con un join(outer)
    #bom_filter.to_csv(r'C:\Users\CECHEVARRIAMENDOZA\OneDrive - Brunswick Corporation\Documents\Proyectos_Python\PysimpleGUI\Proyectos\comparador\csv\bom.csv',index=False)  # guardamos el dataframe
    bom_filter['Reference'] = bom_filter['Reference'].str.split()                                       # de la columna reference desglamos los elementos en elementos unicos es decir
    bom_filter = bom_filter.explode('Reference')                                                        # desempaquetamos la lista de referencias
    bom_filter.reset_index(drop=True,inplace=True)                                                      # reseta el indice,debido que al desempaquetar agregamos mas elemenetos al dataframe
    #bom_filter.to_csv(r'comparador\csv\bom_filter.csv',index=False)
    #************************************************************** BOM 2 ******************************************************************
    syteline2 = pd.read_excel(ruta_bom2, engine='openpyxl')                                               # leemos el archivo excel de syteline
    bom2 = pd.DataFrame(syteline2)                                                                       # convertimos a dataframe
    bom2.rename(columns={'Designators ':'Reference'},inplace=True)                                      # Renombramos columna(Designators a Reference) para que conicida con Placement
    bom2.rename(columns={'Item':'Part Number'},inplace=True)                                            # Renombramos columna(Item a Part Number) para que conicida con Placement
    bom2 = bom2[['Operation','Part Number','Description','Reference']]                                   # seleccionamos las columnas deseadas(Operation,Part Number,Description,Reference)
    bom2_op20 = bom2[bom2['Operation']==20.0]                                                           # creamos un dataframe filtrando por operacion 20
    bom2_op10 = bom2[bom2['Operation']==10.0]                                                           # creamos un dataframe filtrando por operacion 10
    bom2_filter = bom2_op20.merge(bom2_op10,how='outer')                                                # creamos un dataframe combinando con un join(outer)
    bom2_filter['Reference'] = bom2_filter['Reference'].str.split()                                    # de la columna reference desglamos los elementos en elementos unicos es decir
    bom2_filter = bom2_filter.explode('Reference')                                                     # desempaquetamos la lista de referencias
    bom2_filter.reset_index(drop=True,inplace=True)                                                     # reseta el indice,debido que al desempaquetar agregamos mas elemenetos al dataframe
    #bom2_filter.to_csv(r'comparador\csv\bom2_filter.csv',index=False)
    # ******************************************************************* COMPARACION ******************************************************************
    logger.info(f"Comienza la comparacion con el archivo {ruta_bom} vs {ruta_bom2}")
    comparacion = bom_filter.merge(bom2_filter,how='outer',suffixes = ('_izq', '_der'),indicator=True)
    comparacion.rename(columns={'_merge':'Comparacion'},inplace=True)
    comparacion['Comparacion'] = comparacion['Comparacion'].replace({
        'left_only': 'BOM_izq',
        'right_only': 'BOM_der',
        'both': 'En ambos archivos'
    })
    bom_izq = comparacion[comparacion['Comparacion']=='left_only']
    bom_der = comparacion[comparacion['Comparacion']=='right_only']
    
    comparacion_final = comparacion[comparacion['Comparacion']!='En ambos archivos']
    # si no hay diferencias solo alerta un Pop up completado con exito!, si hay diferencias crea el archivo csv
    if comparacion_final.empty:
        sg.popup('No hay diferencias entre los BOM :)')
        logger.info(f'No se encontraron diferencias entre los BOM {ruta_bom} y {ruta_bom2}')
        return None,None
    else:
        logger.info(f'Se encontraron diferencias entre los BOM {ruta_bom} y {ruta_bom2}')
        sg.popup('Se han encontrado diferencias entre los BOM :O')
        ## Comparacion final sera un dataframe que contenga los datos que sean diferentes en ambos archivos pero no es necesario mostrar los no.part que contenga NOT IN BOM
        nombre_excel_sin_extension = os.path.splitext(os.path.basename(ruta_bom2))[0]
        carpeta_nombre_archivo = r"H:\Ingenieria\SMT\Flexa_vs_BOM\BOM\{nombre_excel_sin_extension}".format(nombre_excel_sin_extension=nombre_excel_sin_extension)
        os.makedirs(carpeta_nombre_archivo, exist_ok=True)
        ruta_csv = os.path.join(carpeta_nombre_archivo,f"{nombre_excel_sin_extension}.csv")
        comparacion_final = comparacion[comparacion['Comparacion'] != 'En ambos archivos']
        comparacion_final.to_csv(ruta_csv,index=False)
        logger.info(f'Se realizo la comparacion entre los BOM y se genero el CSV {ruta_csv}')
        logger.info('--------------------------------------------------------------\n')   
        return True, ruta_csv
    
    
def table(data_to_display,skipeados):
    # Creamos el diseño de la tabla utilizando PySimpleGUI
    layout = [[sg.Table(values=data_to_display,
                        headings=skipeados.columns.tolist(),
                        display_row_numbers=False,
                        justification='center',
                        auto_size_columns=True,
                        num_rows=min(25, len(data_to_display)))],
              [sg.Button("Cerrar")]]

    # Creamos la ventana del popup
    window = sg.Window("Componentes con Skip", layout)

    # Mostramos el popup y esperamos a que el usuario lo cierre
    while True:
        event, values = window.read()
        if event == sg.WIN_CLOSED or event == "Cerrar":
            break

    # Cerramos la ventana
    window.close()