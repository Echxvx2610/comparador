import pandas as pd
import openpyxl
from openpyxl import workbook,load_workbook
import csv
import PySimpleGUI as sg


#......................:::: CONFIGURACION DEL DATAFRAME ::::..................
pd.set_option('display.max_columns', None)
pd.set_option('display.max_rows', None)
pd.set_option('display.width', None)
#pd.set_option('expand_frame_repr', False)

#...........................:::: variales globales ::::....................
skipeados = ""
data_to_display = ""

def comparador(ruta_bom,ruta_flexa):    
    #************************************************************** SYTELINE ******************************************************************
    #Carga y conversion de excel syteline a dataframe
    nombre_excel = r'C:\Users\CECHEVARRIAMENDOZA\OneDrive - Brunswick Corporation\Documents\Proyectos_Python\PysimpleGUI\Proyectos\comparador\SV9_007.xlsx'
    #print(len(nombre_excel)) #133 caracteres(nombre de archivo [125:])
    syteline = pd.read_excel(ruta_bom, engine='openpyxl')                                               # leemos el archivo excel de syteline 
    bom = pd.DataFrame(syteline)                                                                        # convertimos a dataframe
    bom.rename(columns={'Designators ':'Reference'},inplace=True)                                       # Renombramos columna(Designators a Reference) para que conicida con Placement
    bom.rename(columns={'Item':'Part Number'},inplace=True)                                             # Renombramos columna(Item a Part Number) para que conicida con Placement
    bom = bom[['Operation','Part Number','Description','Reference']]                                    # seleccionamos las columnas deseadas(Operation,Part Number,Description,Reference)
    bom_op20 = bom[bom['Operation']==20.0]                                                              # creamos un dataframe filtrando por operacion 20
    bom_op10 = bom[bom['Operation']==10.0]                                                              # creamos un dataframe filtrando por operacion 10
    bom_filter = bom_op20.merge(bom_op10,how='outer')                                                   # creamos un dataframe combinando con un join(outer)
    nombre_sin_extencion = os.path.splitext(nombre_excel)[0]
    print(nombre_sin_extencion)
    bom_filter.to_csv(ruta_csv,index=False)  # guardamos el dataframe
    bom_filter['Reference'] = bom_filter['Reference'].str.split()                                       # de la columna reference desglamos los elementos en elementos unicos es decir
    bom_filter = bom_filter.explode('Reference')                                                        # desempaquetamos la lista de referencias
    bom_filter.reset_index(drop=True,inplace=True)                                                      # reseta el indice,debido que al desempaquetar agregamos mas elemenetos al dataframe
    bom_filter.to_csv(r'C:\Users\CECHEVARRIAMENDOZA\OneDrive - Brunswick Corporation\Documents\Proyectos_Python\PysimpleGUI\Proyectos\comparador\csv\bom_filter.csv',index=False)

    #******************************************************************* PLACEMENT ******************************************************************
    #carga y conversion de placement flexa a dataframe
    nombre_placement = r'C:\Users\CECHEVARRIAMENDOZA\OneDrive - Brunswick Corporation\Documents\Proyectos_Python\PysimpleGUI\Proyectos\comparador\29736-1B.xlsx'
    flexa = pd.read_excel(ruta_flexa, engine='openpyxl')                                                # leemos el archivo excel de flexa
    placement = pd.DataFrame(flexa)                                                                     # convertimos a dataframe
    placement.rename(columns={'Ref.':'Reference'},inplace=True)                                         # Renombramos columna(Ref. a Reference)                                          # Renombramos columna(Board a Board)
    placement = placement[['Board','Part Number','Reference','Skip']]
    if "Yes" in placement['Skip'].values:
        #print(placement['Skip'].values == "Yes") 
        title = "! Alerta !"
        message = """Se encontraron componentes skipeados en el archivo"""
        sg.popup(message, title=title)
        skipeados = placement[placement['Skip']=='Yes']
        print(skipeados)
        print(skipeados.columns)
        data_to_display = skipeados.values.tolist()
        table(data_to_display,skipeados)
        respuesta = sg.popup_yes_no("Desea continuar?",title=title)
        if respuesta == "Yes":
            pass
        else:
            exit()
    placement.to_csv(r'C:\Users\CECHEVARRIAMENDOZA\OneDrive - Brunswick Corporation\Documents\Proyectos_Python\PysimpleGUI\Proyectos\comparador\csv\placement.csv',index=False)
    #******************************************************************* COMPARACION ******************************************************************
    comparacion = bom_filter.merge(placement, on = ['Part Number','Reference'], how='outer',suffixes=('_izq', '_der'), indicator=True)
    comparacion.rename(columns={'_merge':'Comparacion'},inplace=True)
    comparacion['Comparacion'] = comparacion['Comparacion'].replace({
    'left_only': 'Solo en BOM',
    'right_only': 'Solo en Placement',
    'both': 'En ambos archivos'
    })
    only_bom = comparacion[comparacion['Comparacion'] == 'left_only']
    only_placement = comparacion[comparacion['Comparacion'] == 'right_only']
    comparacion.to_csv(r'C:\Users\CECHEVARRIAMENDOZA\OneDrive - Brunswick Corporation\Documents\Proyectos_Python\PysimpleGUI\Proyectos\comparador\csv\comparacion.csv',index=False)
    '''
    #Eliminamos filas duplicadas en el segundo DataFrame
    df2_unique_refs = df2.drop_duplicates(subset='Ref.', keep='first')

    #Comparamos referencias únicas del segundo DataFrame con el primer DataFrame
    references_in_df1 = df1['Reference'].isin(df2_unique_refs['Ref.'])

    #Identificamos si una referencia única del segundo DataFrame está en el primer DataFrame
    df2_unique_refs['Is_Present_In_df1'] = references_in_df1.tolist()

    print(df2_unique_refs)
    '''
    
    
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
    window = sg.Window("Filas con 'Yes' en la columna 'Skip'", layout)

    # Mostramos el popup y esperamos a que el usuario lo cierre
    while True:
        event, values = window.read()
        if event == sg.WIN_CLOSED or event == "Cerrar":
            break

    # Cerramos la ventana
    window.close()
    
    
import os
import pandas as pd

def guardar_en_carpeta(nombre_archivo_excel):
    # Leer el archivo Excel y generar el DataFrame filtrado
    # (Coloca aquí el código para leer y filtrar el Excel)
    # dataframe_filtrado = ...

    # Extraer el nombre del archivo sin la extensión
    nombre_sin_extension = os.path.splitext(nombre_archivo_excel)[0]

    # Crear una carpeta con el nombre del archivo Excel
    if not os.path.exists(nombre_sin_extension):
        os.mkdir(nombre_sin_extension)

    # Ruta completa del archivo CSV dentro de la carpeta
    ruta_csv = os.path.join(nombre_sin_extension, f"{nombre_sin_extension}.csv")

    # Guardar el DataFrame filtrado como CSV
    dataframe_filtrado.to_csv(ruta_csv, index=False)