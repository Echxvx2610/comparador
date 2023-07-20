import pandas as pd
import numpy as np
import openpyxl
from openpyxl import workbook,load_workbook
#from openpyxl.utils import get_column_letter
#from openpyxl.styles import Font, Alignment, Border, Side
#from datetime import datetime,time
import asyncio
import csv


#......................:::: CONFIGURACION DEL DATAFRAME ::::..................
pd.set_option('display.max_columns', None)
pd.set_option('display.max_rows', None)
pd.set_option('display.width', None)
#pd.set_option('expand_frame_repr', False)


#************************************************************** SYTELINE ******************************************************************
#carga y conversion de excel syteline a dataframe
nombre_excel = r'C:\Users\CECHEVARRIAMENDOZA\OneDrive - Brunswick Corporation\Documents\Proyectos_Python\PysimpleGUI\Proyectos\comparador\excel\SV9_007.xlsx'
syteline = pd.read_excel(nombre_excel, engine='openpyxl')
#imprimir solo las columnas deseadas
bom = pd.DataFrame(syteline)
bom.rename(columns={'Designators ':'Reference'},inplace=True)
bom.rename(columns={'Item':'Part Number'},inplace=True)
bom = bom[['Operation','Part Number','Description','Reference']]
bom_op20 = bom[bom['Operation']==20.0]
bom_op10 = bom[bom['Operation']==10.0]
bom_filter = bom_op20.merge(bom_op10,how='outer')
bom_filter.to_csv(r'C:\Users\CECHEVARRIAMENDOZA\OneDrive - Brunswick Corporation\Documents\Proyectos_Python\PysimpleGUI\Proyectos\comparador\csv\bom.csv',index=False)
bom_filter['Reference'] = bom_filter['Reference'].str.split()
bom_filter = bom_filter.explode('Reference')
bom_filter.reset_index(drop=True,inplace=True)
bom_filter.to_csv(r'C:\Users\CECHEVARRIAMENDOZA\OneDrive - Brunswick Corporation\Documents\Proyectos_Python\PysimpleGUI\Proyectos\comparador\csv\bom_filter.csv',index=False)
#print(bom_filter)



#******************************************************************* PLACEMENT ******************************************************************
#carga y conversion de placement flexa a dataframe
nombre_placement = r'C:\Users\CECHEVARRIAMENDOZA\OneDrive - Brunswick Corporation\Documents\Proyectos_Python\PysimpleGUI\Proyectos\comparador\excel\29736-1B.xlsx'
flexa = pd.read_excel(nombre_placement, engine='openpyxl')
placement = pd.DataFrame(flexa)
placement.rename(columns={'Ref.':'Reference'},inplace=True)
placement.rename(columns={'Board ':'Board'},inplace=True)
placement = placement[['Board','Part Number','Reference']]
placement.to_csv(r'C:\Users\CECHEVARRIAMENDOZA\OneDrive - Brunswick Corporation\Documents\Proyectos_Python\PysimpleGUI\Proyectos\comparador\csv\placement.csv',index=False)
placement_filter = placement[placement['Board']==1]

# Eliminar la columna "Board" del DataFrame filtrado
placement_filter.drop(columns=["Board"], inplace=True)

# Restablece el índice del DataFrame si es necesario
placement_filter.reset_index(drop=True, inplace=True)
#crear csv
placement_filter.to_csv(r'C:\Users\CECHEVARRIAMENDOZA\OneDrive - Brunswick Corporation\Documents\Proyectos_Python\PysimpleGUI\Proyectos\comparador\csv\placement_filtrado.csv',index=False)
#print(placement_filter)


#******************************************************************* COMPARACION ******************************************************************
comparacion = bom_filter.merge(placement_filter, on = ['Part Number','Reference'], how='outer',suffixes=('_izq', '_der'), indicator=True)
comparacion['_merge'] = comparacion['_merge'].replace({
    'left_only': 'Solo en BOM',
    'right_only': 'Solo en Placement',
    'both': 'En ambos archivos'
})

only_bom = comparacion[comparacion['_merge'] == 'left_only']
only_placement = comparacion[comparacion['_merge'] == 'right_only']
comparacion.to_csv(r'C:\Users\CECHEVARRIAMENDOZA\OneDrive - Brunswick Corporation\Documents\Proyectos_Python\PysimpleGUI\Proyectos\comparador\csv\comparacion.csv',index=False)
#print("only_bom:","\n",only_bom)
#print('only_placement:','\n',only_placement)
#print(comparacion)

