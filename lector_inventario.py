import os
from IPython.display import display_html
import qrcode
from PIL import Image, ImageDraw, ImageFont
import numpy as np
import pandas as pd
import xlsxwriter
import openpyxl as xl
from datetime import datetime
from openpyxl import load_workbook
fecha = '29101738'

inventario_html = 'G:\\Unidades compartidas\\Centro Excelencia Logística CCU S.A\\1. Productividad Bodegas\\22. WEP-Modelo\\26. Inventario\\'+fecha+'.xls'
inventario = pd.read_html(inventario_html, header=0)[0]

#for index, row in inventario.iloc[0:].iterrows():
#        code ="00"+str(row['SSCC'])
#        print(code)
df_inventario = pd.DataFrame(inventario, columns=['Fecha','Area','UbicaciÃ³n','SSCC'])
directory = 'G:\\Unidades compartidas\\Centro Excelencia Logística CCU S.A\\1. Productividad Bodegas\\22. WEP-Modelo\\26. Inventario\\'
file = str(fecha)+'_vdefinitiva.csv'
if not os.path.exists(directory):
    os.makedirs(directory)
df_inventario.to_csv(os.path.join(directory, file))
