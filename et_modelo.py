import qrcode
from PIL import Image, ImageDraw, ImageFont
import numpy as np
import pandas as pd
import xlsxwriter
import openpyxl as xl
from datetime import datetime
from openpyxl import load_workbook

colores = {"Z":(255,255,255),"A":(213,43,30),"B":(0,133,66),"C":(0,101,189),"D":(240,171,0),"F":(215,31,133),"G":(117,48,119),"H":(255,88,0),"I":(249,227,0),"J":(0,0,0),"P":(0,38,100),"Q":(104,69,13),"M":(198,191,110),"L":(78,84,87),"N":(178,175,175),"S":(0,161,222),"T":(127,127,126)}
excel_file = r'C:\Users\jarellan\Desktop\qr_label\IN\Ubicaciones_MOD'
ubicaciones = pd.read_excel(excel_file+".xlsx")

# My image is a 200x374 jpeg that is 102kb large
logo_ccu = Image.open('logo_ccu.png')

# I downsize the image with an ANTIALIAS filter (gives the highest quality)
logo_ccu = logo_ccu.resize((100, 100), Image.ANTIALIAS)
logo_ccu.save("logo_ccu_escalado.png", quality=95)
# The saved downsized image size is 24.8kb
logo_ccu.save("logo_ccu_escalado_opt.png", optimize=True, quality=95)
# The saved downsized image size is 22.9kb
id_etiqueta = 0
k=0
# hacer diccionario de colores

for index, row in ubicaciones.iloc[0:].iterrows():
    images = []
    code = row['Posicion']
    abr = str(row['Abr'])
    print(str(row['Letra']))
    color = colores[str(row['Letra'])]
    img = Image.new('RGB', (797,1124), color=color)
    fnt = ImageFont.truetype(r'C:\Users\jarellan\Desktop\qr_label\IN\Futura Bold font.ttf', 280)
    d = ImageDraw.Draw(img)
    if str(row['Letra']) == 'I' or str(row['Letra']) == 'N'or str(row['Letra']) == 'M'or str(row['Letra']) == 'D' or str(row['Letra']) == 'Z':
        d.text((25, 10), abr, font=fnt, fill=(0, 0, 0))
    else:
        d.text((25, 10), abr, font=fnt, fill=(255, 255, 255))
    # d.line((70, 190, 750, 190), fill=(0, 0, 0), width=1)

    qr_big = qrcode.QRCode(error_correction=qrcode.constants.ERROR_CORRECT_H, box_size=25, border=2)
    qr_big.add_data(code)
    qr_big.make(fit=True)
    img_qr_big = qr_big.make_image().convert('RGB')
    pos = ((img_qr_big.size[0] - logo_ccu.size[0]) // 2, (img_qr_big.size[1] - logo_ccu.size[1]) // 2)
    img_qr_big.paste(logo_ccu, pos)

    pos2 = ((img.size[0] - img_qr_big.size[0]) // 2, 370)
    img.paste(img_qr_big,pos2)
    id_etiqueta += 1
    img.save('C:\\Users\\jarellan\\Desktop\\qr_label\\OUT\\MODELO\\STAGES\\'+str(row['Posicion']) + '_' + str(id_etiqueta + k) + '.pdf',dpi=(600, 600))
    k += 1