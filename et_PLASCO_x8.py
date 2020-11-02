import qrcode
from PIL import Image, ImageDraw, ImageFont
import numpy as np
import pandas as pd
import xlsxwriter
import openpyxl as xl
from datetime import datetime
from openpyxl import load_workbook

excel_file = r'C:\Users\jarellan\Desktop\qr_label\IN\Ubicaciones_PLASCO'
ubicaciones = pd.read_excel(excel_file+".xlsx")

# My image is a 200x374 jpeg that is 102kb large
logo_ccu = Image.open('logo_ccu.png')

# I downsize the image with an ANTIALIAS filter (gives the highest quality)
logo_ccu = logo_ccu.resize((20, 20), Image.ANTIALIAS)
logo_ccu.save("logo_ccu_escalado.png", quality=95)
# The saved downsized image size is 24.8kb
logo_ccu.save("logo_ccu_escalado_opt.png", optimize=True, quality=95)
# The saved downsized image size is 22.9kb
id_etiqueta = 0
k=1
# hacer diccionario de colores
imgs = []
for index, row in ubicaciones.iloc[0:].iterrows():
    code = row['Posicion']
    img = Image.new('RGB', (250,70), color=(255, 255, 255))
    fnt = ImageFont.truetype(r'C:\Users\jarellan\Desktop\qr_label\IN\Futura Bold font.ttf', 25)
    d = ImageDraw.Draw(img)


    qr_big = qrcode.QRCode(error_correction=qrcode.constants.ERROR_CORRECT_H, box_size=2, border=0)
    qr_big.add_data(code)
    qr_big.make()
    img_qr_big = qr_big.make_image().convert('RGB')

    d.text((img_qr_big.size[0]+15, img.size[1] // 2-20), code, font=fnt, fill=(0, 0, 0))
    d.line((0, 65, 300, 65), fill=(0, 0, 0), width=5)
    pos2 = (10, (img.size[1] - img_qr_big.size[1]) // 2)
    img.paste(img_qr_big,pos2)
    id_etiqueta += 1
    imgs.append(img)
    #img.save('C:\\Users\\jarellan\\Desktop\\qr_label\\OUT\\PLASCO\\etiqueta' + row['Posicion'] + '_' + str(id_etiqueta + k) + '.pdf',dpi=(300, 300))
    print('Impresión de ' + row['Posicion'] + '_' + str(k))
    if k%8==0:
        img_vf = Image.new('RGB', (500,280), color=(255, 255, 255))
        posv_2 = (0,0)
        posv_3 = (0,70)
        posv_4 = (0, 140)
        posv_5 = (0, 210)
        posv_6 = (250, 0)
        posv_7 = (250, 70)
        posv_8 = (250, 140)
        posv_9 = (250, 210)
        img_vf.paste(imgs[0],posv_2)
        img_vf.paste(imgs[1],posv_3)
        img_vf.paste(imgs[2], posv_4)
        img_vf.paste(imgs[3], posv_5)
        img_vf.paste(imgs[4], posv_6)
        img_vf.paste(imgs[5], posv_7)
        img_vf.paste(imgs[6], posv_8)
        img_vf.paste(imgs[7], posv_9)
        imgs.clear()
        img_vf.save('C:\\Users\\jarellan\\Desktop\\qr_label\\OUT\\PLASCO\\etiqueta_doble'+row['Posicion'] + '_' + str(id_etiqueta + k)+'.jpg', dpi=(300, 300))
    k += 1

