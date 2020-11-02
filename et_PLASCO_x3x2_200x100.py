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

id_etiqueta = 0
k=1
# hacer diccionario de colores
imgs_v2 = []
for index, row in ubicaciones.iloc[0:].iterrows():
    images = []
    # Con crop corto la imagen
    # face = Image.open('logo_ccu.png')
    code = row['Posicion']
    img = Image.new('RGB', (360, 108), color=(255, 255, 255))
    fnt = ImageFont.truetype(r'C:\Users\jarellan\Desktop\qr_label\IN\Futura Bold font.ttf', 30)
    fnt2 = ImageFont.truetype(r'C:\Users\jarellan\Desktop\qr_label\IN\Futura Bold font.ttf', 20)
    d = ImageDraw.Draw(img)
    d.text((50, 10), "PLASCO", font=fnt2, fill=(0, 0, 0))
    d.text((5, 40), code, font=fnt, fill=(0, 0, 0))

    qr_big = qrcode.QRCode(error_correction=qrcode.constants.ERROR_CORRECT_H, box_size=3, border=1)
    qr_big.add_data(code)
    qr_big.make()
    img_qr_big = qr_big.make_image().convert('RGB')
    img.paste(img_qr_big, (240, 20))

    # save that beautiful picture
    id_etiqueta += 1
    # imgs_comb.save('C:\\Users\\jarellan\\Desktop\\qr_label\\OUT\\PLASCO\\etiqueta' + row['Posicion'] + '_' + str(id_etiqueta + k) + '.jpg', dpi=(300, 300))
    print('Impresión de ' + row['Posicion'] + '_' + str(k))
    imgs_v2.append(img)
    # df.loc[last_row]=[str(code),str(rectangle),str(circle),str(triangle),int(id_etiqueta + k),row['Posicion'] + '_' + str(id_etiqueta + k)]
    # print(df)
    if k % 6 == 0:
        img_vf = Image.new('RGB', (750, 325), color=(255, 255, 255))
        d2 = ImageDraw.Draw(img_vf)
        posv_1 = (0, 0)
        posv_2 = (0, 110)
        posv_3 = (0, 220)
        posv_4 = (380, 0)
        posv_5 = (380, 110)
        posv_6 = (380, 220)
        img_vf.paste(imgs_v2[0], posv_1)
        img_vf.paste(imgs_v2[1], posv_2)
        img_vf.paste(imgs_v2[2], posv_3)
        img_vf.paste(imgs_v2[3], posv_4)
        img_vf.paste(imgs_v2[4], posv_5)
        img_vf.paste(imgs_v2[5], posv_6)
        d2.line((375, 0, 375, 325), fill=(0, 0, 0), width=3)
        d2.line((0, 115, 750, 115), fill=(0, 0, 0), width=3)
        d2.line((0, 225, 750, 225), fill=(0, 0, 0), width=3)

        imgs_v2.clear()
        img_vf.save('C:\\Users\\jarellan\\Desktop\\qr_label\\OUT\\PLASCO\\'+ row['Posicion'] + '_' + str(
            id_etiqueta + k) + '.jpg', dpi=(300, 300))
    k += 1
