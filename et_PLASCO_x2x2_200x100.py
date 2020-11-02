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
imgs_v2 = []
for index, row in ubicaciones.iloc[0:].iterrows():
    images = []
    # Con crop corto la imagen
    # face = Image.open('logo_ccu.png')
    code = row['Posicion']
    img = Image.new('RGB', (360, 150), color=(255, 255, 255))
    fnt = ImageFont.truetype(r'C:\Users\jarellan\Desktop\qr_label\IN\Futura Bold font.ttf', 30)
    fnt2 = ImageFont.truetype(r'C:\Users\jarellan\Desktop\qr_label\IN\Futura Bold font.ttf', 40)
    d = ImageDraw.Draw(img)
    d.text((50, 10), "PLASCO", font=fnt2, fill=(0, 0, 0))
    d.text((5, 60), code, font=fnt, fill=(0, 0, 0))

    qr_big = qrcode.QRCode(error_correction=qrcode.constants.ERROR_CORRECT_H, box_size=6, border=0)
    qr_big.add_data(code)
    qr_big.make()
    img_qr_big = qr_big.make_image().convert('RGB')
    pos = ((img_qr_big.size[0] - logo_ccu.size[0]) // 2, (img_qr_big.size[1] - logo_ccu.size[1]) // 2)
    img_qr_big.paste(logo_ccu, pos)
    img.paste(img_qr_big, (210, 0))

    # save that beautiful picture
    id_etiqueta += 1
    # imgs_comb.save('C:\\Users\\jarellan\\Desktop\\qr_label\\OUT\\PLASCO\\etiqueta' + row['Posicion'] + '_' + str(id_etiqueta + k) + '.jpg', dpi=(300, 300))
    print('Impresión de ' + row['Posicion'] + '_' + str(k))
    imgs_v2.append(img)
    # df.loc[last_row]=[str(code),str(rectangle),str(circle),str(triangle),int(id_etiqueta + k),row['Posicion'] + '_' + str(id_etiqueta + k)]
    # print(df)
    if k % 4 == 0:
        img_vf = Image.new('RGB', (750, 325), color=(255, 255, 255))
        d2 = ImageDraw.Draw(img_vf)
        posv_1 = (0, 0)
        posv_2 = (0, 170)
        posv_3 = (380, 0)
        posv_4 = (380, 170)
        img_vf.paste(imgs_v2[0], posv_1)
        img_vf.paste(imgs_v2[1], posv_2)
        img_vf.paste(imgs_v2[2], posv_3)
        img_vf.paste(imgs_v2[3], posv_4)
        d2.line((0, 325, 750, 325), fill=(0, 0, 0), width=3)
        d2.line((375, 0, 375, 325), fill=(0, 0, 0), width=3)
        d2.line((0, 160, 750, 160), fill=(0, 0, 0), width=3)
        imgs_v2.clear()
        img_vf.save('C:\\Users\\jarellan\\Desktop\\qr_label\\OUT\\PLASCO\\'+ row['Posicion'] + '_' + str(
            id_etiqueta + k) + '.jpg', dpi=(300, 300))
    k += 1
