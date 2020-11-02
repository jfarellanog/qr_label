import qrcode
from PIL import Image, ImageDraw, ImageFont
import numpy as np
import pandas as pd
import xlsxwriter
import openpyxl as xl
from datetime import datetime
from openpyxl import load_workbook

excel_file = r'C:\Users\jarellan\Desktop\qr_label\IN\Ubicaciones_MOD_SELEC_P'
ubicaciones = pd.read_excel(excel_file+".xlsx")

ubicacion = []
fecha = datetime.today().strftime("%d-%m-%Y")
k=1
id_etiqueta = 0
imgs_v2 = []
for index, row in ubicaciones.iloc[0:].iterrows():
        images = []
        # Con crop corto la imagen
        # face = Image.open('logo_ccu.png')
        code = str(row['Posicion'])
        img = Image.new('RGB', (534,293), color=(255, 255, 255))
        fnt = ImageFont.truetype(r'C:\Users\jarellan\Desktop\qr_label\IN\Futura Bold font.ttf', 65)
        fnt2 = ImageFont.truetype(r'C:\Users\jarellan\Desktop\qr_label\IN\Futura Bold font.ttf', 50)
        d = ImageDraw.Draw(img)
        d.text((90, 0), str(code), font=fnt, fill=(0, 0, 0))

        qr_big = qrcode.QRCode(error_correction=qrcode.constants.ERROR_CORRECT_H, box_size=8, border=1)
        qr_big.add_data(code)
        qr_big.make()
        img_qr_big = qr_big.make_image().convert('RGB')
        img.paste(img_qr_big, (300, 80))

        r = int(row['R'])
        c = int(row['C'])
        t = int(row['T'])

        if r < 10:
                rectangle = '0' + str(r)
        else:
                rectangle = str(r)
        if c < 10:
                circle = '0' + str(c)
        else:
                circle = str(c)
        if t < 10:
                triangle = '0' + str(t)
        else:
                triangle = str(t)

        d.rectangle((50, 100, 140, 190), fill="black", outline="black")
        d.text((70, 110), rectangle, font=fnt2, fill=(255, 255, 255))
        d.ellipse((170, 100, 260, 190), fill="black", outline="black")
        d.text((190, 115), circle, font=fnt2, fill=(255, 255, 255))
        d.polygon([(160, 205), (100, 285), (220, 285)], fill="black", outline="black")
        d.text((140, 225), triangle, font=fnt2, fill=(255, 255, 255))

        # save that beautiful picture
        id_etiqueta += 1
        #imgs_comb.save('C:\\Users\\jarellan\\Desktop\\qr_label\\OUT\\PLASCO\\etiqueta' + row['Posicion'] + '_' + str(id_etiqueta + k) + '.jpg', dpi=(300, 300))
        print('Impresión de ' + str(row['Posicion']) + '_' + str(k))
        imgs_v2.append(img)
        # df.loc[last_row]=[str(code),str(rectangle),str(circle),str(triangle),int(id_etiqueta + k),row['Posicion'] + '_' + str(id_etiqueta + k)]
        # print(df)

        img.save('C:\\Users\\jarellan\\Desktop\\qr_label\\OUT\\MODELO\\SELECTIVO-P\\' + str(row['Posicion']) + '_'+str(k)+'.jpg', dpi=(300, 300))
        k += 1
