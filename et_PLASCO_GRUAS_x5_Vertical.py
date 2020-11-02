import qrcode
from PIL import Image, ImageDraw, ImageFont
import numpy as np
import pandas as pd
import xlsxwriter
import openpyxl as xl
from datetime import datetime
from openpyxl import load_workbook

excel_file = r'C:\Users\jarellan\Desktop\qr_label\IN\GRUAS_PLASCO'
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
        code = row['Posicion']
        abr = str(row['Descr'])
        abr2 = str(row['Descr2'])
        if abr2 == 'nan':
                abr2=""
        img = Image.new('RGB', (325,150), color=(255, 255, 255))
        fnt = ImageFont.truetype(r'C:\Users\jarellan\Desktop\qr_label\IN\Futura Medium Italic font.ttf', 25)
        fnt2 = ImageFont.truetype(r'C:\Users\jarellan\Desktop\qr_label\IN\Futura Bold font.ttf', 30)
        d = ImageDraw.Draw(img)
        d.text((5, 5), str(code), font=fnt2, fill=(0, 0, 0))
        d.polygon([(185, 20),(185,30),(195, 25)],  fill=(0, 0, 0))
        d.text((5, 50), str(abr), font=fnt, fill=(0, 0, 0))
        d.text((5, 100), str(abr2), font=fnt, fill=(0, 0, 0))

        qr_big = qrcode.QRCode(error_correction=qrcode.constants.ERROR_CORRECT_H, box_size=5, border=1)
        qr_big.add_data(code)
        qr_big.make()
        img_qr_big = qr_big.make_image().convert('RGB')
        img.paste(img_qr_big, (200, 5))


        # save that beautiful picture
        id_etiqueta += 1
        #imgs_comb.save('C:\\Users\\jarellan\\Desktop\\qr_label\\OUT\\PLASCO\\etiqueta' + row['Posicion'] + '_' + str(id_etiqueta + k) + '.jpg', dpi=(300, 300))
        print('Impresión de ' + str(row['Posicion']) + '_' + str(k))
        imgs_v2.append(img)
        # df.loc[last_row]=[str(code),str(rectangle),str(circle),str(triangle),int(id_etiqueta + k),row['Posicion'] + '_' + str(id_etiqueta + k)]
        # print(df)
        if k % 5 == 0:
                img_vf = Image.new('RGB', (325, 750), color=(255, 255, 255))
                d2 = ImageDraw.Draw(img_vf)
                posv_1 = (0, 0)
                posv_2 = (0, 150)
                posv_3 = (0, 300)
                posv_4 = (0, 450)
                posv_5 = (0, 600)
                img_vf.paste(imgs_v2[0], posv_1)
                d2.line((0, 150, 325, 150), fill=(0, 0, 0), width=3)
                img_vf.paste(imgs_v2[1], posv_2)
                d2.line((0, 300, 325, 300), fill=(0, 0, 0), width=3)
                img_vf.paste(imgs_v2[2], posv_3)
                d2.line((0, 450, 325, 450), fill=(0, 0, 0), width=3)
                img_vf.paste(imgs_v2[3], posv_4)
                d2.line((0, 600, 325, 600), fill=(0, 0, 0), width=3)
                img_vf.paste(imgs_v2[4], posv_5)
                imgs_v2.clear()
                img_vf.save('C:\\Users\\jarellan\\Desktop\\qr_label\\OUT\\PLASCO\\' + str(row['Posicion']) + '_' + str(id_etiqueta + k) + '.jpg', dpi=(300, 300))
        k += 1
