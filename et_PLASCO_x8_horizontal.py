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
        abr = str(row['Texto'])
        abr2 = str(row['Texto2'])
        abr3 = str(row['Texto3'])
        img = Image.new('RGB', (90,325), color=(255, 255, 255))
        fnt = ImageFont.truetype(r'C:\Users\jarellan\Desktop\qr_label\IN\Futura Bold font.ttf', 25)
        fnt2 = ImageFont.truetype(r'C:\Users\jarellan\Desktop\qr_label\IN\Futura Bold font.ttf', 45)
        d = ImageDraw.Draw(img)
        d.text((10, 50), str(abr), font=fnt, fill=(0, 0, 0))
        d.text((20, 100), str(abr2), font=fnt, fill=(0, 0, 0))
        d.text((20, 150), "-0"+str(abr3), font=fnt2, fill=(0, 0, 0))


        qr_big = qrcode.QRCode(error_correction=qrcode.constants.ERROR_CORRECT_H, box_size=3, border=1)
        qr_big.add_data(code)
        qr_big.make()
        img_qr_big = qr_big.make_image().convert('RGB')
        img.paste(img_qr_big,(5,220))


        # save that beautiful picture
        id_etiqueta += 1
        #imgs_comb.save('C:\\Users\\jarellan\\Desktop\\qr_label\\OUT\\PLASCO\\etiqueta' + row['Posicion'] + '_' + str(id_etiqueta + k) + '.jpg', dpi=(300, 300))
        print('Impresión de ' + row['Posicion'] + '_' + str(k))
        imgs_v2.append(img)
        # df.loc[last_row]=[str(code),str(rectangle),str(circle),str(triangle),int(id_etiqueta + k),row['Posicion'] + '_' + str(id_etiqueta + k)]
        # print(df)
        if k % 8 == 0:
                img_vf = Image.new('RGB', (750, 325), color=(255, 255, 255))
                d2 = ImageDraw.Draw(img_vf)
                posv_1 = (0, 0)
                posv_2 = (93, 0)
                posv_3 = (186, 0)
                posv_4 = (279, 0)
                posv_5 = (372, 0)
                posv_6 = (465, 0)
                posv_7 = (558, 0)
                posv_8 = (651, 0)
                img_vf.paste(imgs_v2[0], posv_1)
                d2.line((93, 0, 93, 325), fill=(0, 0, 0), width=3)
                img_vf.paste(imgs_v2[1], posv_2)
                d2.line((186, 0, 186, 325), fill=(0, 0, 0), width=3)
                img_vf.paste(imgs_v2[2], posv_3)
                d2.line((279, 0, 279, 325), fill=(0, 0, 0), width=3)
                img_vf.paste(imgs_v2[3], posv_4)
                d2.line((372, 0, 372, 325), fill=(0, 0, 0), width=3)
                img_vf.paste(imgs_v2[4], posv_5)
                d2.line((465, 0, 465, 325), fill=(0, 0, 0), width=3)
                img_vf.paste(imgs_v2[5], posv_6)
                d2.line((558,0, 558, 325), fill=(0, 0, 0), width=3)
                img_vf.paste(imgs_v2[6], posv_7)
                d2.line((651, 0, 651, 325), fill=(0, 0, 0), width=3)
                img_vf.paste(imgs_v2[7], posv_8)
                imgs_v2.clear()
                img_vf.save('C:\\Users\\jarellan\\Desktop\\qr_label\\OUT\\PLASCO\\etiqueta_doble' + row['Posicion'] + '_' + str(id_etiqueta + k) + '.jpg', dpi=(300, 300))
        k += 1
