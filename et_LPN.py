import qrcode
from PIL import Image, ImageDraw, ImageFont
import numpy as np
import pandas as pd
import xlsxwriter
import openpyxl as xl
from datetime import datetime
from openpyxl import load_workbook

excel_file = r'C:\Users\jarellan\Desktop\qr_label\IN\Ubicaciones_MOD_selec'
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
        code =str(row['Posicion'])
        img = Image.new('RGB', (260,260), color=(255, 255, 255))
        fnt = ImageFont.truetype(r'C:\Users\jarellan\Desktop\qr_label\IN\Futura Bold font.ttf', 70)
        d = ImageDraw.Draw(img)
        d.text((20, 5), str(code), font=fnt, fill=(0, 0, 0))

        qr_big = qrcode.QRCode(error_correction=qrcode.constants.ERROR_CORRECT_H, box_size=6, border=2)
        qr_big.add_data(code)
        qr_big.make()
        img_qr_big = qr_big.make_image().convert('RGB')
        img.paste(img_qr_big, (30, 60))


        # save that beautiful picture
        id_etiqueta += 1
        #imgs_comb.save('C:\\Users\\jarellan\\Desktop\\qr_label\\OUT\\PLASCO\\etiqueta' + row['Posicion'] + '_' + str(id_etiqueta + k) + '.jpg', dpi=(300, 300))
        print('Impresión de ' + str(row['Posicion']) + '_' + str(k))
        imgs_v2.append(img)
        # df.loc[last_row]=[str(code),str(rectangle),str(circle),str(triangle),int(id_etiqueta + k),row['Posicion'] + '_' + str(id_etiqueta + k)]
        # print(df)

        img.save('C:\\Users\\jarellan\\Desktop\\qr_label\\OUT\\MODELO\\' + str(row['Posicion']) + '_' + str(id_etiqueta + k) + '.jpg', dpi=(300, 300))
        k += 1
