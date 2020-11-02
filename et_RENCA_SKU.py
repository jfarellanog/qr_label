import qrcode
from PIL import Image, ImageDraw, ImageFont
import numpy as np
import pandas as pd
import xlsxwriter
import openpyxl as xl
from datetime import datetime
from openpyxl import load_workbook

excel_file = r'C:\Users\jarellan\Desktop\qr_label\IN\SKU_Renca'
sku = pd.read_excel(excel_file+".xlsx")

# My image is a 200x374 jpeg that is 102kb large
logo_ccu = Image.open('logo_ccu.png')
logo_ccu = logo_ccu.resize((80, 70), Image.ANTIALIAS)
logo_ccu.save("logo_ccu_escalado.png", quality=95)
logo_ccu.save("logo_ccu_escalado_opt.png", optimize=True, quality=95)
ubicacion = []
fecha = datetime.today().strftime("%d-%m-%Y")
k=1
id_etiqueta = 0
for index, row in sku.iloc[0:].iterrows():
        images = []
        id_sku_inv = row['id_sku_inv']
        id_sku_venta = row['id_sku_venta']
        descr_sku = row['descr_sku']
        id_foto = row['foto']
        try:
                foto = Image.open('C:\\Users\\jarellan\\Desktop\\qr_label\\IN\\DELFOS\\'+id_foto)
        except:
                foto = Image.open('C:\\Users\\jarellan\\Desktop\\qr_label\\IN\\DELFOS\\Para SKUs sin foto.jpg')
        foto = foto.resize((200, 300), Image.ANTIALIAS)
        foto.save('C:\\Users\\jarellan\\Desktop\\qr_label\\IN\\DELFOS\\ESCALADO\\'+id_sku_inv+"_escalado.png",  optimize=True,quality=95)

        cajas_por_pallet = row['cajas_por_pallet']
        cajas_por_capa = row['cajas_por_capa']
        capas_x_pallet = row['capas_x_pallet']
        ubicacion = row['Ubication']
        vida_util = row['vida_util']
        if int(vida_util)>=999:
                vida_util=' '


        img = Image.new('RGB', (500, 325), color=(255, 255, 255))
        fnt = ImageFont.truetype(r'C:\Users\jarellan\Desktop\qr_label\IN\Futura Bold font.ttf', 45)
        fnt2 = ImageFont.truetype(r'C:\Users\jarellan\Desktop\qr_label\IN\Futura Light font.ttf', 15)
        fnt3 = ImageFont.truetype(r'C:\Users\jarellan\Desktop\qr_label\IN\Futura Light font.ttf', 62-len(str(descr_sku)))
        fnt4 = ImageFont.truetype(r'C:\Users\jarellan\Desktop\qr_label\IN\Futura Bold font.ttf', 60)
        d = ImageDraw.Draw(img)
        print(len(str(descr_sku)))

        d.text((5, 35), "ID_SKU_INVENTARIO", font=fnt2, fill=(0, 0, 0))
        d.text((110, 5), str(id_sku_inv), font=fnt4, fill=(0, 0, 0))
        d.line((0, 80, 500, 80), fill=(0, 0, 0), width=1)

        d.text((5, 90), "DESCRIPCIÓN", font=fnt2, fill=(0, 0, 0))
        d.text((5, 110), str(descr_sku), font=fnt3, fill=(0, 0, 0))


        d.line((0, 160, 500, 160), fill=(0, 0, 0), width=1)

        d.text((5, 170), "CAJAS_POR_PALLET", font=fnt2, fill=(0, 0, 0))
        d.text((5, 180), str(cajas_por_pallet), font=fnt, fill=(0, 0, 0))

        d.text((150, 170), "CAMADAS_POR_PALLET", font=fnt2, fill=(0, 0, 0))
        d.text((160, 180), str(capas_x_pallet), font=fnt, fill=(0, 0, 0))

        d.text((270, 170), "CAJAS_POR_CAMADA", font=fnt2, fill=(0, 0, 0))
        d.text((280, 180), str(cajas_por_capa), font=fnt, fill=(0, 0, 0))

        d.text((400, 170), "VIDA_UTIL", font=fnt2, fill=(0, 0, 0))
        d.text((400, 180), str(vida_util), font=fnt, fill=(0, 0, 0))

        d.line((0, 240, 500, 240), fill=(0, 0, 0), width=1)

        d.text((5, 250), "UBIC REGLA REAP 20.09.2020", font=fnt2, fill=(0, 0, 0))
        d.text((5, 270), str(ubicacion), font=fnt, fill=(0, 0, 0))

        qr_big = qrcode.QRCode(error_correction=qrcode.constants.ERROR_CORRECT_H, box_size=6, border=2)
        qr_big.add_data(id_sku_inv)
        qr_big.make()
        img_qr_big = qr_big.make_image().convert('RGB')
        img.paste(img_qr_big, (360, 5))


        img_vf = Image.new('RGB', (750, 325), color=(255, 255, 255))
        d2 = ImageDraw.Draw(img_vf)
        posv_1 = (0, 0)
        posv_2 = (550, 0)
        img_vf.paste(img, posv_1)
        img_vf.paste(foto, posv_2)

        id_etiqueta += 1
        #img_vf.save('C:\\Users\\jarellan\\Desktop\\qr_label\\OUT\\RENCA\\' + str(id_sku_inv) + '_' + str(ubicacion) + '.jpg', dpi=(300, 300))
        img_vf.save('C:\\Users\\jarellan\\Desktop\\qr_label\\OUT\\RENCA\\' + str(k) +'.jpg', dpi=(300, 300))
        print('Impresión de ' + str(id_sku_inv) + '_' + str(ubicacion))
        k += 1
