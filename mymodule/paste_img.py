import openpyxl as px
from openpyxl.drawing.image import Image

#
xxx_name = '002'
xlfile_path = './Result/1st_Wit_' + xxx_name + '.xlsx'
imgfile_path = './Data/Image_001-030/img_' + xxx_name + '/' + xxx_name + '_'

wb = px.load_workbook(xlfile_path)

sheet_names1 = ['0','1','2','3','4']
sheet_names2 = {'A': ' (1)','B': ' (2)','C': ' (3)','D': ' (4)','E': ' (5)','F': ' (6)','G': ' (7)','H': ' (8)','I': ' (9)','J': ' (10)'}

for sheet_name in sheet_names1:

    ws = wb[sheet_name]
    img = Image(imgfile_path + sheet_name + '.JPG')
    img.width = 1378 #35 cm (100dpi環境下)
    img.height = 775 #19.69 cm

    ws.add_image(img, 'A1')

for s_name, f_name in sheet_names2.items():

    ws = wb[s_name]
    img = Image(imgfile_path + f_name + '.png')
    img.width = 1378 #35 cm (100dpi環境下)
    img.height = 775 #19.69 cm

    ws.add_image(img, 'A1')
    
wb.save(xlfile_path)