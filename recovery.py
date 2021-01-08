#画像のパラメター(角度・時間)を，過去fileから，templateを変えるなどして新規作成したfileに移す
import openpyxl as px
from mymodule import my_round

xxx_name = ['001','002','003','004','005','006','007','008','009','010','011','012','013','014','015','016','017','018','019','020','021','022','023','024','025','026','027','028','029','030']

for i in range(0,1):
    print(xxx_name[i])
    witxl_path = './Result/001-030/Wit/1st_Wit_' + xxx_name[i] + '.xlsx'
    fxl_path = './Result/xlsx_template/failed/pre_Wit_' + xxx_name[i] + '.xlsx'

    wbC = px.load_workbook(witxl_path)
    wbF = px.load_workbook(fxl_path)

    sheet_names1 = ['0','1','2','3','4']
    sheet_names2 = ['A','B','C','D','E','F','G','H','I','J']

    #past_img.py len関数でやればsheet_name1,2という区別はいらない
    for sheet_name in sheet_names1:
        wsC = wbC[sheet_name]
        wsF = wbF[sheet_name]
        wsC['V3'].value = wsF['V3'].value 
        wsC['V4'].value = wsF['V4'].value 
    
    for sheet_name in sheet_names2:
        wsC = wbC[sheet_name]
        wsF = wbF[sheet_name]
        wsC['V3'].value = wsF['V3'].value 
        wsC['V4'].value = wsF['V4'].value 
        wsC['V5'].value = wsF['V5'].value 
        
    wbC.save(witxl_path)