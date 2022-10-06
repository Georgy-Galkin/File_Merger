# File_Merger

import pandas as pd
import datetime
from dateutil.relativedelta import relativedelta
import PySimpleGUI as sg

import xlwings as xw
sg.theme('DarkBlue')

layout = [
    [sg.T("Input File for processing:", s=25,justification="r"), sg.I(key="-IN-", s=70), sg.FileBrowse(file_types=(("Excel Files", "*.xls*")))],
    [sg.T("Input Main file with all data:",s=25, justification="r"), sg.I(key="-MAIN-"), sg.FileBrowse(file_types=(("Text Files", "*.txt"),))],
    [sg.T("Input Storage Folder:", s=25,justification="r"), sg.I(key="-FOLDER-"), sg.FolderBrowse()],
    [sg.T("Input File with GEO:", s=25,justification="r"), sg.I(key="-GEO-", s=70), sg.FileBrowse(file_types=(("Excel Files", "*.xls*")))],
    [sg.Text("Choose data modification type",s=25,justification="r")],
    [sg.Listbox(values=['Еженедельные стоки ТТ', 'Еженедельные стоки РЦ', 'Еженедельные Продажи', 'Ежемесячные Продажи RMC', 'Ежемесячные Продажи RRP', 'Ежемесячные Остатки'], size=(60, 10), select_mode='single', key='-DESTINATION-')],
     
    [sg.Submit()]
]

window = sg.Window('Программа обработки продаж Магнит', layout)

event, values = window.read()

user_path=str(values['-IN-'])
selection=str(values['-DESTINATION-'][0])
main_filepath=str(values['-MAIN-'])
result_file=str(values['-FOLDER-'])
result_file_geo=str(values['-GEO-'])
window.close()

user_path=user_path.replace("\\","\\\\")
main_filepath=main_filepath.replace("\\","\\\\")
result_file=result_file.replace("\\","\\\\")
result_file_geo=result_file_geo.replace("\\","\\\\")


print("Extracting Data")
print("----")

if selection=="Ежемесячные Продажи RMC":

    try:
        #--------------------------
        print("Начали обработку")
        print("----")
        #--------------------------
        sales=pd.read_excel(user_path,index_col=None,header=0)
        header_row=sales.index[sales.iloc[:,0] == 'Магазин'].tolist()
        header_row=header_row[0]
        header=sales.iloc[header_row]
        all_rows=header_row+1
        sales= sales[all_rows:]
        sales.columns=header
        sales=sales[sales['Магазин'].notnull()]
        print("Прочитали содержимое файла")
        print("----")
        #------------------------------------------------------------------------

        geo=sales[['Магазин','Формат','Филиал','РЦ (ОС)']]
        geo=geo.rename(columns={"Магазин": "Наименование ТТ", "РЦ (ОС)":"РЦ"})
        geo=geo.drop_duplicates()
        geo=geo[geo['Наименование ТТ']!="Grand Total"]
        geo=geo[geo['Наименование ТТ']!="Общий Итог"]
        geo=geo[geo['Наименование ТТ']!="Общий итог"]
        sales=sales.drop(['Формат','Филиал','РЦ (ОС)'],axis=1)





        #-----------------------------------------------------------------------
        sales=pd.melt(sales,id_vars='Магазин',var_name="SKU", value_name='sales')
        sales=sales[sales.sales.notnull()]
        sales=sales[sales['Магазин']!="Grand Total"]
        sales=sales[sales['Магазин']!="Общий Итог"]
        sales=sales[sales['Магазин']!="Общий итог"]
        sales['SKU']=sales['SKU'].str.lower()
        user_path=user_path.replace(".xlsm","")
        sales["DATE"]="01/"+user_path[-4:-2]+"/2022"

        #--------------------------
        print("Транспонировали и присвоили дату")
        print("----")
        print(sales['DATE'].unique()[0], " - Присвоенная дата")
        #--------------------------
        print("----")
        sales[['mrp1','mrp2']]=sales.SKU.str.split('мрц',expand=True)
        sales['mrp2']=sales['mrp2'].str.replace(" ","")
        sales['mrp2']=sales['mrp2'].str.slice(stop=3)
        sales['mrp2']=sales.mrp2.str.extract('(\d+)')
        sales['mrp2']=sales.mrp2.fillna(0)
        sales=sales.rename(columns={"mrp2": "MRP"})
        sales=sales.drop(['mrp1'],axis=1)
        #--------------------------
        print("Добавили МРЦ")
        print("----")
        print(sales[sales["MRP"]==0].SKU.unique()," - SKU без МРЦ")
        print("Обработали ",user_path)
        print("----")
        print("так выглядит обработанный массив")
        print("----")
        print(sales.head(5).to_string(index=False))
        print("----")
        #--------------------------
        print(user_path[-4:-2]," Полученная сумма продаж без учета блоков: ",sales.sales.sum())
        #--------------------------
    except:
        pass
        print(user_path,"НЕ ПОЛУЧИЛОСЬ")
    frame=sales 
    frame = frame.astype({"Магазин": str})
    wb = xw.Book(result_file_geo)
    ws = wb.sheets["GEO Monthly RMC"]
    ws.clear()
    ws["A1"].options(pd.DataFrame, header=1, index=False, expand='table').value = geo
    tbl_range = ws.range("A1").expand('table')
    ws.api.ListObjects.Add(1, ws.api.Range(tbl_range.address))
    ws.tables[0].name = "Monthly_RMC"
    
    wb.save()
    wb.close()
    #--------------------------    
    print("Теперь соединим файл с исходным")
    print("----")
    #current table
    current_df=pd.read_csv(main_filepath , delimiter="\t")
    current_df = current_df.astype({"Магазин": str})
    new_total=pd.concat([frame,current_df])
    new_total=new_total[new_total['SKU'].notnull()]
    print("Прочитали Файл")
    print("----")
    print("Добавили файл к исходному")
    #--------------------------
    new_total.to_csv(result_file+"\\Current_RMC.txt", index=None, sep='\t', mode='w+')
    print("Сохранили файл")
    print("----")
    print("Можно закрывать программу")
    #--------------------------
    
elif selection=="Ежемесячные Продажи RRP":

    try:
        #--------------------------
        print("Начали обработку")
        print("----")
        #--------------------------
        sales=pd.read_excel(user_path,index_col=None,header=0)
        
        header_row=sales.index[sales.iloc[:,0] == 'Магазин'].tolist()
        header_row=header_row[0]
        header=sales.iloc[header_row]
        all_rows=header_row+1
        sales= sales[all_rows:]
        sales.columns=header
        print("Прочитали содержимое файла")
        print("----")
        #------------------------------------------------------------------------

        geo=sales[['Магазин','Формат','Филиал','РЦ (ОС)']]
        geo=geo.rename(columns={"Магазин": "Наименование ТТ", "РЦ (ОС)":"РЦ"})
        geo=geo.drop_duplicates()
        geo=geo[geo['Наименование ТТ']!="Grand Total"]
        geo=geo[geo['Наименование ТТ']!="Общий Итог"]
        geo=geo[geo['Наименование ТТ']!="Общий итог"]
        

        sales=sales.drop(['Формат','Филиал','РЦ (ОС)'],axis=1)





        #-----------------------------------------------------------------------
        sales=pd.melt(sales,id_vars='Магазин',var_name="SKU", value_name='sales')
        sales=sales[sales.sales.notnull()]
        sales=sales[sales['Магазин']!="Grand Total"]
        sales=sales[sales['Магазин']!="Общий Итог"]
        sales=sales[sales['Магазин']!="Общий итог"]
        sales['SKU']=sales['SKU'].str.lower()
        user_path=user_path.replace(".xlsm","")
        sales["DATE"]="01/"+user_path[-4:-2]+"/2022"
        #--------------------------
        print("Транспонировали и присвоили дату")
        print("----")
        print(sales['DATE'].unique()[0], " - Присвоенная дата")
        print("----")
        #--------------------------

        print("Обработали ",user_path)
        print("----")
        print("так выглядит обработанный массив")
        print("----")
        print(sales.head(5).to_string(index=False))
        print("----")
        print(user_path[-4:-2],"  Полученная сумма продаж без учета блоков: ",sales.sales.sum())
        
    except:
        pass
        print(user_path,"НЕ ПОЛУЧИЛОСЬ")
    frame=sales
    frame = frame.astype({"Магазин": str})
    #----------------------------------------------------------------------------------
    wb = xw.Book(result_file_geo)
    ws = wb.sheets["GEO Monthly RRP"]
    ws.clear()
    ws["A1"].options(pd.DataFrame, header=1, index=False, expand='table').value = geo
    tbl_range = ws.range("A1").expand('table')
    ws.api.ListObjects.Add(1, ws.api.Range(tbl_range.address))
    ws.tables[0].name = "Monthly_RRP"
    wb.save()
    wb.close()
    #--------------------------   
    #--------------------------    
    print("Теперь соединим файл с исходным")
    #current table
    current_df=pd.read_csv(main_filepath , delimiter="\t")
    current_df = current_df.astype({"Магазин": str})
    print("Прочитали Файл")
    new_total=pd.concat([frame,current_df])
    new_total=new_total[new_total['SKU'].notnull()]
    print("Добавили файл к исходному")
    #--------------------------
    new_total.to_csv(result_file+"\\Current_RRP.txt", index=None, sep='\t', mode='w+')
    print("Сохранили файл")
    print("----")
    print("Можно закрывать файл")
    #--------------------------
    
elif selection=="Ежемесячные Остатки":    
    try:
        #--------------------------
        print("Начали обработку")
        print("----")
        #--------------------------
        sales=pd.read_excel(user_path,index_col=None,header=0)
        header_row=sales.index[sales.iloc[:,0] == 'Наименование ТТ'].tolist()
        header_row=header_row[0]
        header=sales.iloc[header_row]
        all_rows=header_row+1
        sales= sales[all_rows:]
        sales.columns=header
        print("Прочитали содержимое файла")
        print("----")
        sales=sales[sales['Наименование ТТ'].notnull()]
        #------------------------------------------------------------------------

        geo=sales[['Наименование ТТ',"Формат","РЦ"]]
       
        geo=geo.drop_duplicates()
        geo=geo[geo['Наименование ТТ']!="Grand Total"]
        geo=geo[geo['Наименование ТТ']!="Общий Итог"]
        geo=geo[geo['Наименование ТТ']!="Общий итог"]
        sales=sales.drop(['Формат','РЦ'],axis=1)





        #-----------------------------------------------------------------------
        sales=pd.melt(sales,id_vars='Наименование ТТ',var_name="SKU", value_name='stock')
        sales=sales[sales.stock.notnull()]
        sales=sales[sales['Наименование ТТ']!="Grand Total"]
        sales=sales[sales['Наименование ТТ']!="Общий Итог"]
        sales=sales[sales['Наименование ТТ']!="Общий итог"]
        sales['SKU']=sales['SKU'].str.lower()
        user_path=user_path.replace(".xlsm","")
        sales["DATE"]="01/"+user_path[-4:-2]+"/2022"
        #--------------------------
        print("Транспонировали и присвоили дату")
        print("----")
        print(sales['DATE'].unique()[0], " - Присвоенная дата")
        print("----")
        #--------------------------
        sales[['mrp1','mrp2']]=sales.SKU.str.split('мрц',expand=True)
        sales['mrp2']=sales['mrp2'].str.replace(" ","")
        sales['mrp2']=sales['mrp2'].str.slice(stop=3)
        sales['mrp2']=sales.mrp2.str.extract('(\d+)')
        sales['mrp2']=sales.mrp2.fillna(0)
        sales=sales.rename(columns={"mrp2": "MRP"})
        sales=sales.drop(['mrp1'],axis=1)
        #--------------------------
        print("Добавили МРЦ")
        print("----")
        print(sales[sales["MRP"]==0].SKU.unique()," - SKU без МРЦ")
        print("Обработали ",user_path)
        print("----")
        print("так выглядит обработанный массив")
        print("----")
        print(sales.head(5).to_string(index=False))
        print("----")
        
        print(user_path[-4:-2],"  Полученная сумма продаж без учета блоков: ",sales.stock.sum())
        
    except:
        pass
        print(user_path,"НЕ ПОЛУЧИЛОСЬ")
    frame=sales
    frame = frame.astype({"Наименование ТТ": str})
    #----------------------------------------------------------------------------------
    wb = xw.Book(result_file_geo)
    ws = wb.sheets["GEO Monthly Stock"]
    ws.clear()
    ws["A1"].options(pd.DataFrame, header=1, index=False, expand='table').value = geo
    tbl_range = ws.range("A1").expand('table')
    ws.api.ListObjects.Add(1, ws.api.Range(tbl_range.address))
    ws.tables[0].name = "Monthly_Stock"
    
    
    
    wb.save()
    wb.close()
    #----------------------------------------------------------------------------------
    #--------------------------    
    print("Теперь соединим файл с исходным")
    #current table
    current_df=pd.read_csv(main_filepath , delimiter="\t")
    current_df = current_df.astype({"Наименование ТТ": str})
    print("Прочитали Файл")
    new_total=pd.concat([frame,current_df])
    new_total=new_total[new_total['SKU'].notnull()]
    print("Добавили файл к исходному")
    new_total.to_csv(result_file+"\\Current_RMC_Stock.txt", index=None, sep='\t', mode='w+')
    print("Сохранили файл")
    #--------------------------
    print("----")
    print("Можно закрывать файл")
    
    

elif selection=="Еженедельные стоки РЦ":

    
    try:
        #--------------------------
        print("Начали обработку")
        #--------------------------


        sales=pd.read_excel(user_path,index_col=None,header=0)
        header_row=sales.index[sales.iloc[:,0] == 'РЦ'].tolist()
        header_row=header_row[0]
        header=sales.iloc[header_row]
        all_rows=header_row+1
        sales= sales[all_rows:]
        sales.columns=header
        sales=sales[sales['РЦ'].notnull()]
        print("Прочитали содержимое файла")
        print("----")
        sales=pd.melt(sales,id_vars='РЦ',var_name="SKU", value_name='stock')
        sales=sales[sales.stock.notnull()]
        sales=sales[sales['РЦ']!="Grand Total"]
        sales=sales[sales['РЦ']!="Общий Итог"]
        sales=sales[sales['РЦ']!="Общий итог"]
        sales['SKU']=sales['SKU'].str.lower()
        user_path=user_path.replace(".xlsm","")
        week = int(user_path[-2:])
        year = 2022
        date = datetime.date(year, 1, 1) + relativedelta(weeks=+week)
        #--------------------------
        print("Транспонировали и присвоили дату")
        #--------------------------
        print("----")

        sales["DATE"]=date
        sales["week"]=week
        sales[['mrp1','mrp2']]=sales.SKU.str.split('мрц',expand=True)
        sales['mrp2']=sales['mrp2'].str.replace(" ","")
        sales['mrp2']=sales['mrp2'].str.slice(stop=3)
        sales['mrp2']=sales.mrp2.str.extract('(\d+)')
        sales['mrp2']=sales.mrp2.fillna(0)
        sales=sales.rename(columns={"mrp2": "MRP"})
        sales=sales.drop(['mrp1'],axis=1)
        #--------------------------
        
        print("Присвоили дату: ",date)
        print("----")
        print("Добавили МРЦ")
        print("----")
        print(sales[sales["MRP"]==0].SKU.unique()," - SKU без МРЦ")
        print("Обработали ",user_path)
        print("----")
        print("так выглядит обработанный массив")
        print("----")
        print(sales.head(5).to_string(index=False))
        print("----")
        print(user_path[-2:]," Сумма без учета блоков: ",sales.stock.sum())
        print("----")
        
    except:
        pass
        print(user_path,"НЕ ПОЛУЧИЛОСЬ")
    frame=sales
    #--------------------------
    print("Теперь соединим файл с исходным")
    #current table
    current_df=pd.read_csv(main_filepath , delimiter="\t")
    
    print("Прочитали Файл")
    new_total=pd.concat([frame,current_df])
    new_total=new_total[new_total['SKU'].notnull()]
    print("Добавили файл к исходному")
    #--------------------------
    new_total.to_csv(result_file+"\\Current_weekly_Stock_DC.txt", index=None, sep='\t', mode='w+')
    print("Сохранили файл")
    print("----")
    print("Можно закрывать файл")
    
elif selection=="Еженедельные стоки ТТ":

    try:
        #--------------------------
        print("Начали обработку")
        #--------------------------
        sales=pd.read_excel(user_path,index_col=None,header=0)
        header_row=sales.index[sales.iloc[:,0] == 'Наименование ТТ'].tolist()
        header_row=header_row[0]
        header=sales.iloc[header_row]
        all_rows=header_row+1
        sales= sales[all_rows:]
        sales.columns=header
        sales=sales[sales['Наименование ТТ'].notnull()]
        print("Прочитали содержимое файла")
        print("----")
        #ADD GEO TABLE-----------------------------------------------------------
        geo=sales[['Наименование ТТ','РЦ','Филиал','Формат']]
        
        geo=geo.drop_duplicates()
        geo=geo[geo['Наименование ТТ']!="Grand Total"]
        geo=geo[geo['Наименование ТТ']!="Общий Итог"]
        geo=geo[geo['Наименование ТТ']!="Общий итог"]
        sales=sales.drop(['РЦ','Филиал','Формат'],axis=1)
        #ADD GEO TABLE-----------------------------------------------------------
        sales=pd.melt(sales,id_vars='Наименование ТТ',var_name="SKU", value_name='stock')
        sales=sales[sales.stock.notnull()]
        sales=sales[sales['Наименование ТТ']!="Grand Total"]
        sales=sales[sales['Наименование ТТ']!="Общий Итог"]
        sales=sales[sales['Наименование ТТ']!="Общий итог"]
        sales['SKU']=sales['SKU'].str.lower()
        user_path=user_path.replace(".xlsm","")
        week = int(user_path[-2:])
        year = 2022
        date = datetime.date(year, 1, 1) + relativedelta(weeks=+week)
        
        sales["DATE"]=date
        sales["week"]=week
        sales[['mrp1','mrp2']]=sales.SKU.str.split('мрц',expand=True)
        sales['mrp2']=sales['mrp2'].str.replace(" ","")
        sales['mrp2']=sales['mrp2'].str.slice(stop=3)
        sales['mrp2']=sales.mrp2.str.extract('(\d+)')
        sales['mrp2']=sales.mrp2.fillna(0)
        sales=sales.rename(columns={"mrp2": "MRP"})
        sales=sales.drop(['mrp1'],axis=1)
        #--------------------------
        print("Транспонировали и присвоили дату")
        print("----")
        #--------------------------
        print("Присвоили дату: ",date)
        print("----")
        print("Добавили МРЦ")
        print("----")
        print("SKU без МРЦ : ",sales[sales["MRP"]==0].SKU.unique())
        print("----")
        print("Обработали ",user_path)
        print("----")
        print("так выглядит обработанный массив")
        print("----")
        
        print(sales.head(5).to_string(index=False))
        
        print("----")
        print(user_path[-2:]," Сумма без учета блоков: ",sales.stock.sum())
        print("--------------------------------------")
        
    except:
        pass
        print(user_path,"НЕ ПОЛУЧИЛОСЬ")
    
    frame=sales
    frame = frame.astype({"Наименование ТТ": str})
    wb = xw.Book(result_file_geo)
    ws = wb.sheets["GEO Weekly ST"]
    ws.clear()
    ws["A1"].options(pd.DataFrame, header=1, index=False, expand='table').value = geo
    tbl_range = ws.range("A1").expand('table')
    ws.api.ListObjects.Add(1, ws.api.Range(tbl_range.address))
    ws.tables[0].name = "Weekly_ST"
    
    wb.save()
    wb.close()
    
    #--------------------------
    print("Теперь соединим файл с исходным")
    #current table
    current_df=pd.read_csv(main_filepath , delimiter="\t")
    current_df = current_df.astype({"Наименование ТТ": str})
    print("Прочитали Файл")
    print("----")
    new_total=pd.concat([frame,current_df])
    new_total=new_total[new_total['SKU'].notnull()]
    print("Добавили файл к исходному")
    print("----")
    #--------------------------
    new_total.to_csv(result_file+"\\Current_weekly_Stock.txt", index=None, sep='\t', mode='w+')
    print("Сохранили файл")
    #--------------------------
    print("Можно закрывать программу")
    #--------------------------
elif selection=="Еженедельные Продажи":
    
    try:
        #--------------------------
        print("Начали обработку")
        #--------------------------
        sales=pd.read_excel(user_path,index_col=None,header=0)
        header_row=sales.index[sales.iloc[:,0] == 'Магазин'].tolist()
        header_row=header_row[0]
        header=sales.iloc[header_row]
        all_rows=header_row+1
        sales= sales[all_rows:]
        sales.columns=header
        sales=sales[sales['Магазин'].notnull()]
        print("Прочитали содержимое файла")
        print("----")
        #ADD GEO TABLE-----------------------------------------------------------
        #------------------------------------------------------------------------
        geo=sales[['Магазин','FRMT','Филиал','РЦ (ОС)']]
        geo=geo.rename(columns={"Магазин": "Наименование ТТ", "FRMT":"Формат","РЦ (ОС)":"РЦ"})
        geo=geo.drop_duplicates()
        geo=geo[geo['Наименование ТТ']!="Grand Total"]
        geo=geo[geo['Наименование ТТ']!="Общий Итог"]
        geo=geo[geo['Наименование ТТ']!="Общий итог"]

        sales=sales.drop(['FRMT','Филиал','РЦ (ОС)'],axis=1)
        #------------------------------------------------------------------------
        #ADD GEO TABLE-----------------------------------------------------------
        sales=pd.melt(sales,id_vars='Магазин',var_name="SKU", value_name='sales')
        sales=sales[sales.sales.notnull()]
        sales=sales[sales['Магазин']!="Grand Total"]
        sales=sales[sales['Магазин']!="Общий Итог"]
        sales=sales[sales['Магазин']!="Общий итог"]
        sales['SKU']=sales['SKU'].str.lower()
        user_path=user_path.replace(".xlsm","")
        
        week = int(user_path[-2:])
        year = 2022
        date = datetime.date(year, 1, 1) + relativedelta(weeks=+week)
        
        sales["DATE"]=date
        sales["week"]=week
        sales[['mrp1','mrp2']]=sales.SKU.str.split('мрц',expand=True)
        sales['mrp2']=sales['mrp2'].str.replace(" ","")
        sales['mrp2']=sales['mrp2'].str.slice(stop=3)
        sales['mrp2']=sales.mrp2.str.extract('(\d+)')
        sales['mrp2']=sales.mrp2.fillna(0)
        sales=sales.rename(columns={"mrp2": "MRP"})
        sales=sales.drop(['mrp1'],axis=1)
        #--------------------------
        print("Транспонировали и присвоили дату")
        print("----")
        #--------------------------
        print("Присвоили дату: ",date)
        print("----")
        
        print("Добавили МРЦ")
        print("----")
        print(sales[sales["MRP"]==0].SKU.unique()," - SKU без МРЦ")
        print("----")
        print("Обработали ",user_path)
        
        print("----")
        print(user_path[-2:]," Сумма без учета блоков: ",sales.sales.sum())
        print("----")
        
    except:
        pass
        print(user_path,"НЕ ПОЛУЧИЛОСЬ")
    frame=sales
    frame = frame.astype({"Магазин": str})
    wb = xw.Book(result_file_geo)
    ws = wb.sheets["GEO Weekly Sales"]
    ws.clear()
    ws["A1"].options(pd.DataFrame, header=1, index=False, expand='table').value = geo
    tbl_range = ws.range("A1").expand('table')
    ws.api.ListObjects.Add(1, ws.api.Range(tbl_range.address))
    ws.tables[0].name = "Weekly_Sales"
    
    wb.save()
    wb.close()
    #--------------------------
    print("Теперь соединим файл с исходным")
    #current table
    
    current_df=pd.read_csv(main_filepath , delimiter="\t")
    current_df = current_df.astype({"Магазин": str})
    print("Прочитали Файл")
    new_total=pd.concat([frame,current_df])
    new_total=new_total[new_total['SKU'].notnull()]
    print("Добавили файл к исходному")
    #--------------------------
    new_total.to_csv(result_file+"\\Current_weekly_Sales.txt", index=None, sep='\t', mode='w+')
    print("Сохранили файл")
    
    #--------------------------
    #--------------------------
    
