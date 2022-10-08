
import pandas as pd
import re
import glob

import PySimpleGUI as sg

sg.theme('Dark Blue')

layout = [
    
    [sg.T("Input Storage Folder:", s=25,justification="r"), sg.I(key="-FOLDER-"), sg.FolderBrowse()],
    [sg.Text('Specify Sheet Name', size =(25, 1),justification="r"),sg.I(key="-IN-")],
    [sg.Text('Specify File Type in lowercase', size =(25, 1),justification="r"),sg.I(key="-FILE-")], 
    [sg.Submit()],
    [sg.Text('© JTI', size =(80, 1),justification="r")]
    
]

window = sg.Window('File Merger (xlsx/xlsm)', layout)

event, values = window.read()


result_file=str(values['-FOLDER-'])
sheet=str(values['-IN-'])
file=str(values['-FILE-'])
window.close()

result_file=result_file.replace("\\","\\\\")



all_files=glob.glob(result_file+"/*."+file)
li=[]


for filename in all_files:
    try:
        print("Reading ",filename)
        sales=pd.read_excel(filename,sheet_name=sheet )
        print("Reading ",sheet)
        filename=filename.replace("."+file,"")
        filename=filename.replace(result_file,"")
        
        sales["DATE"]=filename[-4:]
        
        alt=re.findall(r'\d+', filename)
        alt="-".join(alt)
        print("Extracting Date")
        sales['ALT_Date']=alt
        print("Given Date: ",alt)
        li.append(sales)
        print("Appending File")
        print("--------------")
    except:
        print(filename,"ERROR, smth went wrong")
print("Appending all dataframes")
frame=pd.concat(li,axis=0,ignore_index=True)
print("Total Rows Amount: ",frame.shape[0])
if frame.shape[0]>900000:
    frame.to_csv(result_file+"\\Merged_files.txt", index=None, sep='\t', mode='w+')
else:
    frame.to_csv(result_file+"\\Merged_files.txt", index=None, sep='\t', mode='w+')
    frame.to_excel(result_file+"\\Merged_files.xlsx", index=False)
print("Done")
