

import openpyxl


file_path = "C:\\Users\\m_asem\\Desktop\\Book1.xlsx"
workbook = openpyxl.load_workbook(file_path)


print("Sheets are:")
print(workbook.sheetnames)



import os

sheet=workbook['Sheet1']
sheet['A1']='Hello'
workbook.save(file_path)
#os.startfile(file_path)

import pandas as pd


data_table = pd.DataFrame({
    "ID": [1, 2, 3],
    "Name": ["John", "Jane", "Alex"],
    "Age": [28, 22, 32]
})

data_table.to_excel(file_path,sheet_name='Sheet1', startrow=8, startcol=2,index=False)
print(data_table)
workbook.close()
os.close(file_path)







