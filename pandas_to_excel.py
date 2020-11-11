import pandas as pd
writer = pd.ExcelWriter('C:\\Users\\muhammedt\\Desktop\\demo.xlsx', engine='xlsxwriter')
writer.save()

# burada ilk dosya yaratılır.
Print("===================================================")



# dataframe Name and Age columns
df = pd.DataFrame({'Name': ['A', 'B', 'C', 'D'],
                   'Age': [10, 0, 30, 50]})

# Create a Pandas Excel writer using XlsxWriter as the engine.
writer = pd.ExcelWriter('C:\\Users\\muhammedt\\Desktop\\demo.xlsx', engine='xlsxwriter')

# Convert the dataframe to an XlsxWriter Excel object.
df.to_excel(writer, sheet_name='Sheet1', index=False)

# Close the Pandas Excel writer and output the Excel file.
writer.save()

# Burada dosyaya  dict veri yazılır.
Print("===================================================")


import pandas as pd
from openpyxl import load_workbook
# new dataframe with same columns
df = pd.DataFrame({'Name': ['E','F','G','H'],
                   'Age': [1000,70,40,60]})
writer = pd.ExcelWriter('C:\\Users\\muhammedt\\Desktop\\demo.xlsx', engine='openpyxl')
# try to open an existing workbook
writer.book = load_workbook('C:\\Users\\muhammedt\\Desktop\\demo.xlsx')
# copy existing sheets
writer.sheets = dict((ws.title, ws) for ws in writer.book.worksheets)
# read existing file
reader = pd.read_excel(r'C:\\Users\\muhammedt\\Desktop\\demo.xlsx')
# write out the new sheet
df.to_excel(writer,sheet_name='Sheet1',index=False,header=False,startrow=len(reader)+1)

writer.close()


# Burada dosyaya veri ilave edilir dikkat engine= openpyxl olacak ve sheet name= aynı sheet

Print("===================================================")


"""
import pandas as pd

# import Pandas to take datas from excel file and make a dataframe

import xlrd

# import xlrd to read excel file "xlsx", for txt or csv you don't need xlrd

from pandas import ExcelWriter
from pandas import ExcelFile



import xlsxwriter 
# Create a Pandas Excel writer using XlsxWriter as the engine.
    
# Create a Pandas dataframe from some data.
  

# Create a Pandas Excel writer using XlsxWriter as the engine.
writer = pd.ExcelWriter('C:\\Users\\muhammedt\\Desktop\\pandas_CREATED_file_ruzgar.xlsx', engine='xlsxwriter')
#writer = pd.ExcelWriter('C:/Users/90534/Desktop/pandas_simple.xlsx', engine='xlsxwriter')

# Convert the dataframe to an XlsxWriter Excel object.

dfbahce.to_excel(writer, sheet_name='Bahce')
dfhasanbeyli.to_excel(writer, sheet_name='Hasanbeyli')  
dfnurdagi.to_excel(writer, sheet_name='Nurdagi') 


    # Close the Pandas Excel writer and output the Excel file.
writer.save()
    ################ Data Frame Ends ##################"""
    
# Burada data frame den veri xlswriter ile bir dosyaya yazdırılır ama üzerine ekleme değil sıfırdan yazdırma 
Print("===================================================")
