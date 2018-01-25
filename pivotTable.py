import pandas as pd
import numpy as np
from openpyxl import load_workbook


df = pd.read_excel('datax.xlsx')
liste=[(1,2),(2,2),(1,2)]
xx= pd.DataFrame(liste, columns=['kod', '2'])

#table = pd.pivot_table(xx, values=['2'], index=['kod'], aggfunc=np.sum)



table = pd.pivot_table(df, values=['cins','model'], index=['SAG','tasarım'], aggfunc={'cins' : 'count', 'model' : np.sum})
print(table)
#print (table.values)

table.columns.name = None               #remove categories
table = table.reset_index()                #index to columns

df3=table.as_matrix()
liste= df3.tolist()
print(liste)


path=r"boom.xlsx"
book = load_workbook(path)
writer = pd.ExcelWriter(path, engine = 'openpyxl')
writer.book = book
writer.sheets = dict((ws.title, ws) for ws in book.worksheets)
table.to_excel(writer, "shi", index=['cins', 'model'])

writer.save()
writer.close()
