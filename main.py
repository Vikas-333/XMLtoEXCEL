# import packages

import xml.etree.ElementTree as ETree
import pandas as pd
import xlsxwriter
# load and parse the input file

Tree = ETree.parse('C:/Users/Vikas/OneDrive/Desktop/PRO 2/inputxmlFile.xml')
root = Tree.getroot()

A = []
for ele in root:
    B = {}
    for i in list(ele):
        B.update({i.tag: i.text})
        A.append(B)

df = pd.DataFrame(A)
df.drop_duplicates(keep='first', inplace=True)
df.reset_index(drop=True, inplace=True)
writer = pd.ExcelWriter('C:/Users/Vikas/OneDrive/Desktop/PRO 2/OUTPUT.xlsx', engine='xlswriter')

df.to_excel(writer, sheet_name='sheet1')
worksheet = writer.sheets['sheet1']
worksheet.set_colmun('B:Z', 30)
writer.save('C:/Users/Vikas/OneDrive/Desktop/PRO 2/OUTPUT.xlsx')

print("XML file converted into Excel success")