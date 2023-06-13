#XML TO EXCEL FILE
from fileinput import filename
import xml.etree.ElementTree as ET
from openpyxl import Workbook
import os 
def readFile(filename):
	'''
		Checks if file exists, parses the file and extracts the needed data
		returns a 2 dimensional list without "header"
	'''
	if not os.path.exists(filename): return
	tree = ET.parse(filename)
	root = tree.getroot()
	#you may need to adjust the keys based on your file structure
	dict_keys = ["id","first_name","last_name","email","gender","ip_address" ] #all keys to be extracted from xml
	mdlist = []
	for child in root:
		temp = []
		for key in dict_keys:
			temp.append(child.find(key).text)
		mdlist.append(temp)
	return mdlist

def to_Excel(mdlist):
	'''
		Generates excel file with given data
		mdlist: 2 Dimenusional list containing data
	'''
	wb = Workbook()
	ws = wb.active
	for i,row in enumerate(mdlist):
		for j,value in enumerate(row):
			ws.cell(row=i+1, column=j+1).value = value
	newfilename = os.path.abspath("./xml_to_excel.xlsx")
	wb.save(newfilename)
	print("complete")
	return

result = readFile(r"C:\Users\Vikas\OneDrive\Desktop\PRO 2\inputxmlFile.xml")
if result:
	to_Excel(result)