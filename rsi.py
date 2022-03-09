from urllib.request import urlopen, Request
from bs4 import BeautifulSoup
from openpyxl import load_workbook
import ssl
import pandas as pd

def scrap(empresa):
	ctx = ssl.create_default_context()
	ctx.check_hostname = False
	ctx.verify_mode = ssl.CERT_NONE

	tds = []
	print(empresa)
	try:
		req = Request('https://finviz.com/quote.ashx?t='+empresa, headers={'User-Agent': 'Mozilla/5.0'})
		html = urlopen(req, context=ctx).read()
		soup = BeautifulSoup(html, "html.parser")
		divs = soup.findAll("table", {"class": "snapshot-table2"})
		for div in divs:
			rows = div.findAll('tr')
			for row in rows:
				tds.append(row.findAll('td'))
		rsi_value = str(tds[8][9].getText())
		print(rsi_value)
	except: 
		rsi_value = '0.0'
	return rsi_value

def readExcel(cols):
	excel_name = 'rsi_out.xlsx'
	sheets = ['Empresas']    
	list_of_dfs = []

	for sheet in sheets:
		df = pd.read_excel(excel_name, sheet_name=sheet, skiprows=1, usecols=cols)
		list_of_dfs.append(df)

	df_final = pd.concat(list_of_dfs, ignore_index=True)
	df_final.columns = ['Empresa']
	df_final['RSI'] = df_final.apply(lambda row : scrap(row['Empresa']), axis = 1)
	df_final = df_final.sort_values(by=['RSI'], ascending = False)
	df_final['RSI'] = df_final['RSI'].astype(float)
	df_final = df_final[df_final.RSI != 0]
	
	print(df_final)
	return df_final

def toExcel(df_out, fst_col):
	excel_name = 'rsi_out.xlsx'
	book = load_workbook(excel_name)
	writer = pd.ExcelWriter(excel_name, engine='openpyxl') 
	writer.book = book
	writer.sheets = dict((ws.title, ws) for ws in book.worksheets)
	df_out.to_excel(writer, sheet_name='TABLON', startrow=1, startcol=fst_col, index=False)
	writer.save()


### Start
cols_list = ['B','E','H','K','N']
cont = 1
for cols in cols_list:
	df_final = readExcel(cols)
	toExcel(df_final, cont)
	cont+=3
	print('-----------------')
