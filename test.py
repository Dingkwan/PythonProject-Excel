from openpyxl import load_workbook

book=load_workbook("下单案件统计20220114.xlsx")

sheet=book["Sheet5"]

max_rows=sheet.max_row				#表格的最大行数
max_cols=sheet.max_column			#表格的最大列数

def kehubianhao():
	for rows in range(1,max_rows):
		val=