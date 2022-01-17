from pickle import FALSE
from types import NoneType
from openpyxl import load_workbook

book=load_workbook("下单案件统计20220114.xlsx")

sheet=book["Sheet5"]

max_rows=sheet.max_row				#表格的最大行数
max_cols=sheet.max_column			#表格的最大列数



def getOrderNo():
	orderList=[]
	for rows in range(1,max_rows+1):
		val=sheet["C%d" %rows].value
		if (type(val)==NoneType):
			continue
		orderList.append(val)
	orderList.remove("下单编号")
	return orderList

orderNo=getOrderNo()
print(orderNo)
