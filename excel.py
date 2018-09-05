# encoding:utf-8
import sys
import io
import os
import time
from selenium import webdriver

from xlrd import open_workbook

# 系统编码格式
reload(sys)
sys.setdefaultencoding('utf-8')

# 默认下载地址
defaultPath = os.getcwd() + '\\'
handleFile  = defaultPath + time.strftime("%Y-%m-%d", time.localtime()) + '.order.xls'

# 全局变量
classification_of_the_scenic_filename      = './景区分类.txt'.decode('utf-8').encode('gbk')
classification_of_the_scenic               = {}
supplier_partial_invoices_provide_filename = './部分开票供应商.txt'.decode('utf-8').encode('gbk')
supplier_partial_invoices_provide          = {}
supplier_all_invoices_provide_filename     = './全开票供应商.txt'.decode('utf-8').encode('gbk')
supplier_all_invoices_provide              = []
object_index                               = ''
rows_counts = 0
cols_counts = 0
excel_sum_list = []


def AnalysisExcel():
	print('【###】OpenExcelFile：'.decode('utf-8').encode('gbk') + handleFile)
	# 打开Excel文件并读取内容
	data=open_workbook(handleFile)
	# 选择第一张工作表
	sheet=data.sheets()[0]
	# 工作表总行数
	rows_counts = sheet.nrows
	# 每一行的值
	row_values = sheet.row_values
	# 工作表总列数
	cols_counts = sheet.ncols
	# 每一列的值
	col_values = sheet.col_values
	print(rows_counts)
	print(cols_counts)
	# 统计总和，循环变量需要-1
	for i in range(0, rows_counts-1):
		excel_sum_list.append({})
		for j in range(0, cols_counts):
			excel_sum_list[i][col_values(j)[0]] = row_values(i+1)[j]
		print(excel_sum_list[i])





def GetSupplierAllInvoicesProvide():
	print('【###】GetSupplierAllInvoicesProvide'.decode('utf-8').encode('gbk'))
	with io.open(supplier_all_invoices_provide_filename, 'r') as f:
		for line in f.readlines():
			line = line.strip("\r\n")
			if line.strip()=='':
				continue
			else :
				supplier_all_invoices_provide.append(line)
	print supplier_all_invoices_provide
	AnalysisExcel()

def GetSupplierPartialInvoicesProvide():
	print('【###】GetSupplierPartialInvoicesProvide'.decode('utf-8').encode('gbk'))
	with io.open(supplier_partial_invoices_provide_filename, 'r') as f:
		for line in f.readlines():
			line = line.strip("\r\n")
			if line.strip()=='':
				continue
			elif line[0] == "#":
				object_index = line[1:-1]
				supplier_partial_invoices_provide[object_index] = []
			else :
				supplier_partial_invoices_provide[object_index].append(line.split('，'))
	print supplier_partial_invoices_provide
	GetSupplierAllInvoicesProvide()

def GetClassificationOfTheScenic():
	print('【###】GetClassificationOfTheScenic'.decode('utf-8').encode('gbk'))
	with io.open(classification_of_the_scenic_filename, 'r') as f:
		for line in f.readlines():
			line = line.strip("\r\n")
			if line.strip()=='':
				continue
			elif line[0] == "#":
				object_index = line[1:-1]
				classification_of_the_scenic[object_index] = []
			else :
				classification_of_the_scenic[object_index].append(line.split('，'))
	print classification_of_the_scenic
	GetSupplierPartialInvoicesProvide()

def main():
	GetClassificationOfTheScenic()

if __name__ == '__main__':
	main()