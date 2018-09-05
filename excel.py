# encoding:utf-8
import sys
import io
import os
import time
from selenium import webdriver
from xlrd import open_workbook
import operator
from functools import reduce

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

def AnalysisExcel():
	print('【###】OpenExcelFile：'.decode('utf-8').encode('gbk') + handleFile)

	# 总行列数
	rows_counts = 0
	cols_counts = 0

	# 各个字段的下标
	supplier_index = 29
	distributors_index = 26
	orderState_index = 21
	salesAmount_index = 20
	orderPerson_index = 18
	productName_index = 17

	# 各个数据统计结果
	excel_scenic_sum = {}
	sales_amount_sum = 0
	order_person_sum = 0
	sales_invoice_sum = 0

	# 打开Excel文件并读取内容
	data=open_workbook(handleFile)
	# 选择第一张工作表
	sheet=data.sheets()[0]
	# 工作表总行数
	rows_counts = sheet.nrows
	# 每一行的值的集合
	row_values = sheet.row_values

	for kind in classification_of_the_scenic:
		excel_scenic_sum[kind] = {}
		if kind != '其他景区':
			for scenic in classification_of_the_scenic[kind]:
				excel_scenic_sum[kind][reduce(operator.add, scenic)] = 0

	# 统计总和，循环变量从1到rows_counts -> (注意列表下标不要out of range)
	for i in range(1, rows_counts):

		curr_val = row_values(i)

		# 排除测试单、自供单以及未完成订单
		if '江西旅游科技' in curr_val[supplier_index]:
			continue
		elif '测试' in curr_val[distributors_index]:
			continue
		elif '已取消' in curr_val[orderState_index] or '待支付' in curr_val[orderState_index]:
			continue

		# 统计人数和流水金额
		sales_amount_sum += curr_val[salesAmount_index]
		order_person_sum += curr_val[orderPerson_index]

		# 统计实际营收
		if curr_val[supplier_index] in supplier_all_invoices_provide:
			sales_invoice_sum += curr_val[salesAmount_index]
		elif curr_val[supplier_index] in supplier_partial_invoices_provide.keys():
			item = supplier_partial_invoices_provide[curr_val[supplier_index]]
			result = True
			for j in xrange(0, len(item)-1):
				if item[j] in curr_val[productName_index]:
					result = result and True
				else :
					result = result and False
			if result:
				sales_invoice_sum += curr_val[salesAmount_index]

		# 景区统计
		for kind in classification_of_the_scenic:
			for scenic in classification_of_the_scenic[kind]:
				flag = True
				print(scenic)
				name = reduce(operator.add, scenic)
				# print(name.encode('gbk'))
				for subname in scenic:
					if subname in curr_val[productName_index]:
						flag = flag and True
					else :
						flag = flag and False
				if flag and excel_scenic_sum[kind].has_key(name):

					excel_scenic_sum[kind][name] += int(curr_val[orderPerson_index])
					print('!!--', excel_scenic_sum[kind][name])
					# if name == '灵山':
					# 	print('灵山'.encode('gbk'), excel_scenic_sum[kind][name])
				elif flag:
					excel_scenic_sum[kind][name] = int(curr_val[orderPerson_index])
					print(excel_scenic_sum[kind][name])


	print('流水金额：'.decode('utf-8').encode('gbk'))
	print(sales_amount_sum)
	print('订单人数：'.decode('utf-8').encode('gbk'))
	print(int(order_person_sum))
	print('实际营收：'.decode('utf-8').encode('gbk'))
	print(sales_invoice_sum)

	print(excel_scenic_sum)

	with open('./今日景区汇总.txt', 'w') as f:
		f.write('景区情况：\n'.decode())
		f.write('订单人数：'.decode() + str(int(order_person_sum)) + '张\n'.decode())
		f.write('流水金额：'.decode() + sales_amount_sum + '元\n'.decode())
		f.write('实际营收：'.decode() + sales_invoice_sum + '元（可开发票）\n'.decode())

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