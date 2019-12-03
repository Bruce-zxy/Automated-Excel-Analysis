# encoding:utf-8
import sys
import os
import io
import operator
from time import strftime, localtime, sleep
from selenium import webdriver
from xlrd import open_workbook
from functools import reduce

# 系统编码格式
reload(sys)
sys.setdefaultencoding('utf-8')

YY = str(int(strftime("%Y", localtime())))
mm = str(int(strftime("%m", localtime())))
dd = str(int(strftime("%d", localtime())) - 1)

# 默认下载地址
defaultPath = os.getcwd() + '\\'
handleFile  = defaultPath + (YY+'-'+mm+'-'+dd) + '.order.xls'

# 全局变量
classification_of_the_scenic_filename      = defaultPath + '景区分类.txt'.decode('utf-8').encode('gbk')
classification_of_the_scenic               = {}
supplier_partial_invoices_provide_filename = defaultPath + '部分开票供应商.txt'.decode('utf-8').encode('gbk')
supplier_partial_invoices_provide          = {}
supplier_all_invoices_provide_filename     = defaultPath + '全开票供应商.txt'.decode('utf-8').encode('gbk')
supplier_all_invoices_provide              = []
object_index                               = ''

def AnalysisExcel(browser):
	print('【...】正在打开Excel文件：'.decode('utf-8').encode('gbk') + handleFile)

	if not os.path.exists(handleFile):
		sleep(5)
		print('【!!!】Excel文件不存在，可能正在下载中，请稍候...'.decode('utf-8').encode('gbk'))
		return AnalysisExcel(browser)
	else:
		browser.quit()

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
		elif '云锦庄' in curr_val[productName_index]:
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
			for j in xrange(0, len(item)):
				for k in xrange(0,len(item[j])):
					if item[j][k] in curr_val[productName_index]:
						result = result and True
					else :
						result = result and False
			if result:
				sales_invoice_sum += curr_val[salesAmount_index]

		# 景区统计
		for kind in classification_of_the_scenic:
			for scenic in classification_of_the_scenic[kind]:
				flag = True
				name = reduce(operator.add, scenic)
				for subname in scenic:
					if subname in curr_val[productName_index]:
						flag = flag and True
					else :
						flag = flag and False
				if flag and excel_scenic_sum[kind].has_key(name):
					excel_scenic_sum[kind][name] += int(curr_val[orderPerson_index])
				elif flag:
					excel_scenic_sum[kind][name] = int(curr_val[orderPerson_index])


	with open(defaultPath + '【每日销售情况汇报-'.decode('utf-8').encode('gbk') + (mm+'.'+dd) + '】.txt'.decode('utf-8').encode('gbk'), 'w') as f:
		f.write('【每日销售情况汇报-'.decode() + (mm+'.'+dd) + '】\n\n'.decode())
		f.write('一、景区情况：\n'.decode())
		f.write('订单人数：'.decode() + str(int(order_person_sum)) + '张\n'.decode())
		f.write('流水金额：'.decode() + str(int(sales_amount_sum)) + '元\n'.decode())
		f.write('实际营收：'.decode() + str(int(sales_invoice_sum)) + '元（可开发票）\n'.decode())
		serial = 1
		text = ''
		for kind in excel_scenic_sum:
			text += '\n'.decode() + str(serial) + '.'.decode() + kind + '销量：\n'.decode()
			serial += 1
			for scenic_name in excel_scenic_sum[kind]:
				if scenic_name == '丫山':
					text += scenic_name + '助销客销售中、'.decode()
				else :
					text += scenic_name + str(excel_scenic_sum[kind][scenic_name]) + '张、'.decode()
			text = text[:-1] + '\n'
		f.write(text)
	# os.remove(handleFile)

def GetSupplierAllInvoicesProvide(browser):
	print('【...】正在获取能全部开发票的供应商列表...'.decode('utf-8').encode('gbk'))
	with io.open(supplier_all_invoices_provide_filename, 'r') as f:
		for line in f.readlines():
			line = line.strip("\r\n")
			if line.strip()=='':
				continue
			else :
				supplier_all_invoices_provide.append(line)
	AnalysisExcel(browser)

def GetSupplierPartialInvoicesProvide(browser):
	print('【...】正在获取能部分开发票的供应商列表...'.decode('utf-8').encode('gbk'))
	with io.open(supplier_partial_invoices_provide_filename, 'r') as f:
		for line in f.readlines():
			line = line.strip("\r\n")
			if line.strip()=='':
				continue
			elif line[0] == "#":
				object_index = line[1:]
				supplier_partial_invoices_provide[object_index] = []
			else :
				supplier_partial_invoices_provide[object_index].append(line.split('，'))
	GetSupplierAllInvoicesProvide(browser)

def GetClassificationOfTheScenic(browser):
	print('【...】正在获取所有景区的分类...'.decode('utf-8').encode('gbk'))
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
	GetSupplierPartialInvoicesProvide(browser)


def startWebdriver(uname, upass):
	print("【###】您在config.txt中配置的【用户名】为：".decode('UTF-8').encode('GBK') + uname)
	print("【###】您在config.txt中配置的【密  码】为：".decode('UTF-8').encode('GBK') + upass)
	# 配置浏览器驱动
	options = webdriver.ChromeOptions()
	# options.set_headless()
	options.add_experimental_option('prefs', {'profile.default_content_settings.popups': 0, 'download.default_directory': defaultPath})

	# 启动浏览器驱动
	browser = webdriver.Chrome(executable_path=defaultPath + 'chromedriver.exe', chrome_options=options)
	# 打开网页
	browser.get('http://www.goyoto.com.cn/')

	# 输入用户名
	username = browser.find_element_by_name('username')
	username.clear()
	username.send_keys(uname)

	# 输入密码
	password = browser.find_element_by_name('password')
	password.clear()
	password.send_keys(upass)

	# 点击登录
	browser.find_element_by_css_selector('.sub_btn input').submit()

	# 等待跳转
	browser.implicitly_wait(3)
	# 解决弹窗
	browser.switch_to_alert().accept();

	# 切换frame
	browser.switch_to_frame(browser.find_element_by_name('fracmd'))
	browser.find_element_by_id('12').find_element_by_tag_name('a').click()

	# 切换回主内容`
	browser.switch_to_default_content()

	# 切换frame
	browser.switch_to_frame(browser.find_element_by_name('main'))
	# browser.find_elements_by_css_selector('#day_span a')[1].click()
	sdate = browser.find_element_by_id('sdate')
	sdate.clear()
	sdate.send_keys(YY+'-'+mm+'-'+dd)

	edate = browser.find_element_by_id('edate')
	edate.clear()
	edate.send_keys(YY+'-'+mm+'-'+dd)
	browser.find_element_by_name('export').click()

	# 将导出的表进行分析
	GetClassificationOfTheScenic(browser)
	
def main():
	if os.path.exists(handleFile):
		try: 
			os.remove(handleFile)
		except Exception as e:
			print('【!!!】请检查下列文件是否被其他程序占用！'.decode('utf-8').encode('gbk') + handleFile)
			return
	try:
		userInfo = []
		f=open(defaultPath + 'config.txt', 'r')
		for line in f.readlines():
			userInfo.append(line.strip('\n'))
		f.close()
		startWebdriver(userInfo[0], userInfo[1])
	except Exception as e:
		print(e)
		exit()
	finally:
		if f:
			f.close()

if __name__ == '__main__':
	main()