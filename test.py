# encoding:utf-8
import sys
import os
import time
import xlrd
from selenium import webdriver

# 系统编码格式
reload(sys)
sys.setdefaultencoding('utf-8')

# 默认下载地址
defaultPath = os.getcwd() + '\\'
handleFile  = defaultPath + time.strftime("%Y-%m-%d", time.localtime()) + '.order.xls'

def AnalysisExcel():
	while 1:
		if not os.path.exists(handleFile):
			time.sleep(3)
			print(handleFile)
		else :
			break
	data=xlrd.open_workbook(handleFile)
	sheet=data.sheets()[0]
	rows = sheet.row_values(3)
	print(rows[6].encode('gbk'))


	# os.remove(handleFile)

def startWebdriver(uname, upass):
	print("您的用户名：".decode('UTF-8').encode('GBK') + uname)
	print("您的密  码：".decode('UTF-8').encode('GBK') + upass)
	# 配置浏览器驱动
	options = webdriver.ChromeOptions()
	# options.set_headless()
	options.add_experimental_option('prefs', {'profile.default_content_settings.popups': 0, 'download.default_directory': defaultPath})

	# 启动浏览器驱动
	browser = webdriver.Chrome(executable_path='./chromedriver.exe', chrome_options=options)
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

	# 切换回主内容
	browser.switch_to_default_content()

	# 切换frame
	browser.switch_to_frame(browser.find_element_by_name('main'))
	browser.find_elements_by_css_selector('#day_span a')[0].click()
	browser.find_element_by_name('export').click()

	# 将导出的表进行分析
	AnalysisExcel()

	browser.quit()
	
def main():
	if os.path.exists(handleFile):
		os.remove(handleFile)
	try:
		userInfo = []
		f=open('./config.txt', 'r')
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