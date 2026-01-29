import common

import configparser
import queue
import threading
import time
import openpyxl
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font
import re
from html import unescape
from DrissionPage import Chromium, ChromiumOptions
from DrissionPage import ChromiumPage, ChromiumOptions
from DrissionPage._pages.chromium_page import ChromiumPage
import traceback
import os
import json
from pathlib import Path
import shutil
import base64
import requests
from urllib.parse import urlparse

SUMMARY_PROMPT = (
	"你是一名科研助理。请阅读我上传的论文PDF并返回："
	"1) 中文摘要；2) 关键贡献点（3-5条）；3) 方法概述；"
	"4) 主要实验或结果；5) 局限性或未来工作。"
	"输出使用清晰的分段文本，不要使用JSON格式。"
)

def format_time_string(timestamp):
	local_time = time.localtime(timestamp)
	return time.strftime("%Y-%m-%d %H:%M:%S", local_time)

def wait_page_loading(page):
	common.log("等待页面加载完成...")
	time.sleep(2.5)
	while page.states.ready_state != 'complete':
		time.sleep(1)
	time.sleep(2.5)

# 模拟点击
def simulate_click(page, ele):
	if not ele.states.is_covered and ele.states.is_clickable and ele.states.has_rect:
		page.actions.move_to(ele).click()
	else:
		if ele.states.has_rect:
			ele.hover()
			time.sleep(0.1)
		if ele.states.is_covered:
			ele.click(True)
		else:
			ele.click()
# 模拟按下回车
def press_enter(tab):
	tab.actions.key_down('ENTER')
	time.sleep(0.5)
	tab.actions.key_up('ENTER')

def is_url(value):
	parsed = urlparse(value)
	return parsed.scheme in ('http', 'https') and parsed.netloc != ""

def sanitize_filename(filename):
	filename = unescape(filename)
	filename = re.sub(r'[^A-Za-z0-9._-]', '_', filename)
	return filename or "paper.pdf"

def download_pdf(url, output_dir):
	os.makedirs(output_dir, exist_ok=True)
	parsed = urlparse(url)
	filename = sanitize_filename(os.path.basename(parsed.path) or "paper.pdf")
	if not filename.lower().endswith(".pdf"):
		filename += ".pdf"
	file_path = os.path.join(output_dir, filename)
	response = requests.get(url, stream=True, timeout=60)
	response.raise_for_status()
	content_type = response.headers.get("Content-Type", "").lower()
	if "pdf" not in content_type and not filename.lower().endswith(".pdf"):
		raise ValueError("链接不是PDF文件")
	with open(file_path, "wb") as f:
		for chunk in response.iter_content(chunk_size=1024 * 1024):
			if chunk:
				f.write(chunk)
	return file_path

def request_ai_summary(prompt_content, file_path):
	with open(file_path, 'rb') as f:
		file_data = f.read()
	file_base64 = base64.b64encode(file_data).decode('utf-8')
	mime_type = mimeTypes.get(os.path.splitext(file_path)[1].lower(), "application/octet-stream")
	filename = os.path.basename(file_path)
	api_url = "https://api.gpt.ge/v1/chat/completions"
	payload = {
		"model": "gemini-2.5-pro",
		"messages": [{
			"role": "user",
			"content": [
				{
					"type": "text",
					"text": prompt_content
				},
				{
					"type": "file",
					"file": {
						"filename": filename,
						"file_data": f"data:{mime_type};base64,{file_base64}"
					}
				}
			]
		}],
		"max_tokens": 6000,
		"temperature": 0.5,
		"stream": False
	}
	response = requests.post(api_url, json=payload, headers=headers, timeout=120)
	response.raise_for_status()
	result = response.json()
	content_text = result.get('choices', [{}])[0].get('message', {}).get('content', "")
	if not content_text:
		raise ValueError("AI接口未返回内容")
	return content_text

def summarize_paper_from_url(url, prompt_content, output_dir):
	pdf_path = download_pdf(url, output_dir)
	summary_text = request_ai_summary(prompt_content, pdf_path)
	summary_path = f"{os.path.splitext(pdf_path)[0]}_summary.txt"
	with open(summary_path, 'w', encoding='utf-8') as f:
		f.write(summary_text)
	return summary_path

# 统一的excel处理入口
EXCEL_COLUMN_ORIGINAL_SPECS = 7
EXCEL_COLUMN_ORIGINAL_IMAGES = 8
EXCEL_COLUMN_STATUS = 9
processor_excel = None
def processor_excel_function(data):
	excel_data_content(data['row'], data['column'], data['content'])
### Excel 文件处理
excel_lock = threading.Lock()
# 店铺状态设置
def excel_data_content(row, column, content):
	data_file_path = system_config['File']
	with excel_lock:
		excel = openpyxl.load_workbook(data_file_path)
		sheet = excel['Sheet1']
		sheet.cell(row = row, column = column, value = content)
		excel.save(data_file_path)
		excel.close()
# 数据处理流程，统一线程
class DataProcessor:
	def __init__(self, process_function):
		self.data_queue = queue.Queue()
		self.processor_thread = None
		self.running = False
		self.process_function = process_function
	def start(self):
		self.running = True
		self.processor_thread = threading.Thread(target=self._process_loop)
		self.processor_thread.start()
	def stop(self):
		self.running = False
		if self.processor_thread:
			self.processor_thread.join(timeout=5)
	def add_data(self, data):
		self.data_queue.put(data)
	def _process_loop(self):
		while self.running:
			try:
				# 阻塞获取数据，最多等待1秒
				data = self.data_queue.get(timeout=1)
				self._process_data(data)
				self.data_queue.task_done()
			except queue.Empty:
				# 队列为空，继续循环
				continue
	def _process_data(self, data):
		try:
			self.process_function(data)
		except Exception as e:
			common.log(f"数据处理异常: {e}")


# 其他全局数据
headers = {
	'Content-Type': 'application/json',
	'Authorization': ''  # 将在init_config时设置
}
mimeTypes = {
	".pdf":  "application/pdf",
	".mp3":  "audio/mp3",
	".mp4":  "video/mp4",
	".wav":  "audio/wav",
	".png":  "image/png",
	".jpg":  "image/jpeg",
	".jpeg": "image/jpeg",
	".txt":  "text/plain",
	".mov":  "video/mov",
	".mpeg": "video/mpeg",
	".mpg":  "video/mpg",
	".avi":  "video/avi",
	".wmv":  "video/wmv",
	".flv":  "video/flv",
}

# 配置数据
system_config = {}
# 初始化配置
def init_config():
	global system_config
	# 读取配置
	config = configparser.ConfigParser()
	config.read('config.ini', encoding="utf-8")
	system_config = config['System']

def open_browser(port, using_port_directory = True):
	CO = ChromiumOptions()
	# 设置端口和用户数据目录
	if not using_port_directory:
		CO.set_user_data_path("user_data")
	else:
		CO.set_user_data_path(f"user_data_{port}")
	CO.set_local_port(port)
	CO.set_pref("credentials_enable_service", False)
	# 启动浏览器
	browser = Chromium(CO)
	common.log("浏览器已启动")
	return browser

def close_browser(browser):
	common.log("关闭浏览器")
	browser.quit()


# 线程处理
def thread_process(index):
	thread_no = index + 1
	port_base = int(system_config['Port'])
	port = port_base + index
	key = system_config['Key']
	prompt_file = system_config['Prompt_File']
	prompt_file_path = os.path.join(common.get_exe_dir(), prompt_file)

	prompt_content = ""
	# 读取prompt文件内容
	with open(prompt_file_path, 'r', encoding='utf-8') as f:
		prompt_content = f.read()
	# 设置Authorization header
	headers['Authorization'] = f"Bearer {system_config['Key']}"
	
	browser = open_browser(port)
	main_page = browser.latest_tab
	page = browser.new_tab()
	page.get("https://oa.deeproute.cn/")
	wait_page_loading(page)
	page.set.window.max()
	try:
		while True:
			common.log_warning(f'''请先确保进入CW06页面''')
			while True:
				common.log_warning(f'''请输入要登记流程的目录（点击右键粘贴），并按【回车键】继续：''')
				directory = input()
				if is_url(directory):
					common.log_warning(f'''检测到PDF链接，开始下载并总结论文：{directory}''')
					try:
						summary_path = summarize_paper_from_url(directory, SUMMARY_PROMPT, common.get_exe_dir())
						common.log(f'''论文总结已保存：{summary_path}''')
					except Exception as e:
						common.log_error(f"论文总结失败: {e}")
					continue
				if os.path.isdir(directory):
					break
				else:
					common.log_error(f'''输入的目录不存在，请重新输入！''')
			common.log(f'''开始处理目录：{directory}''')

			try:
				# 在directory目录下创建一个Excel文件
				result_excel_path = os.path.join(directory, "CW06费用报销数据.xlsx")
				if os.path.isfile(result_excel_path):
					common.log(f'''发现已存在的结果文件，先备份''')
					backup_excel_path = os.path.join(directory, f"CW06费用报销数据备份_{int(time.time())}.xlsx")
					shutil.move(result_excel_path, backup_excel_path)
				common.log(f'''创建结果文件：{result_excel_path}''')
				result_excel = Workbook()
				sheet = result_excel.active
				sheet.title = "明细2"
				# 写入表头
				sheet.cell(row = 1, column = 1, value = "序号")
				sheet.cell(row = 1, column = 2, value = "发生事由##fssy1")
				sheet.cell(row = 1, column = 3, value = "币种##bz1")
				sheet.cell(row = 1, column = 4, value = "申报金额##sbje1")
				sheet.cell(row = 1, column = 5, value = "发票张数##fpzs1")
				sheet.cell(row = 1, column = 6, value = "备注##bz2")
				sheet.cell(row = 1, column = 7, value = "开始日期##ksrq1")
				sheet.cell(row = 1, column = 8, value = "结束日期##jsrq1")
				sheet.cell(row = 1, column = 9, value = "出差城市##cccs1")
				sheet.cell(row = 1, column = 10, value = "费用类型##fylx1")
				sheet.cell(row = 1, column = 11, value = "项目名称##xmmc")
				result_excel.save(result_excel_path)
				result_excel.close()
				# 读取该目录下（仅当前目录）所有的pdf/jpg/jpeg/png文件
				files = os.listdir(directory)
				file_list = []
				for file in files:
					file_path = os.path.join(directory, file)
					if os.path.isfile(file_path):
						ext = os.path.splitext(file)[1].lower()
						if ext in ['.pdf', '.jpg', '.jpeg', '.png']:
							file_list.append(file_path)
				# 通过AI提取关键信息存储到Excel中
				reason = ""
				start_time = ""
				end_time = ""
				city = ""
				row_index = 2
				for file_path in file_list:
					try:
						common.log(f'''利用AI处理文件 {file_path}''')
						# 读取文件内容（二进制）
						file_base64 = ""
						with open(file_path, 'rb') as f:
							file_data = f.read()
							file_base64 = base64.b64encode(file_data).decode('utf-8')
						# 获取文件MIME类型
						mime_type = mimeTypes.get(os.path.splitext(file_path)[1].lower(), "application/octet-stream")
						# 调用API
						filename = os.path.basename(file_path)
						api_url = f"https://api.gpt.ge/v1/chat/completions"
						payload = {
							"model": "gemini-2.5-pro",
							"messages": [{
								"role": "user",
								"content": [
									{
										"type": "text",
										"text": prompt_content
									},
									{
										"type": "file",
										"file": {
											"filename": filename,
											"file_data": f"data:{mime_type};base64,{file_base64}"
										}
									}
								]
							}],
							"max_tokens": 6000,
							"temperature": 0.5,
							"stream": False
						}
						response = requests.post(api_url, json=payload, headers=headers)
						result = response.json()
						# 提取返回的JSON内容
						content_text = result['choices'][0]['message']['content']
						common.log(f'''AI接口返回内容: {content_text}''')
						json_match = re.search(r'\{.*\}', content_text, re.DOTALL)
						if json_match:
							data = json.loads(json_match.group())
							if filename.find("出差申请") != -1:
								# 出差申请单，提取开始日期、结束日期、出差城市
								start_time_gotten = data.get("开始日期", "")
								if start_time_gotten != "":
									start_time = start_time_gotten
								end_time_gotten = data.get("结束日期", "")
								if end_time_gotten != "":
									end_time = end_time_gotten
								city_gotten = data.get("出差城市", "")
								if city_gotten != "":
									city = city_gotten
								reason_gotten = data.get("发生事由", "")
								if reason_gotten != "":
									reason = reason_gotten
							amount_gotten = data.get("申报金额", "")
							if amount_gotten == "" or amount_gotten == None:
								continue
							# 写入Excel
							excel = None
							if os.path.exists(result_excel_path):
								excel = openpyxl.load_workbook(result_excel_path)
							else:
								excel = Workbook()
							sheet = excel.active
							sheet.cell(row=row_index, column=1, value=row_index - 1)
							sheet.cell(row=row_index, column=3, value=data.get("币种", ""))
							sheet.cell(row=row_index, column=4, value=data.get("申报金额", ""))
							sheet.cell(row=row_index, column=5, value=data.get("发票张数", ""))

							# 不录入“备注”
							sheet.cell(row=row_index, column=6, value="")

							sheet.cell(row=row_index, column=10, value=data.get("费用类型", ""))
							sheet.cell(row=row_index, column=11, value="零散杂项")
							excel.save(result_excel_path)
							excel.close()
							row_index += 1
					except Exception as e:
						common.log_error(f"处理文件 {file} 失败: {e}")
				# 写入发生事由
				if not os.path.exists(result_excel_path):
					common.log_error(f"未能从文件中解析生成Excel文件，忽略本次执行")
					continue
				excel = openpyxl.load_workbook(result_excel_path)
				sheet = excel.active
				for r in range(2, row_index):
					sheet.cell(row=r, column=7, value=start_time)
					sheet.cell(row=r, column=8, value=end_time)
					sheet.cell(row=r, column=9, value=city)
					sheet.cell(row=r, column=2, value=reason)
				excel.save(result_excel_path)
				excel.close()
				common.log(f"目录数据处理完成，结果已保存到: {result_excel_path}")
				# 开始操作页面
				latest_tab = browser.latest_tab
				common.log(f"点击“明细导入”")
				latest_tab.ele('xpath://span[text()="明细导入"]').click(True)
				time.sleep(1)
				ele_message_to_save_record = latest_tab.ele('xpath://div[text()="流程数据还未保存，现在保存吗？"]', timeout=3)
				if ele_message_to_save_record != None:
					ele_message_to_save_record.parent('css:div[role="dialog"]').ele('css:button').click(True)
					time.sleep(1)
					wait_page_loading(latest_tab)
				# 等待“明细导入”对话框出现
				common.log(f"等待“明细导入”对话框出现")
				ele_title = None
				while True:
					ele_title = latest_tab.ele('xpath://div[text()="明细导入"]', timeout=1)
					if ele_title != None:
						break
					time.sleep(1)
				ele_dialog = ele_title.parent('css:div[role="dialog"]')
				common.log(f"上传Excel到附件")
				ele_dialog.ele('xpath://span[text()="上传附件"]').parent("css:button").click.to_upload(result_excel_path)
				time.sleep(3)
				ele_dialog.ele('xpath://span[text()="明细导入"]').parent("css:button").click(True)
				time.sleep(1)
				ele_dialog_import_result = None
				while True:
					ele_title_import_result = latest_tab.ele('xpath://div[text()="导入结果"]', timeout=1)
					if ele_title_import_result != None:
						ele_dialog_import_result = ele_title_import_result.parent('css:div[role="dialog"]')
						break
					time.sleep(1)
				time.sleep(1)
				# 检查导入结果
				ele_error_tip = ele_dialog_import_result.ele('xpath://*[contains(text(), "导入失败")]', timeout=1)
				if ele_error_tip != None:
					common.log_error(f"导入失败，动作已暂停，请检查Excel内容是否正确！")
				else:
					common.log(f"导入成功，关闭对话框")
					ele_dialog_import_result.ele('xpath://span[text()="确 定"]').parent("css:button").click(True)
					time.sleep(1)
					wait_page_loading(latest_tab)
					# 填入“借款余额”
					common.log(f"填入“借款余额”为0")
					latest_tab.ele('css:div[data-fieldname="jkye"] input').input("0")
					time.sleep(1)
					# 填入“定额备用金金额”
					common.log(f"填入“定额备用金金额”为0")
					latest_tab.ele(f'css:div[data-fieldname="cxje"] input').input("0")
					# 上传各个附件
					common.log(f"上传各个附件")
					for file_path in file_list:
						latest_tab.ele('xpath://span[text()="上传附件"]').parent("css:button").click.to_upload(file_path)
						time.sleep(3)
					# 保存
					latest_tab.ele('css:button[title="保存"]').click(True)
					time.sleep(3)
					wait_page_loading(latest_tab)
					common.log(f"数据已保存完成")
			except Exception as e:
				common.log_error(traceback.format_exc())
				common.log_error(e)
				common.log_error("忽略本次执行")
	except Exception as e:
		common.log_error(traceback.format_exc())
		common.log_error(e)
	finally:
		close_browser(browser)

threads = []
def run():
	# 创建数据处理器
	global processor_excel
	processor_excel = DataProcessor(processor_excel_function)
	processor_excel.start()

	global threads
	thread_number = 1
	# 创建线程
	for i in range(thread_number):
		t = threading.Thread(target=thread_process, args=(i,))
		threads.append(t)
		t.start()
	# 等待所有线程执行完成
	for t in threads:
		t.join()

def main():

	# 初始化日志
	common.init_log()

	# 初始化配置
	init_config()

	# 运行
	run()

#######################################
if __name__ == "__main__":
	main()
