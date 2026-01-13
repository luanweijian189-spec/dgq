import logging
from logging.handlers import TimedRotatingFileHandler
import os
import json
import sys
import shutil

def read_file_to_list(file_path, encoding='utf-8', strip_whitespace=True):
    try:
        with open(file_path, 'r', encoding=encoding) as file:
            if strip_whitespace:
                lines = [line.strip() for line in file.readlines()]
            else:
                lines = [line.rstrip('\n') for line in file]
        return lines
    except FileNotFoundError:
        return []
    except Exception as e:
        return []

def save_list_to_file(data_list, file_path, encoding='utf-8', add_newline=True):
    try:
        with open(file_path, 'w', encoding=encoding) as file:
            for item in data_list:
                if add_newline:
                    file.write(str(item) + '\n')
                else:
                    file.write(str(item))
        return True
    except Exception as e:
        return False

def get_image_files(directory):
    # 定义图片文件的扩展名
    image_extensions = ('.jpg', '.jpeg', '.png', '.gif', '.bmp', '.tiff')
    # 获取目录下的所有文件
    files = os.listdir(directory)
    # 筛选出图片文件
    image_files = [f for f in files if f.lower().endswith(image_extensions)]
    # 返回图片文件的完整路径
    image_files = [os.path.join(directory, f) for f in image_files]
    return image_files

def get_exe_dir():
    # If the script is bundled into an executable (e.g., via PyInstaller),
    # return the directory of the executable. Otherwise return the script directory.
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))

def save_dict_to_file(data: dict, filename: str):
    with open(filename, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=4)
def load_dict_from_file(filename: str) -> dict:
    if not os.path.exists(filename):
        return {}
    try:
        with open(filename, 'r', encoding='utf-8') as f:
            content = f.read().strip()
            return json.loads(content) if content else {}  # 处理空文件情况
    except json.JSONDecodeError:
        return {}

def backup_file(source_file: str):
	if not os.path.exists(source_file):
		return
	file_name, file_ext = os.path.splitext(source_file)
	backup_file = f"{file_name}_backup{file_ext}"  # 添加 _backup 后缀
	shutil.copyfile(source_file, backup_file)

logger = None
def init_log():
    global logger
    try:
        from colorama import init, Fore, Style
        init(autoreset=True)
        COLORAMA = True
    except ImportError:
        COLORAMA = False
    class ColorConsoleHandler(logging.StreamHandler):
        def emit(self, record):
            msg = self.format(record)
            if COLORAMA:
                if record.levelno == logging.INFO:
                    msg = Fore.GREEN + msg + Style.RESET_ALL
                elif record.levelno == logging.WARNING:
                    msg = Fore.YELLOW + msg + Style.RESET_ALL
                elif record.levelno == logging.ERROR:
                    msg = Fore.RED + msg + Style.RESET_ALL
            self.stream.write(msg + self.terminator)
            self.flush()
    # 创建logs目录
    if not os.path.exists("logs"):
        os.makedirs("logs")
    # 创建 logger
    logger = logging.getLogger()
    logger.setLevel(logging.INFO)
    # 创建控制台输出 handler
    console_handler = ColorConsoleHandler(sys.stdout)
    # 创建文件输出 handler（按天分割文件）
    log_file = os.path.join("logs", 'app.log')
    file_handler = TimedRotatingFileHandler(log_file, when='midnight', interval=1, backupCount=14)
    file_handler.setLevel(logging.DEBUG)  # 设置文件日志的级别
    file_handler.suffix = "%Y-%m-%d.log"  # 文件名后缀格式（日期格式）
    # 创建日志格式
    formatter = logging.Formatter('%(asctime)s - %(message)s')
    console_handler.setFormatter(formatter)
    file_handler.setFormatter(formatter)
    # 添加 handler 到 logger
    logger.addHandler(console_handler)
    logger.addHandler(file_handler)

def log(message):
    logger.info(message)

def log_error(message):
    logger.error(message)

def log_warning(message):
    logger.warning(message)