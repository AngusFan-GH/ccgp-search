import io
import math
import os
import sys
import threading
import time
import tkinter as tk
from datetime import datetime, timedelta
from tkinter import messagebox
from urllib.parse import quote

from bs4 import BeautifulSoup
from PIL import Image
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait

RETRY_TIMES = 60
driver = None


class ConsoleRedirector:
    def __init__(self, text_widget):
        self.text_widget = text_widget

    def write(self, str):
        self.text_widget.insert(tk.END, str)
        self.text_widget.see(tk.END)

    def flush(self):
        pass


def generate_params(start_time, end_time, keyword, page_index=1):
    start_time_encoded = quote(start_time, safe='')
    end_time_encoded = quote(end_time, safe='')
    keyword_encoded = quote(keyword, safe='')

    params = {
        'searchtype': '1',
        'page_index': page_index,
        'bidSort': '0',
        'buyerName': '',
        'projectId': '',
        'pinMu': '0',
        'bidType': '0',
        'dbselect': 'bidx',
        'kw': keyword_encoded,
        'start_time': start_time_encoded,
        'end_time': end_time_encoded,
        'timeType': '6',
        'displayZone': '',
        'zoneId': '',
        'pppStatus': '0',
        'agentName': ''
    }

    params_str = '&'.join([f"{key}={value}" for key, value in params.items()])
    return params_str


def init_driver():
    # 获取当前脚本所在的目录
    base_dir = os.path.dirname(os.path.abspath(__file__))
    # 拼接 chromedriver.exe 的路径
    driver_path = os.path.join(base_dir, 'chromedriver.exe')
    chrome_options = Options()
    chrome_options.add_argument(
        "--disable-blink-features=AutomationControlled")
    chrome_options.add_argument("--disable-extensions")
    chrome_options.add_argument("--start-maximized")
    chrome_options.add_experimental_option(
        "excludeSwitches", ["enable-automation"])
    chrome_options.add_experimental_option('useAutomationExtension', False)

    service = Service(driver_path)
    driver = webdriver.Chrome(service=service, options=chrome_options)
    return driver


def get_search_list(date_range, kw, driver, page_index, total_pages=None):
    if page_index == 1:
        print(
            f"正在搜索{date_range[0].replace(':', '-')}到{date_range[1].replace(':', '-')}的包含'{kw}'关键词的内容......")
    url = 'https://search.ccgp.gov.cn/bxsearch?' + \
        generate_params(date_range[0], date_range[1], kw, page_index)
    driver.get(url)
    time.sleep(1)
    html_content = driver.page_source
    soup = BeautifulSoup(html_content, 'html.parser')
    # 如果获取到“频繁”字样，说明被反爬虫系统拦截，等待一段时间后重试
    if '您的访问过于频繁,请稍后再试' in soup.get_text():
        global RETRY_TIMES
        print(f"搜索过于频繁，{RETRY_TIMES}秒后将自动重试。")
        time.sleep(RETRY_TIMES)
        RETRY_TIMES = 2 * RETRY_TIMES
        return get_search_list(date_range, kw, driver, page_index, total_pages)
    else:
        RETRY_TIMES = 10
    if page_index == 1 and total_pages is None:
        total_results = get_total_results(html_content)
        if total_results is None:
            print("搜索结果为空。")
        else:
            print(f"共找到{total_results}条记录。")

        # 总页数
        total_pages = math.ceil(int(total_results) / 20)

    if total_pages > 1:
        print(f"正在获取第{page_index}页的搜索结果......")

    li_tags = soup.select('.vT-srch-result-list .vT-srch-result-list-bid li')
    results = []
    for li in li_tags:
        a_tag = li.find('a')
        if a_tag:
            title = a_tag.get_text(strip=True)
            href = a_tag['href']
            results.append({'title': title, 'href': href})
    if page_index < total_pages:
        res = get_search_list(date_range, kw, driver,
                              page_index + 1, total_pages)
        results.extend(res)
    return results


def get_total_results(html_content):
    soup = BeautifulSoup(html_content, 'html.parser')
    try:
        total_results_tag = soup.select(
            '.vT_z div > p:nth-child(1) > span:nth-child(2)')[0]
        total_results = total_results_tag.get_text(strip=True)
        return int(total_results) if total_results else None
    except (IndexError, ValueError) as e:
        print(f"获取总结果数失败: {e}，请稍后重试。")
        return None


def filter_content(results, driver, filter_keyword):
    source_dir = 'screenshots'
    if not os.path.exists(source_dir):
        os.makedirs(source_dir)
    save_dir = os.path.join(source_dir, str(int(time.time())))
    os.makedirs(save_dir)

    print(f"正在筛选包含'{filter_keyword}'的记录......")

    filter_results = []
    for result in results:
        title = result['title']
        href = result['href']
        driver.get(href)
        time.sleep(1)
        html_content = driver.page_source
        soup = BeautifulSoup(html_content, 'html.parser')
        if filter_keyword in soup.get_text():
            print(f"存在‘{filter_keyword}’字样的页面: {title}。")
            print(f"链接: {href}。")
            filter_results.append(result)
            print(f"正在保存'{title}'的完整网页截图。")
            screenshot_filename = f"{title}.png"
            if os.path.exists(screenshot_filename):
                screenshot_filename = f"{title}_{time.time()}.png"

            # 等待页面加载完成
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.TAG_NAME, "body")))

            # 匹配页面中所有的filter_keyword，高亮显示，以便截图
            driver.execute_script(
                f"document.body.innerHTML = document.body.innerHTML.replace(new RegExp('{filter_keyword}', 'g'), '<span style=\"background-color: yellow;\">{filter_keyword}</span>');")

            # 获取完整网页的滚动高度
            total_height = driver.execute_script(
                "return document.body.scrollHeight")
            viewport_height = driver.execute_script(
                "return window.innerHeight")

            # 初始化图片
            screenshots = []
            for i in range(0, total_height, viewport_height):
                driver.execute_script(f"window.scrollTo(0, {i});")
                time.sleep(0.5)
                screenshot = driver.get_screenshot_as_png()
                screenshots.append(Image.open(io.BytesIO(screenshot)))

            # 拼接图片
            full_screenshot = Image.new(
                'RGB', (screenshots[0].width, total_height))
            offset = 0
            for screenshot in screenshots:
                full_screenshot.paste(screenshot, (0, offset))
                offset += screenshot.height

            # 保存完整网页截图到save_dir文件夹中
            full_screenshot.save(os.path.join(save_dir, screenshot_filename))
            print(f"已保存截图: {screenshot_filename}。")
        # 保存筛选结果到save_dir文件夹中
        with open(os.path.join(save_dir, '筛选结果.txt'), 'w', encoding='utf-8') as f:
            for result in filter_results:
                f.write(f"{result['title']}\n{result['href']}\n\n")
    if not filter_results:
        # 删除txt文件
        os.remove(os.path.join(save_dir, '筛选结果.txt'))
        # 如果没有筛选结果，删除save_dir文件夹
        os.rmdir(save_dir)
        print(f"未找到包含'{filter_keyword}'的记录。")
    else:
        print(f"筛选完成，共找到{len(filter_results)}条记录，已保存到{save_dir}文件夹。")


def split_date_range_backward(start_date_str, end_date_str):
    start_date = datetime.strptime(start_date_str, '%Y:%m:%d')
    end_date = datetime.strptime(end_date_str, '%Y:%m:%d')

    # 确保起始日期不晚于截止日期
    if start_date > end_date:
        raise ValueError("起始日期必须不晚于截止日期")

    date_ranges = []
    current_end_date = end_date
    while current_end_date > start_date:
        current_start_date = max(
            current_end_date - timedelta(days=365), start_date)
        date_ranges.insert(0, (current_start_date.strftime(
            '%Y:%m:%d'), current_end_date.strftime('%Y:%m:%d')))
        current_end_date = current_start_date - timedelta(days=1)

    if current_end_date == start_date:
        # 将第一个日期范围的起始日期调整为start_date
        date_ranges[0] = (start_date_str, date_ranges[0][1])

    return date_ranges


def main(start_time, end_time, keyword, filter_keyword):
    global driver
    driver = init_driver()
    results = []

    try:
        date_ranges = split_date_range_backward(start_time, end_time)
        for date_range in date_ranges:
            results.extend(get_search_list(
                date_range, keyword, driver, page_index=1))
    except ValueError as e:
        print(e)
        driver.quit()
        return

    filter_content(results, driver, filter_keyword)
    driver.quit()


def on_submit():
    start_time = f"{start_year.get()}:{start_month.get()}:{start_day.get()}"
    end_time = f"{end_year.get()}:{end_month.get()}:{end_day.get()}"
    keyword = keyword_entry.get()
    filter_keyword = filter_keyword_entry.get()

    if start_time and end_time and keyword and filter_keyword:
        console_text.delete(1.0, tk.END)
        threading.Thread(target=main, args=(
            start_time, end_time, keyword, filter_keyword)).start()
        # 清除控制台输出
    else:
        messagebox.showwarning("输入错误", "请填写所有字段")


def create_spinbox(row, column, default_value):
    spinbox = tk.Spinbox(root, from_=default_value,
                         to=default_value + 9, width=4)
    spinbox.grid(row=row, column=column)
    return spinbox


def on_closing():
    global driver
    if driver:
        driver.quit()
    root.destroy()


root = tk.Tk()
root.title("搜索工具")

# 设置窗口关闭事件处理程序
root.protocol("WM_DELETE_WINDOW", on_closing)

# 设置窗口大小和位置
root.geometry("600x400+500+300")

# 设置字体和颜色
font = ("Arial", 12)
bg_color = "#f0f0f0"
fg_color = "#000000"
root.configure(bg=bg_color)


# 结束日期选择(默认为当前日期)
tk.Label(root, text="开始日期", font=font, bg=bg_color,
         fg=fg_color).grid(row=0, column=0, padx=10, pady=5)
end_year = create_spinbox(1, 1, datetime.now().year)
end_month = create_spinbox(1, 2, datetime.now().month)
end_day = create_spinbox(1, 3, datetime.now().day)

# 开始日期选择(默认为当前日期前一年)
tk.Label(root, text="结束日期", font=font, bg=bg_color,
         fg=fg_color).grid(row=1, column=0, padx=10, pady=5)
start_year = create_spinbox(0, 1, datetime.now().year - 1)
start_month = create_spinbox(0, 2, datetime.now().month)
start_day = create_spinbox(0, 3, datetime.now().day)

# 关键词输入
tk.Label(root, text="关键词", font=font, bg=bg_color,
         fg=fg_color).grid(row=2, column=0, padx=10, pady=5)
keyword_entry = tk.Entry(root, font=font)
keyword_entry.grid(row=2, column=1, columnspan=3)

# 筛选词输入
tk.Label(root, text="筛选词", font=font, bg=bg_color,
         fg=fg_color).grid(row=3, column=0, padx=10, pady=5)
filter_keyword_entry = tk.Entry(root, font=font)
filter_keyword_entry.grid(row=3, column=1, columnspan=3)

# 提交按钮
submit_button = tk.Button(
    root, text="提交", command=on_submit, font=font, bg="#4CAF50", fg="white")
submit_button.grid(row=4, column=0, columnspan=4, padx=10, pady=10)

# 控制台输出文本框
console_text = tk.Text(root, font=font, bg="#ffffff",
                       fg="#000000", height=10, width=60)
console_text.grid(row=5, column=0, columnspan=6, padx=10, pady=10)

# 重定向控制台输出到文本框
sys.stdout = ConsoleRedirector(console_text)
sys.stderr = ConsoleRedirector(console_text)

root.mainloop()
