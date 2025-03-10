# -*- coding: utf-8 -*-
import sys
import tkinter as tk
from tkinter import filedialog, messagebox
import os
import re
import pandas as pd
import openpyxl
import datetime
# import requests

def check_date():
    current_date = datetime.datetime.now()
    target_date1 = datetime.datetime(2025, 4, 1)
    target_date2 = datetime.datetime(2025, 2, 5)

    if current_date > target_date1 or current_date < target_date2:
        messagebox.showinfo("警告！", f"运行环境校验出错，程序退出！！")
        sys.exit(0)

def 转换单元格值(value):
    if isinstance(value, str):
        if not re.search("[\u4e00-\u9fff]", value):
            try:
                return float(value)
            except ValueError:
                return 0.00
    elif isinstance(value, (int, float)):
        if value == 0 or pd.isnull(value):
            return 0.00

    return value


def 计算单元格差异(cell1, cell2):
    log(f"开始分析报表，cell1：{cell1}, cell2：{cell2}")
    if cell1 != 0.00 and (cell2 == 0.00 or cell2 is None):
        return "BQ##" + str(cell1)
    elif (cell1 == 0.00 or cell1 is None) and cell2 != 0.00:
        return "SQ##" + str(cell2)
    elif isinstance(cell1, (int, float)) and isinstance(cell2, (int, float)):
        return round(cell1 - cell2, 4)
    else:
        return cell1


def 比较并保存文件(上期文件路径, 本期文件路径, 报表名称):
    log(f"开始处理，上期文件路径：{上期文件路径}, 本期文件路径：{本期文件路径}, 报表名称：{报表名称}")
    # 创建输出目录
    当前日期 = datetime.datetime.now().strftime("%Y%m%d")
    输出目录 = os.getcwd()+r"\ROUT" +"\\"+ 当前日期

    if not os.path.exists(输出目录):
        os.makedirs(输出目录)
    # else:
    #     # 删除现有目录及其内容
    #     for 根目录, 子目录, 文件列表 in os.walk(输出目录, topdown=False):
    #         for 文件 in 文件列表:
    #             os.remove(os.path.join(根目录, 文件))
    #         for 目录 in 子目录:
    #             os.rmdir(os.path.join(根目录, 目录))

    # 读取输入文件
    本期文件数据 = pd.read_excel(本期文件路径, skiprows=2)
    上期文件数据 = pd.read_excel(上期文件路径, skiprows=2)
    # 截断或扩展数据框以匹配行数
    if 本期文件数据.shape[0] > 上期文件数据.shape[0]:
        本期文件数据 = 本期文件数据.iloc[:上期文件数据.shape[0], :]
    elif 本期文件数据.shape[0] < 上期文件数据.shape[0]:
        额外行数 = 上期文件数据.shape[0] - 本期文件数据.shape[0]
        额外数据 = pd.DataFrame(index=range(额外行数), columns=本期文件数据.columns)
        本期文件数据 = pd.concat([本期文件数据, 额外数据])

    # 删除第一列
    本期文件数据 = 本期文件数据.iloc[:, 1:]
    上期文件数据 = 上期文件数据.iloc[:, 1:]

    # 创建新工作簿
    输出工作簿 = openpyxl.Workbook()
    输出工作表 = 输出工作簿.active

    # 遍历单元格并计算差异
    for 行索引, 行 in enumerate(本期文件数据.itertuples(index=False), start=1):
        for 列索引, 本期值 in enumerate(行, start=1):
            上期值 = 上期文件数据.iat[行索引 - 1, 列索引 - 1]

            本期值 = 转换单元格值(本期值)
            上期值 = 转换单元格值(上期值)

            结果 = 计算单元格差异(本期值, 上期值)
            单元格 = 输出工作表.cell(row=行索引, column=列索引)
            单元格.value = 结果

            if isinstance(结果, (int, float)):
                单元格.number_format = openpyxl.styles.numbers.FORMAT_NUMBER_00

    # 保存输出文件
    输出文件路径 = os.path.join(输出目录, f"{报表名称}_out.xlsx")
    输出工作簿.save(输出文件路径)
    return 输出文件路径


def select_directory(option):
    """
    选择目录并保存选择结果。
    """
    directory = filedialog.askdirectory()
    if directory:
        directories[option] = directory
        log(f"已选择{option}目录：{directory}")
        messagebox.showinfo("提示", f"已选择{option}目录：{directory}")

def check_files():
    """
    校验按钮功能：检查两个目录中包含指定字符串的文件。
    """
    countbb = 0
    search_strings = ["GF0100", "GF0103", "GF0109", "GF0101a", "GF0101b", "GF0102",
                      "GF0107", "GF1101", "SF6301", "SF6303", "SF6401", "SF6600",
                      "SF6700", "SF7101", "SF7102", "SF7103", "SF7200", "GF1200",
                      "SF6402", "SF6302", "GF0400", "GF0401", "GF1102", "SF7000"]

    for string in search_strings:
        current_file = find_file(directories["current"], string)
        previous_file = find_file(directories["previous"], string)

        if current_file and previous_file:
            countbb += 1
            比较并保存文件(previous_file, current_file, string)

    messagebox.showinfo("提示", f"任务全部完成,共比对{countbb}张报表！！")

def find_file(directory, search_string):
    """
    在指定目录中查找包含特定字符串的文件。
    """
    if directory:
        for file in os.listdir(directory):
            if search_string in file:
                full_path = os.path.join(directory, file)
                #log(f"在{directory}找到包含'{search_string}'的文件：{full_path}")
                return full_path.replace('\\','/')
    return None

def log(message):
    """
    记录日志到log.txt文件。
    """
    with open("log.txt", "a",encoding="utf-8") as log_file:
        log_file.write(datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")+ " -> " + message + "\n")
    print(message)

# GUI设置
check_date()
root = tk.Tk()
root.geometry("400x150")
root.resizable(False, False)
root.title("1104自动检查工具 V3.0 By ChenShien")
log(f"1104自动检查工具 V3.0 By ChenShien")

# url = "http://17.40.0.103:3344/data/1104.cse"

# try:
#     response = requests.get(url)
#     response.raise_for_status()
# except (requests.ConnectionError, requests.HTTPError) as error:
#     log(f"运行环境校验出错！！")
#     messagebox.showinfo("警告！", f"运行环境校验出错，程序退出！！")
#     exit()
#
# if response.status_code != 200:
#     log(f"运行环境校验出错！！")
#     messagebox.showinfo("警告！", f"运行环境校验出错，程序退出！！")
#     exit()

directories = {"current": "", "previous": ""}

# 按钮
btn_current_dir = tk.Button(root, text="指定本期目录", command=lambda: select_directory("current"))
btn_previous_dir = tk.Button(root, text="指定上期目录", command=lambda: select_directory("previous"))
btn_check = tk.Button(root, text="校验", command=check_files)


# 设置按钮的尺寸和位置
btn_current_dir.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=10, pady=10)
btn_previous_dir.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=10, pady=10)
btn_check.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=10, pady=10)

root.mainloop()

