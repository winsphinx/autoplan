#!/usr/bin/env python
# -- coding: utf-8 --

import datetime
import json
import os
import time

import colorama
import xlrd
import xlwt
from selenium import webdriver
from selenium.webdriver import ActionChains
from xlutils.copy import copy

colorama.init()
RED = colorama.Fore.RED
GREEN = colorama.Fore.GREEN
END = colorama.Style.RESET_ALL


def get_file_list(file_dir):
    filelist = []
    for parent, _, filenames in os.walk(file_dir):
        for filename in filenames:
            if os.path.splitext(filename)[1] == '.xls':
                filelist.append(os.path.join(parent, filename))
    return filelist


def edit_file(filename):
    book = xlrd.open_workbook(filename, formatting_info=True)
    sheet = book.sheets()[0]
    max_row = sheet.nrows
    max_col = sheet.ncols

    new_book = copy(book)
    new_sheet = new_book.get_sheet(0)

    style = xlwt.XFStyle()
    style.num_format_str = 'YYYY-MM-DD'

    font = xlwt.Font()
    font.colour_index = 2  # 2:红色
    style.font = font

    borders = xlwt.Borders()
    # 细实线:1，小粗实线:2，细虚线:3，中细虚线:4，大粗实线:5，双线:6，细点虚线:7
    # 大粗虚线:8，细点划线:9，粗点划线:10，细双点划线:11，粗双点划线:12，斜点划线:13
    borders.left = 1
    borders.right = 1
    borders.top = 1
    borders.bottom = 1
    style.borders = borders

    for row in range(max_row):
        for col in range(max_col):
            if sheet.cell_type(row, col) == 3:  # 3:日期格式
                new_sheet.write(row, col, datetime.date.today(), style)

    new_book.save(filename)


def change_date():
    file_list = get_file_list(os.getcwd())
    for file in file_list:
        print(file)
        edit_file(file)


def login():
    browser = webdriver.Ie()
    browser.get("http://10.202.3.27/WebApp/emoss/index.jsp")

    n = browser.find_element_by_id("tfAccount")
    old_name = n.get_attribute('value')
    name = input(f"\n名字（回车还用 {old_name} ，或输入新的）: ")

    if name == "":
        name = old_name.lower()
    else:
        name = name.lower()
        n.clear()
        n.send_keys(name)

    with open("password.json", 'r') as f:
        password = json.load(f)
    pswd = password.get(name)

    code = browser.find_element_by_id("checkCode").get_attribute("innerHTML")

    browser.find_element_by_id("tfPassWord").send_keys(pswd)
    browser.find_element_by_id("inputCode").send_keys(code)
    browser.find_element_by_id("btLogin").click()

    time.sleep(5)
    url = r"http://10.202.3.27/WebApp/emoss/files/taskperformcontrol/taskmanagesheetmode.jsp?taskType=0&flowIndex=0"
    browser.get(url)

    return browser


def get_tasklist(browser):
    tasks = list()
    tds = browser.find_elements_by_tag_name("td")
    jobs = tds[0].text.split("\n")
    for i in range(4, len(jobs), 2):
        t = jobs[i].replace("/", "-")
        tasks.append(t)

    return tasks


def do_tasklist(browser, tasks):
    main_windows = browser.current_window_handle
    for i in range(len(tasks)):
        id = browser.find_element_by_id(i + 1)
        ActionChains(browser).double_click(id).perform()
        time.sleep(3)
        windows = browser.window_handles
        # print(windows)
        for w in windows:
            if w != main_windows:
                browser.switch_to.window(w)
                filename = os.path.join(os.getcwd(), "files",
                                        (tasks[i] + ".xls"))
                try:
                    browser.find_element_by_id("FILE").send_keys(filename)
                    browser.find_element_by_id("save").click()
                except Exception as e:
                    print(f"{filename} -- {RED}不成功{END}")
                    print(f"{RED}{e}{END}")
                else:
                    print(f"{filename} -- {GREEN}已上传{END}")
                    time.sleep(2)
                    browser.switch_to.alert.accept()
                finally:
                    browser.close()
                    browser.switch_to.window(main_windows)


if __name__ == '__main__':
    print(f"{RED}修改文件日期。。。{END}")
    change_date()
    print(f"{GREEN}文件日期改好了。。。{END}")

    browser = login()

    last_list = list()

    while (True):
        tasks = get_tasklist(browser)
        if last_list == tasks:
            break
        else:
            last_list = tasks
            do_tasklist(browser, tasks)
            browser.find_element_by_id("imgNext").click()
            time.sleep(5)

    print("\n\n完成！")
