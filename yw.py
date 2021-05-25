#!/usr/bin/env python
# -- coding: utf-8 --
"""
Usage:
  yw
  yw (-h | --help)
  yw (-a | --add) <username> <password>
  yw (-d | --del) <username>
  yw (-u | --use) <username>

Options:
  无参数        根据提示手动执行
  -h, --help    显示帮助
  -a, --add     增加或修改用户名和密码
  -d, --del     删除指定用户名
  -u, --use     使用指定用户名执行

"""

import base64
import datetime
import json
import os
import sys

import colorama
import docopt
import xlrd
import xlwt
from selenium import webdriver
from selenium.webdriver import ActionChains
from selenium.webdriver.support import expected_conditions
from selenium.webdriver.support.ui import WebDriverWait
from xlutils.copy import copy

colorama.init()
RED = colorama.Fore.RED
GREEN = colorama.Fore.GREEN
END = colorama.Style.RESET_ALL

URL = "http://10.202.3.27/WebApp/emoss"


def dec(s):
    return str(base64.b64decode(bytes(s, "utf-8")), "utf-8")


def enc(s):
    return str(base64.b64encode(bytes(s, "utf-8")), "utf-8")


def add_user(user, pswd):
    data = read_data()
    data[user] = enc(pswd)
    write_data(data)


def del_user(user):
    data = read_data()
    try:
        del data[user]
        write_data(data)
    except KeyError:
        print(f"\n{RED}用户名不存在！{END}")


def read_data():
    try:
        with open("password.json", "r") as f:
            return json.load(f)
    except FileNotFoundError:
        open("password.json", "w")
        return {}


def write_data(data):
    with open("password.json", "w") as f:
        json.dump(data, f, indent=2)


def get_file_list(file_dir):
    filelist = []
    for parent, _, filenames in os.walk(file_dir):
        for filename in filenames:
            if os.path.splitext(filename)[1] == ".xls":
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
    style.num_format_str = "YYYY-MM-DD"

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
    print(f"\n{RED}修改文件日期。。。{END}")

    file_list = get_file_list(os.getcwd())
    for file in file_list:
        print(file)
        edit_file(file)

    print(f"\n{GREEN}文件日期改好了！{END}")


def login(username):
    print(f"\n{RED}开始登录系统。。。{END}")
    browser = webdriver.Ie()
    browser.maximize_window()
    browser.get(URL + "/index.jsp")

    n = browser.find_element_by_id("tfAccount")

    if username:
        name = username
    else:
        old_name = n.get_attribute("value")
        name = input(f"\n名字（回车还用 {old_name} ，或输入新的）: ")

        if name == "":
            name = old_name.lower()
        else:
            name = name.lower()

    n.clear()
    n.send_keys(name)

    try:
        with open("password.json", "r") as f:
            password = json.load(f)
            pswd = dec(password.get(name, ""))
            if pswd == "":
                print(f"\n\n{RED}没有这个人！{END}")
                sys.exit(0)
    except FileNotFoundError:
        print(f"\n\n{RED}请先添加用户！{END}")
        sys.exit(0)

    code = browser.find_element_by_id("checkCode").get_attribute("innerHTML")

    browser.find_element_by_id("tfPassWord").send_keys(pswd)
    browser.find_element_by_id("inputCode").send_keys(code)
    browser.find_element_by_id("btLogin").click()

    print(f"\n{GREEN}已成功登录！{END}")

    WebDriverWait(browser, 30).until(
        lambda browser: browser.find_element_by_name("InfoPage")
    )
    browser.get(URL + "/files/taskperformcontrol/taskmanagesheetmode.jsp?taskType=0")

    return browser


def get_tasklist(browser):
    tasks = []
    WebDriverWait(browser, 10).until(
        lambda browser: browser.find_element_by_name("taskForm")
    )
    tds = browser.find_elements_by_tag_name("td")
    jobs = tds[0].text.split("\n")
    for i in range(4, len(jobs), 2):
        t = jobs[i].replace("/", "-")
        tasks.append(t)

    return tasks


def execute_tasks(browser, tasks):
    main_window = browser.window_handles
    for i in range(len(tasks)):
        id = browser.find_element_by_id(i + 1)
        ActionChains(browser).double_click(id).perform()
        WebDriverWait(browser, 10).until(
            expected_conditions.new_window_is_opened(main_window)
        )
        windows = browser.window_handles
        for w in windows:
            if w != main_window[0]:
                browser.switch_to.window(w)
                filename = os.path.join(os.getcwd(), "files", (tasks[i] + ".xls"))
                try:
                    browser.find_element_by_id("FILE").send_keys(filename)
                    browser.find_element_by_id("save").click()
                except Exception as e:
                    print(f"{filename} -- {RED}不成功{END}")
                    print(f"{RED}{e}{END}")
                else:
                    alert = WebDriverWait(browser, 10).until(
                        expected_conditions.alert_is_present()
                    )
                    alert.accept()
                    print(f"{filename} -- {GREEN}已上传{END}")
                finally:
                    browser.close()
                    browser.switch_to.window(main_window[0])


def do_jobs(username=""):
    change_date()

    browser = login(username)

    last_list = list()

    print(f"\n{RED}开始上传文件。。。{END}")

    while True:
        tasks = get_tasklist(browser)
        if last_list == tasks:
            break
        else:
            last_list = tasks
            execute_tasks(browser, tasks)
            browser.find_element_by_id("imgNext").click()

    print(f"\n{GREEN}完成！{END}")


if __name__ == "__main__":
    args = docopt.docopt(__doc__)

    if args["--add"]:
        add_user(args["<username>"], args["<password>"])
    elif args["--del"]:
        del_user(args["<username>"])
    elif args["--use"]:
        do_jobs(args["<username>"])
    else:
        do_jobs()
