# ======================
# -*- coding:utf-8 -*-
# @author:Beall
# @time  :2021/12/8 11:05
# @file  :main.py
# ======================
import pyautogui
import time
import xlrd
import pyperclip


# 定义鼠标事件
def MouseClick(clickTimes, lOrR, img, reTry):
    if reTry == 1:
        while True:
            location = pyautogui.locateCenterOnScreen(img, confidence=0.9)
            if location is not None:
                pyautogui.click(location.x, location.y, clicks=clickTimes, interval=0.2, duration=0.2, button=lOrR)
                break
            print("未找到匹配图片,0.1秒后重试")
            time.sleep(0.1)
    elif reTry == -1:
        while True:
            location = pyautogui.locateCenterOnScreen(img, confidence=0.9)
            if location is not None:
                pyautogui.click(location.x, location.y, clicks=clickTimes, interval=0.2, duration=0.2, button=lOrR)
            time.sleep(0.1)
    elif reTry > 1:
        i = 1
        while i < reTry + 1:
            location = pyautogui.locateCenterOnScreen(img, confidence=0.9)
            if location is not None:
                pyautogui.click(location.x, location.y, clicks=clickTimes, interval=0.2, duration=0.2, button=lOrR)
                print("重复")
                i += 1
            time.sleep(0.1)


# 数据检查
def DataCheck(_sheet):
    """
    cmd_type.value  1.0 左键单击    2.0 左键双击  3.0 右键单击  4.0 输入  5.0 等待  6.0 滚轮
    cmd_type     空：0
              字符串：1
              数字：2
              日期：3
              布尔：4
              error：5
    :param _sheet:
    :return:
    """
    check_cmd = True
    # 行数检查
    if _sheet.nrows < 2:
        print("没数据啊哥")
        check_cmd = False
    # 每行数据检查
    i = 1
    while i < _sheet.nrows:
        # 第1列 操作类型检查
        cmd_type = _sheet.row(i)[0]
        if cmd_type.ctype != 2 or (cmd_type.value != 1.0 and cmd_type.value != 2.0 and cmd_type.value != 3.0 and
                                   cmd_type.value != 4.0 and cmd_type.value != 5.0 and cmd_type.value != 6.0 and
                                   cmd_type.value != 7.0):
            print('第', i + 1, "行,第1列数据有毛病")
            check_cmd = False
        # 第2列 内容检查
        cmd_value = _sheet.row(i)[1]
        # 读图点击类型指令，内容必须为字符串类型
        if cmd_type.value == 1.0 or cmd_type.value == 2.0 or cmd_type.value == 3.0:
            if cmd_value.ctype != 1:
                print('第', i + 1, "行,第2列数据有毛病")
                check_cmd = False
        # 输入类型，内容不能为空
        if cmd_type.value == 4.0:
            if cmd_value.ctype == 0:
                print('第', i + 1, "行,第2列数据有毛病")
                check_cmd = False
        # 等待类型，内容必须为数字
        if cmd_type.value == 5.0:
            if cmd_value.ctype != 2:
                print('第', i + 1, "行,第2列数据有毛病")
                check_cmd = False
        # 滚轮事件，内容必须为数字
        if cmd_type.value == 6.0:
            if cmd_value.ctype != 2:
                print('第', i + 1, "行,第2列数据有毛病")
                check_cmd = False
        # 输入类型，内容不能为空
        if cmd_type.value == 7.0:
            if cmd_value.ctype == 0:
                print('第', i + 1, "行,第2列数据有毛病")
                check_cmd = False
        i += 1
    return check_cmd


# 任务
def MainWork(img):
    i = 1
    while i < sheet1.nrows:
        # 取本行指令的操作类型
        cmd_type = sheet1.row(i)[0]
        if cmd_type.value == 1.0:
            # 取图片名称
            img = 'img/' + sheet1.row(i)[1].value
            re_try = 1
            if sheet1.row(i)[2].ctype == 2 and sheet1.row(i)[2].value != 0:
                re_try = sheet1.row(i)[2].value
            MouseClick(1, "left", img, re_try)
            print("单击左键", img)
        # 2代表双击左键
        elif cmd_type.value == 2.0:
            # 取图片名称
            img = 'img/' + sheet1.row(i)[1].value
            # 取重试次数
            re_try = 1
            if sheet1.row(i)[2].ctype == 2 and sheet1.row(i)[2].value != 0:
                re_try = sheet1.row(i)[2].value
            MouseClick(2, "left", img, re_try)
            print("双击左键", img)
        # 3代表右键
        elif cmd_type.value == 3.0:
            # 取图片名称
            img = 'img/' + sheet1.row(i)[1].value
            # 取重试次数
            re_try = 1
            if sheet1.row(i)[2].ctype == 2 and sheet1.row(i)[2].value != 0:
                re_try = sheet1.row(i)[2].value
            MouseClick(1, "right", img, re_try)
            print("右键", img)
        # 4代表输入
        elif cmd_type.value == 4.0:
            input_value = sheet1.row(i)[1].value
            pyperclip.copy(input_value)
            pyautogui.hotkey('ctrl', 'v')
            time.sleep(0.5)
            print("输入:", input_value)
        # 5代表等待
        elif cmd_type.value == 5.0:
            wait_time = sheet1.row(i)[1].value
            time.sleep(wait_time)
            print("等待", wait_time, "秒")
        # 6代表滚轮
        elif cmd_type.value == 6.0:
            scroll = sheet1.row(i)[1].value
            pyautogui.scroll(int(scroll))
            print("滚轮滑动", int(scroll), "距离")
        # 7代表按键
        elif cmd_type.value == 7.0:
            hotkey = sheet1.row(i)[1].value.split(',')
            if len(hotkey) == 1:
                pyautogui.hotkey(hotkey[0])
            elif len(hotkey) == 2:
                pyautogui.hotkey(hotkey[0], hotkey[1])
            print("按键", hotkey)
        i += 1


if __name__ == '__main__':
    file = 'cmd.xls'
    # 打开文件
    wb = xlrd.open_workbook(filename=file)
    # 通过索引获取表格sheet页
    sheet1 = wb.sheet_by_index(0)
    print('欢迎使用beall牌RPA~')
    # 数据检查
    checkCmd = DataCheck(sheet1)
    if checkCmd:
        key = input('选择功能（1.做一次 2.循环到死）:\t')
        if key == '1':
            # 循环拿出每一行指令
            MainWork(sheet1)
        elif key == '2':
            while True:
                MainWork(sheet1)
                time.sleep(0.1)
                print("等待0.1秒")
    else:
        print('输入有误或者已经退出!')
