# Pyinstaller -F main.py
import pickle
import win32api
import win32gui
from ctypes import *
import time
from datetime import datetime
import win32con
from PIL import ImageGrab
import random

positions_cali = {}


def save_cali_position():
    save_file = open("positionCalibrated.bin", "wb")
    pickle.dump(positions_cali, save_file)  # 顺序存入变量
    save_file.close()


def load_cali_position():
    global positions_cali
    load_file = open("positionCalibrated.bin", "rb")
    positions_cali = pickle.load(load_file)  # 顺序导出变量
    load_file.close()


def get_cur_pos():
    return win32gui.GetCursorPos()


def input_all_cali_position():
    global positions_cali
    print('鼠标移动到刷新按钮上后敲回车')
    input()
    positions_cali[0] = get_cur_pos()
    print('鼠标移动到座位1上后敲回车')
    input()
    positions_cali[1] = get_cur_pos()
    print('鼠标移动到座位2上后敲回车')
    input()
    positions_cali[2] = get_cur_pos()


def test_positions():
    time.sleep(2)
    move_cur_pos(positions_cali[0])
    time.sleep(2)
    move_cur_pos(positions_cali[1])
    time.sleep(2)
    move_cur_pos(positions_cali[2])


def click_left_cur():
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN | win32con.MOUSEEVENTF_LEFTUP, 0, 0)
    time.sleep(0.1)


def move_cur_pos(position):
    windll.user32.SetCursorPos(position[0], position[1])
    time.sleep(0.03)


def roll_to_bottom():
    for i in range(1, 1500):
        win32api.mouse_event(win32con.MOUSEEVENTF_WHEEL, 0, 0, -1)


def tap_seat(num):
    move_cur_pos(positions_cali[num])
    click_left_cur()
    print('尝试座位:')
    print(num)


def get_button_color():
    image = ImageGrab.grab()
    color1 = image.getpixel((positions_cali[1][0], positions_cali[1][1]))
    color2 = image.getpixel((positions_cali[2][0], positions_cali[2][1]))
    return color1, color2


def refresh():
    while True:
        move_cur_pos(positions_cali[0])
        click_left_cur()
        move_cur_pos(positions_cali[1])
        roll_to_bottom()

        color1, color2 = get_button_color()
        if color1[0] == color1[1] and color1[0] == color1[2] and color1[0] == 172:
            break


def start():
    while 1:
        print('系统时间为：' + datetime.now().strftime('%Y-%m-%d %H:%M:%S'))
        refresh()

        for i in range(3000 + random.randint(0, 50)):
            color1, color2 = get_button_color()
            if color1[2] > color1[0]:
                tap_seat(1)
                break
            if color2[2] > color2[0]:
                tap_seat(2)
                break


tip_str = '输入相应数字进行操作：\n' \
          '1：开始挂机等待抢座\n' \
          '2：重新校准\n' \
          '3：测试校准结果\n' \
          '4：退出程序\n'
help_str = '\n使用帮助：\n' \
            'a.第一次运行请先校准\n' \
            'b.需要校准的三个位置是刷新按钮，座位1，座位2\n' \
            'c.校准生成的positionCalibrated.bin 不删除的话，下次就无需校准\n'

if __name__ == "__main__":
    print(help_str)
    try:
        load_cali_position()
    except IOError:
        print('没有位置校准数据，请先校准。')
    while 1:
        print(tip_str)
        sel = int(input("请输入:"))

        if sel == 1:
            print('已开始')
            start()
        elif sel == 2:
            input_all_cali_position()
            save_cali_position()
        elif sel == 3:
            test_positions()
        elif sel == 4:
            break
