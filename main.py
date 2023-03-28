# This is a sample Python script.
import re

def get_numbers(s):
    numbers = ''
    for c in s:
        if c.isdigit():   numbers += c
    return numbers

def print_hi(name):
    # Use a breakpoint in the code line below to debug your script.
    print(f'Hi, {name}')  # Press Ctrl+F8 to toggle the breakpoint.

# 获取字符串里的数字
# def get_numbers(s):
#     return ''.join(re.findall(r'\d+', s))
# pyinstaller --name "购售电自动填写" --onefile --hidden-import=openpyxl --hidden-import=wx --add-data="./report/GouShouDian.py;./report/" ./MainWindow/MainWindow.py
# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    #     print_hi('PyCharm')
    get_numbers("我是776")
    pass
