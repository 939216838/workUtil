# This is a sample Python script.
import datetime
import re


# Press Shift+F10 to execute it or replace it with your code.
# Press Double Shift to search everywhere for classes, files, tool windows, actions, and settings.


def print_hi(name):
    # Use a breakpoint in the code line below to debug your script.
    print(f'Hi, {name}')  # Press Ctrl+F8 to toggle the breakpoint.

# 获取字符串里的数字
def get_numbers(s):
    return ''.join(re.findall(r'\d+', s))

# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    #     print_hi('PyCharm')
    #
    #     print("\\t" + "shan" + "\\\\" + "t")
    #
    # # See PyCharm help at https://www.jetbrains.com/help/pycharm/
    #     now = datetime.datetime.now().month % 12
    #
    #     # next_month = (now.month % 12) + 1  # 取模12后加上1就是下一个月的月份
    #     print(now)
    print(get_numbers("567不是你"))
    print(get_numbers("5月"))