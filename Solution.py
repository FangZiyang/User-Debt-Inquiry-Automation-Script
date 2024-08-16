import os

import openpyxl
import pyautogui
import time
import logging

# 设置日志
logging.basicConfig(filename='automation.log', level=logging.INFO,
                    format='%(asctime)s - %(levelname)s - %(message)s')

# 1. 读取Excel文件
wb = openpyxl.load_workbook('AAA.xlsx')
sheet = wb.active

# 2. 提取数据
id_list = [cell.value for cell in sheet['B'][1:]]
id_val = {row[1].value: row[2].value for row in sheet.iter_rows(min_row=2, max_col=3)}

# 存储欠费ID
overdue_ids = []

# 断点文件名
CHECKPOINT_FILE = 'checkpoint.txt'


# 读取断点
def read_checkpoint():
    if os.path.exists(CHECKPOINT_FILE):
        with open(CHECKPOINT_FILE, 'r') as f:
            return int(f.read().strip())
    return 0


# 保存断点
def save_checkpoint(index):
    with open(CHECKPOINT_FILE, 'w') as f:
        f.write(str(index))


# 读取上次处理的位置
start_index = read_checkpoint()

# 3. 循环处理
for index, id in enumerate(id_list[start_index:], start_index + 1):
    logging.info(f"正在处理第 {index}/{len(id_list)} 个ID: {id}")

    # 3.1 鼠标移动并击
    pyautogui.click(370, 472)
    time.sleep(2)

    # 3.2 清空输入行
    pyautogui.doubleClick(359, 209)
    pyautogui.hotkey('ctrl', 'a')
    pyautogui.press('delete')
    time.sleep(1)

    # 3.3 输入数据
    pyautogui.write(f"0{id}")

    # 3.4 和 3.5 鼠标双击
    pyautogui.click(436, 347)
    time.sleep(1)
    pyautogui.click(450, 860)

    # 3.6 等待
    time.sleep(8)  # 可根据实际情况调整等待时间

    # 3.7 检查是否欠费
    try:
        if pyautogui.locateOnScreen('out_money.png', region=(0, 0, 583, 247), confidence=0.5):
            overdue_ids.append(id)
            overdue_message = f"发现欠费: {id}"
            print(overdue_message)
            logging.info(overdue_message)

            if id_val.get(id) == 1:
                logging.warning(f"ID {id} 已欠费且在字典中的值为1")
                print("已欠费未订正" + {id})
    except pyautogui.ImageNotFoundException:
        pass
    except Exception as e:
        print(f"发生错误：{e}")

    # 3.8 输出进度日志
    logging.info(f"已完成 {index}/{len(id_list)}")
    save_checkpoint(index)  # 保存当前处理位置

    # 3.9 双击返回
    pyautogui.click(132, 38)
    time.sleep(2)  # 给系统一些响应时间

# 输出结果
logging.info(f"处理完成。共发现 {len(overdue_ids)} 个欠费ID")
logging.info(f"欠费ID列表: {overdue_ids}")
