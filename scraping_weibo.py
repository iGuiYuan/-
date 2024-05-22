import os
import json
import re
import openpyxl
import requests
from datetime import datetime
import schedule
import time

def crawl_and_save():
    # 获取当前文件所在的目录路径
    current_dir = os.path.dirname(os.path.abspath(__file__))
    # 拼接resou文件夹的路径
    resou_dir = os.path.join(current_dir, 'resou')

    # 如果resou文件夹不存在，则创建它
    if not os.path.exists(resou_dir):
        os.makedirs(resou_dir)

    # 设置文件保存路径
    file_path = os.path.join(resou_dir, '热搜.xlsx')

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(['顺序', '热搜分类', '热搜关键词'])

    try:
        url = requests.get("https://weibo.com/ajax/side/hotSearch")
        data = json.loads(url.text)['data']['realtime']
        for i in data:
            try:
                print(f'热搜：{i["realpos"]}, 热搜分类[{i["category"]}], 热搜关键词：{i["word"]}')
                ws.append([i["realpos"], i["category"], i["word"].encode('utf-8')])
            except Exception as e:
                print(f"写入失败: {e}")
    except requests.RequestException as e:
        print(f"请求失败: {e}")
    except json.JSONDecodeError as e:
        print(f"JSON解析失败: {e}")
    finally:
        # 生成当前时间的字符串，形如：2022-01-01_12-00-00
        current_time = datetime.now().strftime('%Y-%m-%d_%H-%M-%S')
        # 拼接完整的文件名
        file_name = f"热搜_{current_time}.xlsx"
        # 拼接完整的文件路径
        file_path = os.path.join(resou_dir, file_name)
        wb.save(file_path)
        wb.close()


# 初始运行一次
crawl_and_save()

# 设定定时任务，每隔一个小时运行一次
schedule.every().hour.do(crawl_and_save)

# 无限循环执行任务
while True:
    schedule.run_pending()
    time.sleep(1)
