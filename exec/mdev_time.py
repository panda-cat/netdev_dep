#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import netmiko
import openpyxl  # Import openpyxl
import getopt
import os
import datetime
from concurrent.futures import ThreadPoolExecutor
import sys
from tqdm import tqdm, colorize  # Import tqdm and colorize


def load_excel(excel_file):
    """Loads device information from an Excel file using openpyxl."""
    devices_info = []
    try:
        workbook = openpyxl.load_workbook(excel_file)
        sheet = workbook.active  # Assuming data is in the first sheet

        # Get the header row (first row)
        header = [cell.value for cell in sheet[1]]

        # Iterate through rows starting from the second row
        for row in sheet.iter_rows(min_row=2, values_only=True):
            device_data = dict(zip(header, row))
            devices_info.append(device_data)
    except FileNotFoundError:
        print(f"Error: 找不到excel: {excel_file}")
        sys.exit(1)
    except Exception as e:
        print(f"Error 读取excel: {e}")
        sys.exit(1)
    return devices_info

# The rest of your code (execute_commands, multithreaded_execution, main) remains largely the same.

def execute_commands(devices):
    ip = devices["host"]
    user = devices["username"]
    dev_type = devices["device_type"]
    passwd = devices["password"]
    secret = devices["secret"]
    read_time = devices.get("readtime", 10)
    cfg_file = devices["dev_cfg"]

    try:
        net_devices = {
            "device_type": dev_type,
            "host": ip,
            "username": user,
            "password": passwd,
            "secret": secret,
            "read_timeout_override": read_time,
        }

        with netmiko.ConnectHandler(**net_devices) as net_connect:
            cmd_out = net_connect.send_config_from_file(cfg_file)
            cmd_out += net_connect.save_config()

        output_dir = f"./result{datetime.datetime.now():%Y%m%d}"
        os.makedirs(output_dir, exist_ok=True)
        with open(os.path.join(output_dir, f"{ip}.txt"), "w", encoding="utf-8") as tmp_fle:
            tmp_fle.write(cmd_out + "\n")
        print(f"{ip} 执行成功") # 标准输出仍然保留，用于日志或其他目的
        return True  # 返回 True 表示执行成功

    except netmiko.exceptions.NetmikoAuthenticationException:
        with open("登录失败列表.txt", "a", encoding="utf-8") as failed_ip:
            failed_ip.write(f"{ip} 用户名密码错误\n")
        print(f"{ip} 用户名密码错误")
    except netmiko.exceptions.NetmikoTimeoutException:
        with open("登录失败列表.txt", "a", encoding="utf-8") as failed_ip:
            failed_ip.write(f"{ip} 登录超时\n")
        print(f"{ip} 登录超时")

    return False # 返回 False 表示执行失败

def multithreaded_execution(devices, num_threads):
    device_futures = [] # 用于保存设备和 future 对象的列表
    with ThreadPoolExecutor(num_threads) as pool:
        for device in devices:
            future = pool.submit(execute_commands, device)
            device_futures.append((device, future)) # 存储设备信息和 future 对象

        for device, future in tqdm(device_futures, desc="设备执行进度"): # 使用 tqdm 包装 device_futures
            ip = device['host']
            result = future.result() # 从 future 对象获取结果
            if result:
                tqdm.write(colorize(f"{ip} 执行成功", color="green")) # 使用 colorize 输出绿色成功消息


def main(argv):
    try:
        opts, args = getopt.getopt(argv, "i:t:", ["input=", "threads="])
    except getopt.GetoptError:
        print("Usage: connexec -i <excel_file> -t <num_threads default:4>")
        sys.exit(2)

    excel_file = ""
    num_threads = 4
    for opt, arg in opts:
        if opt in ("-i", "--input"):
            excel_file = arg
        elif opt in ("-t", "--threads"):
            num_threads = int(arg)

    devices = load_excel(excel_file)
    multithreaded_execution(devices, num_threads)

if __name__ == "__main__":
    main(sys.argv[1:])
