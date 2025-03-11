#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import netmiko
import openpyxl  # 导入 openpyxl
import getopt
import os
import datetime
from concurrent.futures import ThreadPoolExecutor
import sys
from tqdm import tqdm  # 修改导入语句，只导入 tqdm

def load_excel(excel_file):
    """从 Excel 文件加载设备信息 (使用 openpyxl)."""
    devices_info = []
    try:
        workbook = openpyxl.load_workbook(excel_file)
        sheet = workbook.active  # 假设数据在第一个 sheet 中

        # 获取标题行 (第一行)
        header = [cell.value for cell in sheet[1]]

        # 从第二行开始迭代行
        for row in sheet.iter_rows(min_row=2, values_only=True):
            device_data = dict(zip(header, row))
            devices_info.append(device_data)
    except FileNotFoundError:
        print(f"Error: 找不到 excel: {excel_file}")
        sys.exit(1)
    except Exception as e:
        print(f"Error 读取excel: {e}")
        sys.exit(1)
    return devices_info

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
        print(f"{ip} 执行成功")
        return True

    except netmiko.exceptions.NetmikoAuthenticationException:
        with open("登录失败列表.txt", "a", encoding="utf-8") as failed_ip:
            failed_ip.write(f"{ip} 用户名密码错误\n")
        print(f"{ip} 用户名密码错误")
    except netmiko.exceptions.NetmikoTimeoutException:
        with open("登录失败列表.txt", "a", encoding="utf-8") as failed_ip:
            failed_ip.write(f"{ip} 登录超时\n")
        print(f"{ip} 登录超时")

    return False

def multithreaded_execution(devices, num_threads):
    device_futures = []
    with ThreadPoolExecutor(num_threads) as pool:
        for device in devices:
            future = pool.submit(execute_commands, device)
            device_futures.append((device, future))

        for device, future in tqdm(device_futures, desc="设备执行进度"):
            ip = device['host']
            result = future.result()
            if result:
                tqdm.write(f"{ip} 执行成功") # 移除 colorize，直接打印消息

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
