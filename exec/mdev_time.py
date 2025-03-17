#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import netmiko
import openpyxl  # Import openpyxl
import getopt
import os
import datetime
from tqdm import tqdm
from typing import List
from concurrent.futures import ThreadPoolExecutor
import sys  # Import sys if not already present

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
        print(f"文件没有找到: {excel_file}")
        sys.exit(1)
    except Exception as e:
        print(f"读取错误: {e}")
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
   cmds = list(devices["mult_command"].split(";"))

   try:
       net_devices = {
           "device_type": dev_type,
           "host": ip,
           "username": user,
           "password": passwd,
           "secret": secret,
           "read_timeout_override": read_time,
       }
       net_connect = netmiko.ConnectHandler(**net_devices)

       hname = net_connect.find_prompt().strip('#<>[]*:?')
       
       with net_connect:
            if dev_type == "paloalto_panos":
                cmd_out = net_connect.send_multiline(cmds, expect_string=r">", cmd_verify=False)
            elif dev_type in ("huawei", "huawei_telnet", "hp_comware", "hp_comware_telnet"):
                cmd_out = net_connect.send_multiline(cmds, cmd_verify=False)
            elif secret:
                net_connect.enable()
                cmd_out = net_connect.send_multiline(cmds, cmd_verify=False)
            else:
                cmd_out = net_connect.send_multiline(cmds, cmd_verify=False)

       output_dir = f"./result{datetime.datetime.now().strftime('%Y%m%d')}"
       os.makedirs(output_dir, exist_ok=True)
       with open(os.path.join(output_dir, f"{ip}_{hname}.txt"), "w", encoding="utf-8") as tmp_fle:
           tmp_fle.write(cmd_out + "\n")
       print(f"{ip} 执行成功")

   except netmiko.exceptions.NetmikoAuthenticationException:
       with open("登录失败列表.txt", "a", encoding="utf-8") as failed_ip:
           failed_ip.write(f"{ip} 用户名密码错误\n")
       print(f"{ip} 用户名密码错误")
   except netmiko.exceptions.NetmikoTimeoutException:
       with open("登录失败列表.txt", "a", encoding="utf-8") as failed_ip:
           failed_ip.write(f"{ip} 登录超时\n")
       print(f"{ip} 登录超时")

   return None

def multithreaded_execution(devices, num_threads):
   with ThreadPoolExecutor(num_threads) as pool:
       all_results = list(tqdm(pool.map(execute_commands, devices),
                               total=len(devices),
                               desc="执行进度",
                               unit="台"))
        
        # 统计结果
        success = sum(1 for r in all_results if r is not None)
        print(f"\n执行完成: {success}台成功, {len(devices)-success}台失败")

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

   try:
       devices = load_excel(excel_file)
       print(f"成功加载 {len(devices)} 台设备信息")
       multithreaded_execution(devices, num_threads)
   except KeyboardInterrupt:
       print("\n操作已取消")
       sys.exit(0)

if __name__ == "__main__":
   main(sys.argv[1:])
