#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import netmiko
import multiprocessing
import openpyxl
import getopt
import os
import datetime
from concurrent.futures import ThreadPoolExecutor


def load_excel(excel_file):
    """
    使用 openpyxl 读取 Excel 文件并返回字典列表。
    """
    try:
        workbook = openpyxl.load_workbook(excel_file)
        sheet = workbook["Sheet1"]
        header = [cell.value for cell in sheet[1]]
        if not all(header):
            return []
        devices_info = []
        for row in sheet.iter_rows(min_row=2, values_only=True):
            if not all(row):
                continue
            row_dict = dict(zip(header, row))
            devices_info.append(row_dict)
        return devices_info
    except FileNotFoundError:
        print(f"错误：找不到文件 '{excel_file}'")
        return None
    except Exception as e:
        print(f"读取 Excel 文件时发生错误：{e}")
        return None

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

            "read_timeout_override": read_time,
        }
        if secret:
          net_devices["secret"]=secret
        net_connect = netmiko.ConnectHandler(**net_devices)

        
        with net_connect:
            if dev_type == "paloalto_panos":
                cmd_out = net_connect.send_multiline(cmds, expect_string=r">", cmd_verify=False)
            elif dev_type in ("huawei", "huawei_telnet", "hp_comware", "hp_comware_telnet"):
                cmd_out = net_connect.send_multiline(cmds, cmd_verify=False)
            else:
              if secret: #  如果secret非空字符串
                  net_connect.enable()
              cmd_out = net_connect.send_multiline(cmds, cmd_verify=False) # 所有情况都需要发送命令

        output_dir = f"./result{datetime.datetime.now():%Y%m%d}"
        os.makedirs(output_dir, exist_ok=True)
        with open(os.path.join(output_dir, f"{ip}.txt"), "w", encoding="utf-8") as tmp_fle:
            tmp_fle.write(cmd_out + "\n")
        print(f"{ip} 执行成功")
        return cmd_out

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
        pool.map(execute_commands, devices)


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
    
    if not excel_file:
        print("错误：必须指定 Excel 文件路径。")
        sys.exit(2)
    devices = load_excel(excel_file)
    if devices:
        multithreaded_execution(devices, num_threads)
    else:
        print("读取设备信息失败，请检查 Excel 文件。")


if __name__ == "__main__":
    main(sys.argv[1:])
