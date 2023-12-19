#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import netmiko
import multiprocessing
import pandas as pd
import getopt
import os
import datetime
from concurrent.futures import ThreadPoolExecutor


def load_excel(excel_file):
    df = pd.read_excel(excel_file, sheet_name="Sheet1")
    devices_info = df.to_dict(orient="records")
    return devices_info


def execute_commands(devices):
    ip = devices["host"]
    user = devices["username"]
    dev_type = devices["device_type"]
    passwd = devices["password"]
    secret = devices["secret"]
    read_time = devices["readtime"]
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
        if dev_type == "paloalto_panos":
            cmd_out = net_connect.send_multiline(cmds, expect_string=r">")
        elif dev_type in ("huawei", "huawei_telnet", "hp_comware", "hp_comware_telnet"):
            cmd_out = net_connect.send_multiline(cmds)
        else:
            net_connect.enable()
            cmd_out = net_connect.send_multiline(cmds)

        output_dir = f"./result{datetime.datetime.now():%Y%m%d}"
        os.makedirs(output_dir, exist_ok=True)
        with open(os.path.join(output_dir, f"{ip}.txt"), "w", encoding="utf-8") as tmp_fle:
            tmp_fle.write(cmd_out + "\n")
        print(f"{ip} 执行成功")
        return cmd_out

    except netmiko.exceptions.NetmikoAuthenticationException:
        with open("登录失败列表", "a", encoding="utf-8") as failed_ip:
            failed_ip.write(f"{ip} 用户名密码错误\n")
        print(f"{ip} 用户名密码错误")
    except netmiko.exceptions.NetmikoTimeoutException:
        with open("登录失败列表", "a", encoding="utf-8") as failed_ip:
            failed_ip.write(f"{ip}  登录超时\n")
        print(f"{ip} 登录超时")

    return None


def multithreaded_execution(devices, num_threads):
    with ThreadPoolExecutor(num_threads) as pool:
        pool.map(execute_commands, devices)


def main(argv):
    try:
        opts, args = getopt.getopt(argv, "c:t:", ["excel=", "threads="])
    except getopt.GetoptError:
        print("Usage: connexec -c <excel_file> -t <num_threads default:4>")
        sys.exit(2)

    excel_file = ""
    num_threads = 4
    for opt, arg in opts:
        if opt in ("-c", "--excel"):
            excel_file = arg
        elif opt in ("-t", "--threads"):
            num_threads = int(arg)

    devices = load_excel(excel_file)
    multithreaded_execution(devices, num_threads)


if __name__ == "__main__":
    main(sys.argv[1:])
