```python
#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import sys
import getopt
import os
import datetime
import netmiko
import multiprocessing
import pandas as pd
from concurrent.futures import ThreadPoolExecutor


def load_excel(excel_file):
    """
    Load devices information from an excel file.

    Args:
        excel_file (str): The path to the excel file.

    Returns:
        list: A list of dictionaries, where each dictionary represents a device.
    """
    df = pd.read_excel(excel_file, sheet_name="Sheet1")
    devices_info = df.to_dict(orient="records")
    return devices_info


def execute_commands(device, folder):
    """
    Execute commands on a device and save the output to a file.

    Args:
        device (dict): A dictionary representing a device.
        folder (str): The path to the output folder.
    """
    ip = device["host"]
    user = device["username"]
    dev_type = device["device_type"]
    passwd = device["password"]
    secret = device["secret"]
    read_time = device["readtime"]
    cmds = list(device["mult_command"].split(";"))

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

        output_dir = folder

        with open(os.path.join(output_dir, f"{ip}.txt"), "w", encoding="utf-8") as tmp_fle:
            tmp_fle.write(cmd_out + "\n")
        print(f"{ip} 执行成功")
        return None

    except netmiko.exceptions.NetmikoAuthenticationException:
        with open("登录失败列表", "a", encoding="utf-8") as failed_ip:
            failed_ip.write(f"{ip} 用户名密码错误\n")
            print(f"{ip} 用户名密码错误")
    except netmiko.exceptions.NetmikoTimeoutException:
        with open("登录失败列表", "a", encoding="utf-8") as failed_ip:
            failed_ip.write(f"{ip} 登录超时\n")
            print(f"{ip} 登录超时")

    return None


def multithreaded_execution(devices, num_threads, folder):
    """
    Execute commands on devices in a multithreaded manner.

    Args:
        devices (list): A list of dictionaries, where each dictionary represents a device.
        num_threads (int): The number of threads to use.
        folder (str): The path to the output folder.
    """
    with ThreadPoolExecutor(num_threads) as pool:
        pool.map(execute_commands, devices, folder)


def help_man():
    """
    Print help message.
    """
    print("Usage: connexec -t <num_threads default:4> -o <output_folder> -c <excel_file>")
    sys.exit(0)  # Exit script


def main(argv):
    """
    The main function.
    """
    try:
        opts, args = getopt.getopt(argv, "hc:t:o:", ["help", "excel=", "threads=", "output_floder="])
    except getopt.GetoptError:
        help_man()

    excel_file = None
    output_folder = None
    num_threads = 4
    for opt, arg in opts:
        if opt in ("-c", "--excel"):
            excel_file = arg
        elif opt in ("-t", "--threads"):
            num_threads = int(arg)
        elif opt in ("-o", "--output_folder"):
            output_folder = arg

    if excel_file is None:
        help_man()
    if output_folder is None:
        output_folder = f"./result{datetime.datetime.now():%Y%m%d}"
        os.makedirs(output_folder, exist_ok=True)

    folder = output_folder
    devices = load_excel(excel_file)
    multithreaded_execution(devices, num_threads, folder)


if __name__ == "__main__":
    main(sys.argv[1:])
