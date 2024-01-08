#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import argparse
import logging
import multiprocessing
import netmiko
import pandas as pd
import os
import datetime
from concurrent.futures import ThreadPoolExecutor

def load_excel(excel_file):
    df = pd.read_excel(excel_file, sheet_name="Sheet1")
    devices_info = df.to_dict(orient="records")
    return devices_info

def execute_commands(mdev):
    ip = mdev["host"]
    user = mdev["username"]
    dev_type = mdev["device_type"]
    passwd = mdev["password"]
    devsecret = mdev["secret"]
    read_time = mdev.get("readtime", 10)
    cmds = list(mdev["mult_command"].split(";"))

    try:
        net_devices = {
            "device_type": dev_type,
            "host": ip,
            "username": user,
            "password": passwd,
            "secret": devsecret,
            "read_timeout_override": read_time,
        }
        net_connect = netmiko.ConnectHandler(**net_devices)

        with net_connect:
            if dev_type == "paloalto_panos":
                cmd_out = net_connect.send_multiline(cmds, expect_string=r">", cmd_verify=False)
            elif dev_type in ("huawei", "huawei_telnet", "hp_comware", "hp_comware_telnet"):
                cmd_out = net_connect.send_multiline(cmds, cmd_verify=False)
            elif mdev['secret'] != "":
                net_connect.enable()
                cmd_out = net_connect.send_multiline(cmds, cmd_verify=False)
            else:
                cmd_out = net_connect.send_multiline(cmds, cmd_verify=False)

        output_dir = args.output
        if not output_dir:
            output_dir = f"./result{datetime.datetime.now():%Y%m%d}"

        os.makedirs(output_dir, exist_ok=True)
        with open(os.path.join(output_dir, f"{ip}.txt"), "w", encoding="utf-8") as tmp_fle:
            tmp_fle.write(cmd_out + "\n")
        logging.info(f"{ip} 执行成功")
        return cmd_out

    except netmiko.exceptions.NetmikoAuthenticationException:
        with open("登录失败列表", "a", encoding="utf-8") as failed_ip:
            failed_ip.write(f"{ip} 用户名密码错误\n")
        logging.error(f"{ip} 用户名密码错误")
    except netmiko.exceptions.NetmikoTimeoutException:
        with open("登录失败列表", "a", encoding="utf-8") as failed_ip:
            failed_ip.write(f"{ip} 登录超时\n")
        logging.error(f"{ip} 登录超时")

    return None

def multithreaded_execution(devices, num_threads):
    with ThreadPoolExecutor(num_threads) as pool:
        pool.map(execute_commands, devices)

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("-i", "--input", help="Path to the Excel file containing device information", required=True)
    parser.add_argument(
        "-t", "--threads", help="Number of threads to use for parallel execution", type=int, default=4
    )
    parser.add_argument(
        "-o", "--output", help="Directory to save the command outputs. If not specified, a date-based directory will be created in the current directory."
    )
    args = parser.parse_args()

    logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')

    devices = load_excel(args.input)
    multithreaded_execution(devices, args.threads)

if __name__ == "__main__":
    main()
