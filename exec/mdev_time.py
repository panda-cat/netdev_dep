#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import netmiko
import multiprocessing
import openpyxl
import getopt
import os
import datetime
import logging
from concurrent.futures import ThreadPoolExecutor
import sys
import time

# Setup logging
logging.basicConfig(filename='execution.log', level=logging.INFO, 
                    format='%(asctime)s - %(levelname)s - %(message)s')

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
        logging.error(f"Error: Excel file not found: {excel_file}")
        sys.exit(1)
    except Exception as e:
        logging.error(f"Error reading Excel file: {e}")
        sys.exit(1)
    return devices_info

def execute_commands(devices):
    ip = devices["host"]
    user = devices["username"]
    dev_type = devices["device_type"]
    passwd = devices["password"]
    secret = devices["secret"]
    read_time = devices.get("readtime", 10)
    cmds = list(devices["mult_command"].split(";"))

    attempts = 3  # Number of retry attempts
    for attempt in range(attempts):
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

            with net_connect:
                if dev_type == "paloalto_panos":
                    cmd_out = net_connect.send_multiline(cmds, expect_string=r">", cmd_verify=False)
                elif dev_type in ("huawei", "huawei_telnet", "hp_comware", "hp_comware_telnet"):
                    cmd_out = net_connect.send_multiline(cmds, cmd_verify=False)
                elif net_devices['secret'] != "":
                    net_connect.enable()
                    cmd_out = net_connect.send_multiline(cmds, cmd_verify=False)
                else:
                    cmd_out = net_connect.send_multiline(cmds, cmd_verify=False)

            output_dir = f"./result{datetime.datetime.now():%Y%m%d}"
            os.makedirs(output_dir, exist_ok=True)
            with open(os.path.join(output_dir, f"{ip}.txt"), "w", encoding="utf-8") as tmp_fle:
                tmp_fle.write(cmd_out + "\n")
            logging.info(f"{ip} 执行成功")
            return cmd_out

        except netmiko.exceptions.NetmikoAuthenticationException:
            logging.error(f"{ip} 用户名密码错误")
            with open("登录失败列表.txt", "a", encoding="utf-8") as failed_ip:
                failed_ip.write(f"{ip} 用户名密码错误\n")
        except netmiko.exceptions.NetmikoTimeoutException:
            logging.error(f"{ip} 登录超时")
            with open("登录失败列表.txt", "a", encoding="utf-8") as failed_ip:
                failed_ip.write(f"{ip} 登录超时\n")
        except Exception as e:
            logging.error(f"Error executing commands on {ip}: {e}")

        time.sleep(5)  # Wait before retrying

    return None

def multithreaded_execution(devices, num_threads):
    with ThreadPoolExecutor(num_threads) as pool:
        pool.map(execute_commands, devices)

def main(argv):
    try:
        opts, args = getopt.getopt(argv, "i:t:o:l:d:", ["input=", "threads=", "output=", "loglevel=", "device="])
    except getopt.GetoptError:
        print("Usage: connexec -i <excel_file> -t <num_threads default:4> -o <output_dir default:./result> -l <loglevel default:INFO> -d <device>")
        sys.exit(2)

    excel_file = ""
    num_threads = 4
    output_dir = "./result"
    loglevel = "INFO"
    single_device = None
    for opt, arg in opts:
        if opt in ("-i", "--input"):
            excel_file = arg
        elif opt in ("-t", "--threads"):
            num_threads = int(arg)
        elif opt in ("-o", "--output"):
            output_dir = arg
        elif opt in ("-l", "--loglevel"):
            loglevel = arg.upper()
        elif opt in ("-d", "--device"):
            single_device = arg

    logging.getLogger().setLevel(getattr(logging, loglevel, logging.INFO))

    devices = []
    if single_device:
        # Parsing single device details provided via command line
        device_details = single_device.split(",")
        device_keys = ["host", "username", "password", "device_type", "secret", "mult_command"]
        device_info = dict(zip(device_keys, device_details))
        devices.append(device_info)
    else:
        devices = load_excel(excel_file)
    
    multithreaded_execution(devices, num_threads)

if __name__ == "__main__":
    main(sys.argv[1:])
