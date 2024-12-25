#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import pandas as pd
import getopt
import os
import datetime
from nornir import InitNornir
from nornir.core.task import Task, Result
from nornir_netmiko import netmiko_send_command
import sys

def load_excel(excel_file):
    if not os.path.exists(excel_file):
        print(f"Error: Excel file not found at '{excel_file}'")
        sys.exit(2)  # Exit with a non-zero code to indicate an error
    try:
        df = pd.read_excel(excel_file, sheet_name="Sheet1")
        devices_info = df.to_dict(orient="records")
        return devices_info
    except Exception as e:
        print(f"Error: Failed to read or parse Excel file. {e}")
        sys.exit(2)

def execute_commands(devices, nr):
   ip = devices["host"]
   dev_type = devices["device_type"]
   cmds = list(devices["mult_command"].split(";"))
   secret = devices["secret"]
   read_time = devices.get("readtime", 10)

   def send_commands(task: Task, commands):
       output = ""
       for cmd in commands:
            result = task.run(
                task=netmiko_send_command,
                command_string=cmd,
            )
            output += result[0].result
       return Result(host=task.host, result=output)

   try:
       
       result = nr.run(task=send_commands, commands = cmds)
       output = result[ip].result
       output_dir = f"./result{datetime.datetime.now():%Y%m%d}"
       os.makedirs(output_dir, exist_ok=True)
       with open(os.path.join(output_dir, f"{ip}.txt"), "w", encoding="utf-8") as tmp_fle:
           tmp_fle.write(output + "\n")
       print(f"{ip} 执行成功")
       return output

   except Exception as e:
       with open("登录失败列表.txt", "a", encoding="utf-8") as failed_ip:
           failed_ip.write(f"{ip} 执行失败: {e}\n")
       print(f"{ip} 执行失败: {e}")

   return None


def multithreaded_execution(devices, num_threads):  # num_threads is still accepted but ignored here.
    nr = InitNornir(
        inventory={
            "hosts": {
                device["host"]: {
                    "hostname": device["host"],
                    "username": device["username"],
                    "password": device["password"],
                    "platform": device["device_type"],
                    "port": 22,
                    "data":{
                        "secret": device["secret"],
                        "read_time": device.get("readtime", 10)
                    }
                 }
                 for device in devices
            }
        }

    )

    # we simply invoke the execution method and pass in the Nornir object and list of devices, Nornir handles concurrency.
    for device in devices:
        execute_commands(device, nr)


def main(argv):
   try:
       opts, args = getopt.getopt(argv, "i:t:", ["input=", "threads="])
   except getopt.GetoptError:
       print("Usage: main.py -i <excel_file> -t <num_threads default:4>")
       sys.exit(2)

   excel_file = ""
   num_threads = 4 # this variable is not used.
   for opt, arg in opts:
       if opt in ("-i", "--input"):
           excel_file = arg
       elif opt in ("-t", "--threads"):
           num_threads = int(arg) # this variable is not used.


   devices = load_excel(excel_file)
   if devices:
       multithreaded_execution(devices, num_threads) # we still pass num_threads, but it is ignored.

if __name__ == "__main__":
   main(sys.argv[1:])
