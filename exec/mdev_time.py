#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import netmiko
import threading
import pandas as pd
import getopt
import os
import datetime

def load_excel(excel_file):
    df = pd.read_excel(excel_file,sheet_name="Sheet1")
    devices_info = df.to_dict(orient="records")
    return devices_info

def execute_commands(ip,user,dev_type,passwd,secret,read_time,cmds):
    try:
        net_devices = {
                'device_type': dev_type,
                'host': ip,
                'username': user,
                'password': passwd,
                'secret': secret,
                'read_timeout_override': read_time,
            }
        net_connect = netmiko.ConnectHandler(**net_devices)
        if net_devices['secret'] != "huawei" or net_devices['secret'] != "huawei_telnet":
           connect_dev.enable()
        if net_devices['secret'] != "hp_comware" or net_devices['secret'] != "hp_comware_telnet":
           connect_dev.enable()
        cmd_out = connect_dev.send_multiline(cmds)
        os.chdir(os.mkdir("./result"+'{0:%Y%m%d}'.format(datetime.datetime.now())))
        with open (ip + ".txt", "w",encoding="utf-8")  as tmp_fle:
             tmp_fle.write(cmd_out+'\n')
        print(ip + " 执行成功")
        return cmd_out
    except netmiko.exceptions.NetmikoAuthenticationException:
        with open("登录失败列表", "a", encoding="utf-8") as failed_ip:
             failed_ip.write(ip + "  用户名密码错误\n")
             print(ip + " 用户名密码错误")
             return None
    except netmiko.exceptions.NetmikoTimeoutException:
        with open("登录失败列表", "a", encoding="utf-8") as failed_ip:
             failed_ip.write(ip + "       登录超时\n")
             print(ip + " 登录超时")
             return None

def multithreaded_execution(devices, num_threads):
    threads = []
    for dev_info in devices:
        cmds = list(dev_info['mult_command'].split(";"))
        ip = dev_info["host"]
        user = dev_info["username"]
        dev_type = dev_info["device_type"]
        passwd = dev_info["password"]
        secret = dev_info["secret"]
        read_time = dev_info["readtime"]
        t = threading.Thread(target=execute_commands, args=(ip,user,dev_type,passwd,secret,read_time,cmds))
        threads.append(t)

    # Start threads
    for t in threads:
        t.start()

    # Wait for threads to complete
    for t in threads:
        t.join()

    return

def main(argv):
    try:
        opts, args = getopt.getopt(argv, "c:t:", ["excel=", "threads="])
    except getopt.GetoptError:
        print("Usage: program_network_devices.py -c <excel_file> -t <num_threads>")
        sys.exit(2)

    execl_file = ""
    num_threads = 4
    for opt, arg in opts:
        if opt in ("-c", "--excel"):
            execl_file = arg
        elif opt in ("-t", "--threads"):
            num_threads = int(arg)

    devices = load_execl(execl_file)
    #commands = ["show version", "show ip interface brief"]

    #multithreaded_execution(devices, commands, num_threads)
    multithreaded_execution(devices, num_threads)


if __name__ == "__main__":
    main(sys.argv[1:])
