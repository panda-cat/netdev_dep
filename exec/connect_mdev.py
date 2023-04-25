#!/usr/bin/env python3
# -*- coding: utf-8 -*-

from concurrent.futures import ThreadPoolExecutor
import netmiko
import os
from threading import Lock
import pandas as pd

class net_dev():

    def __init__(self,excel_name):
        try :
            os.mkdir("./log")
        except:
            pass
        self.excel_name = excel_name
        self.list = [] # 空列表存储设备信息数据
        self.pool = ThreadPoolExecutor(10) # 初始化线程数量
        self.lock = Lock()  # 添加线程锁，避免写入数据丢失
        self.path = ("./log")  # 创建保存log路径
        #self.mult_config=[] # 创建列表，保存多条命令。用于批量执行命令

    def get_dev_info(self):
        # 获取sheet(设备信息)的dataframe.
        df = pd.read_excel(self.excel_name,sheet_name="Sheet1") # 读取excel的sheet1
        self.list = df.to_dict(orient="records")  # 将数据打印出来，已字典存储的列表数据
        #self.mult_config = list(df['mult_command'].split(";"))
        #mult_conf = df["mult_command"].values.tolist()  # 取一列的值生成列表
        print(self.list)

        # 获取sheet(CMD)的dataframe
        #df1 = pd.read_excel(self.excel_name,sheet_name="Sheet1")
        #result1 = df1.to_dict(orient="list")  # 将数据打印出来,将一列的数据存为一个字典
        #self.mult_config = result1["mult_command"]
        #print(self.mult_config)

    def mult_cmd_in(self,ip,user,dev_type,passwd,secret,cmds):
        try:
            devices = {
                'device_type': dev_type,  # 锐捷os:ruijie_os, 华三：hp_comware 中兴：zte_zxros
                'host': ip,
                'username': user,
                'password': passwd,
                'secret': secret,
            }

            connect_dev = netmiko.ConnectHandler(**devices)
            if devices['secret'] != "":
               connect_dev.enable()
            #for cmd in cmds:
            cmd_out = connect_dev.send_multiline(cmds)
            with open (ip + ".txt", "w",encoding="utf-8")  as tmp_fle:
                 tmp_fle.write(cmd_out+'\n')
            print(ip + " 执行成功")

        except netmiko.exceptions.NetmikoAuthenticationException:
            self.lock.acquire()
            with open("登录失败列表", "a", encoding="utf-8") as failed_ip:
                failed_ip.write(ip + "  用户名密码错误\n")
                print(ip + " 用户名密码错误")
            self.lock.release()
        except netmiko.exceptions.NetmikoTimeoutException:
            self.lock.acquire()
            with open("登录失败列表", "a", encoding="utf-8") as failed_ip:
                failed_ip.write(ip + "       登录超时\n")
                print(ip + " 登录超时")
            self.lock.release()

    def main(self):
        for dev_info in self.list:
            cmds = list(dev_info['mult_command'].split(";"))
            #print(dev_info)
            ip = dev_info["host"]
            #print(ip)
            user = dev_info["username"]
            dev_type = dev_info["device_type"]
            passwd = dev_info["password"]
            secret = dev_info["secret"]
            self.pool.submit(self.mult_cmd_in,ip,user,dev_type,passwd,secret,cmds)
        os.chdir(self.path)
        self.pool.shutdown(True)

yc_use = net_dev("inventory.xlsx")
yc_use.get_dev_info()
yc_use.main()
