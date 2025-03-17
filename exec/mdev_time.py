#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import netmiko
import openpyxl
import getopt
import os
import datetime
import sys
from typing import List, Dict
from concurrent.futures import ThreadPoolExecutor
from threading import Lock
from tqdm import tqdm

# 禁用颜色输出
os.environ["NO_COLOR"] = "1"
# 全局线程锁
write_lock = Lock()


def sanitize_filename(name: str) -> str:
    """清理非法文件名字符"""
    invalid_chars = r'<>:"/\|?*'
    return ''.join(c for c in name if c not in invalid_chars).strip()[:50]

def load_excel(excel_file: str) -> List[Dict]:
    """加载并验证设备信息"""
    required_fields = ['host', 'username', 'password', 'device_type']
    devices_info = []
    
    try:
        workbook = openpyxl.load_workbook(excel_file)
        sheet = workbook.active
        
        # 校验表头
        header = [str(cell.value).strip().lower() for cell in sheet[1]]
        missing = [f for f in required_fields if f not in header]
        if missing:
            print(f"缺少必要列: {', '.join(missing)}")
            sys.exit(1)
            
        # 处理数据行
        for row_idx, row in enumerate(sheet.iter_rows(min_row=2, values_only=True), start=2):
            device_data = {k: str(v).strip() if v else "" for k, v in zip(header, row)}
            
            # 数据验证
            missing_data = [f for f in required_fields if not device_data.get(f)]
            if missing_data:
                print(f"第{row_idx}行缺少数据: {', '.join(missing_data)}")
                sys.exit(1)
                
            devices_info.append(device_data)
            
        return devices_info
        
    except FileNotFoundError:
        print(f"[错误] 文件不存在: {excel_file}")
        sys.exit(1)
    except Exception as e:
        print(f"[错误] 读取失败: {str(e)}")
        sys.exit(1)

def execute_commands(device: Dict) -> str:
    """执行设备命令"""
    ip = device["host"]
    dev_type = device["device_type"]



    try:
        # 命令处理
        cmds = [cmd.strip() for cmd in device.get("mult_command", "").split(";") if cmd.strip()]
        if not cmds:
            print(f"{ip} [警告] 无有效命令")
            return None
            
        # 连接参数
        conn_params = {
            "device_type": dev_type,
            "host": ip,
            "username": device["username"],
            "password": device["password"],
            "secret": device.get("secret", ""),
            "read_timeout_override": int(device.get("readtime", 10)),
            "fast_cli": False
        }
        
        # 建立连接
        with netmiko.ConnectHandler(**conn_params) as conn:

            # 特权模式处理
            if device.get("secret"):
                conn.enable()
                output = conn.send_multiline(cmds, cmd_verify=False)
            # 执行命令
            elif dev_type == "paloalto_panos":
                output = conn.send_multiline(cmds, expect_string=r">", cmd_verify=False)
            else:
                output = conn.send_multiline(cmds, cmd_verify=False)

            # 获取主机名
            prompt = conn.find_prompt()
            hname = sanitize_filename(prompt.strip('#<>[]*:?'))
            
            # 保存结果
            date_str = datetime.datetime.now().strftime('%Y%m%d')
            output_dir = f"./result_{date_str}"
            os.makedirs(output_dir, exist_ok=True)
            
            filename = f"{sanitize_filename(ip)}_{hname}.txt"
            with write_lock:
                with open(os.path.join(output_dir, filename), "w", encoding="utf-8") as f:
                    f.write(f"=== 设备 {ip} 执行结果 ===\n{output}\n")
                    
            print(f"{ip} [成功] 执行完成")
            return output
            
    except netmiko.exceptions.NetmikoAuthenticationException as e:
        error_msg = f"{ip} [错误] 认证失败: {str(e)}"
    except netmiko.exceptions.NetmikoTimeoutException as e:
        error_msg = f"{ip} [错误] 连接超时: {str(e)}"
    except Exception as e:
        error_msg = f"{ip} [错误] 未知错误: {str(e)}"
        
    # 错误记录
    with write_lock:
        with open("error_log.txt", "a", encoding="utf-8") as f:
            f.write(f"{datetime.datetime.now()}|{error_msg}\n")
    print(error_msg)
    return None

def batch_execute(devices: List[Dict], max_workers: int = 4):
    """批量执行任务"""
    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        try:
            results = list(tqdm(executor.map(execute_commands, devices), 
                              total=len(devices),
                              desc="执行进度",
                              unit="台",
                              bar_format="{l_bar}{bar}| {n_fmt}/{total_fmt}"))
                              
            success = sum(1 for r in results if r is not None)
            print(f"\n执行完成: 成功 {success} 台 | 失败 {len(devices)-success} 台")
        except KeyboardInterrupt:
            print("\n正在安全终止程序...")
            executor.shutdown(wait=False)
            sys.exit(1)

def main(argv):
    """主函数"""
    usage = """
网络设备批量配置工具 v2.1

使用方法:
  connexec.py -i <设备清单.xlsx> [-t <并发数>]

参数说明:
  -i, --input    必需  输入Excel文件路径
  -t, --threads  可选  并发线程数 (默认:4, 范围1-20)

示例:
  connexec.py -i devices.xlsx -t 10
"""
    
    try:
        opts, _ = getopt.getopt(argv, "hi:t:", ["help", "input=", "threads="])
    except getopt.GetoptError:
        print(usage)
        sys.exit(2)
        
    excel_file = ""
    num_threads = 4
    
    for opt, arg in opts:
        if opt in ("-h", "--help"):
            print(usage)
            sys.exit()
        elif opt in ("-i", "--input"):
            excel_file = arg
        elif opt in ("-t", "--threads"):
            num_threads = max(1, min(int(arg), 20))
            
    if not excel_file:
        print(usage)
        sys.exit(2)
        
    try:
        devices = load_excel(excel_file)
        print(f"成功加载设备: {len(devices)} 台")
        batch_execute(devices, num_threads)
    except KeyboardInterrupt:
        print("\n操作已取消")
        sys.exit(0)

if __name__ == "__main__":
    main(sys.argv[1:])
