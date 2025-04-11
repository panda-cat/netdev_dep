#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import netmiko
import openpyxl
import argparse
import os
import datetime
import sys
import encodings.idna  # 关键预加载
from typing import List, Dict
from concurrent.futures import ThreadPoolExecutor
from threading import Lock
from tqdm import tqdm

# 环境配置
os.environ["NO_COLOR"] = "1"  # 禁用彩色输出
write_lock = Lock()           # 全局写入锁

def thread_initializer():
    """线程初始化函数（解决编码问题）"""
    import encodings.idna
    encodings.idna  # 防止被优化

def sanitize_filename(name: str) -> str:
    """生成安全文件名"""
    import re
    clean_name = re.sub(r'[<>:"/\\|?*]', '', name).strip()
    return clean_name[:50]  # 限制长度

def validate_device_data(device: Dict, row_idx: int):
    """验证设备数据完整性"""
    required = ['host', 'username', 'password', 'device_type']
    missing = [f for f in required if not device.get(f)]
    if missing:
        print(f"第{row_idx}行缺少字段: {', '.join(missing)}")
        sys.exit(1)

def load_excel(excel_file: str) -> List[Dict]:
    """加载并验证Excel设备信息"""
    devices = []
    try:
        wb = openpyxl.load_workbook(excel_file)
        sheet = wb.active
        
        # 解析表头
        headers = [str(cell.value).lower().strip() for cell in sheet[1]]
        required = ['host', 'username', 'password', 'device_type']
        if any(f not in headers for f in required):
            print(f"缺少必要列: {', '.join(required)}")
            sys.exit(1)
        
        # 处理数据行
        for idx, row in enumerate(sheet.iter_rows(min_row=2, values_only=True), 2):
            device = {k: str(v).strip() if v else "" for k, v in zip(headers, row)}
            validate_device_data(device, idx)
            devices.append(device)
            
        return devices
        
    except Exception as e:
        print(f"Excel处理失败: {str(e)}")
        sys.exit(1)

def connect_device(device: Dict) -> netmiko.BaseConnection:
    """建立设备连接"""
    params = {
        'device_type': device['device_type'],
        'host': device['host'],
        'username': device['username'],
        'password': device['password'],
        'secret': device.get('secret', ''),
        'read_timeout_override': int(device.get('readtime', 10)),
        'fast_cli': False
    }
    
    try:
        conn = netmiko.ConnectHandler(**params)
        if device.get('secret'):
            conn.enable()
        return conn
    except Exception as e:
        log_error(device['host'], str(e))
        return None

def execute_commands(device: Dict, config_set: bool) -> str:
    """执行设备命令主逻辑"""
    ip = device['host']
    
    try:
        cmds = [c.strip() for c in device.get('mult_command', '').split(';') if c.strip()]
        if not cmds:
            print(f"{ip} [警告] 无有效命令")
            return None
            
        conn = connect_device(device)
        if not conn:
            return None
            
        with conn:
            conn.send_command_timing('')

            # 根据 config_set 使用不同的命令发送方法
            if config_set:
                output = conn.send_config_set(cmds, cmd_verify=False)
            else:
                if device['device_type'] == 'paloalto_panos':
                    output = conn.send_multiline(cmds, expect_string=r">", cmd_verify=False)
                else:
                    output = conn.send_multiline(cmds, cmd_verify=False)

            return output
            
    except Exception as e:
        log_error(ip, str(e))
        return None

def save_result(ip: str, prompt: str, output: str, dest_path: str):
    """保存执行结果"""
    date_str = datetime.datetime.now().strftime('%Y%m%d')
    hname = sanitize_filename(prompt.strip('#<>[]*:?'))
    
    output_dir = os.path.join(dest_path, f"result_{date_str}")
    os.makedirs(output_dir, exist_ok=True)
    
    filename = f"{sanitize_filename(ip)}_{hname}.txt"
    content = f"=== 设备 {ip} 执行结果 ===\n{output}"

    with write_lock:
        with open(os.path.join(output_dir, filename), 'w', encoding='utf-8') as f:
            f.write(content)

def log_error(ip: str, error: str):
    """统一错误日志记录"""
    timestamp = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    msg = f"{timestamp} | {ip} | {error}"
    
    with write_lock:
        with open("error_log.txt", 'a', encoding='utf-8') as f:
            f.write(msg + '\n')
    print(f"{ip} [错误] {error}")

def batch_execute(devices: List[Dict], config_set: bool, max_workers: int = 4, destination: str = './'):
    """批量执行入口"""
    with ThreadPoolExecutor(max_workers=max_workers, initializer=thread_initializer) as executor:
        try:
            results = list(tqdm(
                executor.map(lambda dev: execute_commands(dev, config_set), devices),
                total=len(devices),
                desc="执行进度",
                unit="台",
            ))
            success = sum(1 for r in results if r is not None)
            print(f"\n执行完成: 成功 {success} 台 | 失败 {len(devices)-success} 台")
        except KeyboardInterrupt:
            print("\n正在安全终止...")
            executor.shutdown(wait=False)
            sys.exit(1)

def parse_args():
    """解析命令行参数"""
    parser = argparse.ArgumentParser(description="网络设备批量配置工具 v2.4")
    parser.add_argument('-i', '--input', required=True, help='Excel文件路径')
    parser.add_argument('-t', '--threads', type=int, default=4, help='并发线程数')
    parser.add_argument('-cs', '--config_set', action='store_true', help='使用使用配置模式发送命令')
    parser.add_argument('-d', '--destination', type=str, default='./', help='结果保存路径')
    
    # 添加帮助说明
    parser.add_argument('-h', '--help', action='help', 
        help="""功能特性:
- 多线程安全执行
- 编码问题修复
- 完善的错误处理
- 结果自动归档

使用方法:
  connexec -i <设备清单.xlsx> [-t 并发数] [-d 保存结果的目标文件夹路径] [-cs 自动进入配置模式并发送命令]

参数说明:
  -i, --input        必需  Excel文件路径
  -t, --threads      可选  并发线程数（最小值1，默认4）
  -d, --destination  可选  保存结果的目标文件夹路径，否则保存于当前文件夹
  -cs, --config_set  可选  使用配置模式发送命令

示例excel模板:
  host          username  password    device_type  secret   readtime  mult_command
  192.168.1.1   admin     Cisco@123   cisco_ios    enable   15        show version;show run
  10.10.1.1     huawei    HuaWei@123  huawei                10        display version;dis cur

关于netmiko支持平台:
  https://github.com/ktbyers/netmiko/blob/develop/PLATFORMS.md
""")
    return parser.parse_args()

def main():
    """命令行入口"""
    args = parse_args()

    excel_file = args.input
    threads = args.threads
    config_set = args.config_set
    destination = args.destination

    if not os.path.exists(excel_file):
        print("错误: 输入的Excel文件不存在。")
        sys.exit(1)

    try:
        devices = load_excel(excel_file)
        print(f"已加载设备: {len(devices)} 台")
        batch_execute(devices, config_set, threads, destination)
    except KeyboardInterrupt:
        print("\n用户终止操作")
        sys.exit(0)

if __name__ == "__main__":
    main()
