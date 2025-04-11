#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import netmiko
import openpyxl
import argparse
import os
import datetime
import sys
import re
import uuid
import tempfile
from typing import List, Dict, Optional
from concurrent.futures import ThreadPoolExecutor, as_completed
from threading import Lock
from tqdm import tqdm
import encodings.idna
from netmiko import NetmikoTimeoutException, NetmikoAuthenticationException

# 环境配置
os.environ["NO_COLOR"] = "1"  # 禁用彩色输出
write_lock = Lock()           # 全局写入锁
DEFAULT_THREADS = min(500, max(4, (os.cpu_count() or 4)))  # 动态线程数

def thread_initializer() -> None:
    """线程初始化函数（解决编码问题）"""
    import encodings.idna  # noqa: F401
    # 显式调用确保模块加载
    encodings.idna.__name__  # 防止被优化
    encodings.idna.__file__  # 触发实际导入

def sanitize_filename(name: str) -> str:
    """生成安全文件名（带唯一标识）"""
    clean_name = re.sub(r'[<>:"/\\|?*]', '', name).strip()
    unique_id = uuid.uuid4().hex[:6]
    return f"{clean_name[:45]}_{unique_id}"

def validate_device_data(device: Dict[str, str], row_idx: int) -> None:
    """验证设备数据完整性"""
    required_fields = ['host', 'username', 'password', 'device_type']
    missing = [f for f in required_fields if not device.get(f)]
    if missing:
        raise ValueError(f"Row {row_idx} missing fields: {', '.join(missing)}")

def load_excel(excel_file: str) -> List[Dict[str, str]]:
    """加载并验证Excel设备信息（内存优化版）"""
    devices = []
    errors = []
    try:
        # 使用只读模式优化内存
        wb = openpyxl.load_workbook(excel_file, read_only=True)
        sheet = wb.active
        
        headers = [str(cell.value).lower().strip() for cell in sheet[1]]
        required = ['host', 'username', 'password', 'device_type']
        if any(f not in headers for f in required):
            raise ValueError(f"Missing required columns: {', '.join(required)}")

        for row_idx, row in enumerate(sheet.iter_rows(min_row=2, values_only=True), start=2):
            try:
                device = {k: str(v).strip() if v else "" for k, v in zip(headers, row)}
                validate_device_data(device, row_idx)
                devices.append(device)
            except ValueError as e:
                errors.append(str(e))
        
        if errors:
            print("Excel数据校验错误:\n" + "\n".join(errors))
            sys.exit(1)
            
        return devices
        
    except Exception as e:
        print(f"Excel处理失败: {str(e)}")
        sys.exit(1)
    finally:
        if 'wb' in locals():
            wb.close()

def connect_device(device: Dict[str, str]) -> Optional[netmiko.BaseConnection]:
    """建立设备连接（带异常分类处理）"""
    params = {
        'device_type': device['device_type'],
        'host': device['host'],
        'username': device['username'],
        'password': device['password'],
        'secret': device.get('secret', ''),
        'read_timeout_override': int(device.get('readtime', 15)),
        'fast_cli': False
    }
    
    conn = None
    try:
        conn = netmiko.ConnectHandler(**params)
        if device.get('secret'):
            conn.enable()
        return conn
    except NetmikoTimeoutException as e:
        log_error(device['host'], f"Connection timeout: {str(e)}")
    except NetmikoAuthenticationException as e:
        log_error(device['host'], f"Authentication failed: {str(e)}")
    except Exception as e:
        log_error(device['host'], f"Connection error: {str(e)}")
    finally:
        if conn is not None and not conn.is_alive():
            conn.disconnect()
    return None

def execute_commands(device: Dict[str, str], config_set: bool) -> Optional[str]:
    """执行设备命令主逻辑（增强异常处理）"""
    ip = device['host']
    try:
        cmds = [c.strip() for c in device.get('mult_command', '').split(';') if c.strip()]
        if not cmds:
            print(f"{ip} [WARN] No valid commands")
            return None
            
        if (conn := connect_device(device)) is None:
            return None
            
        with conn:
            # 发送初始化空命令清空缓冲区
            conn.send_command_timing('')

            if config_set:
                output = conn.send_config_set(cmds, cmd_verify=False)
            else:
                if device['device_type'] == 'paloalto_panos':
                    output = conn.send_multiline(cmds, expect_string=r">", cmd_verify=False)
                else:
                    output = conn.send_multiline(cmds, cmd_verify=False)

            # 基本结果验证
            if "Invalid input" in output:
                log_error(ip, "Command validation failed")
                return None
                
            return output
            
    except Exception as e:
        log_error(ip, f"Command execution error: {str(e)}")
        return None

def save_result(ip: str, prompt: str, output: str, dest_path: str) -> None:
    """原子化保存执行结果"""
    date_str = datetime.datetime.now().strftime('%Y%m%d')
    hname = sanitize_filename(prompt.strip('#<>[]*:?'))
    
    output_dir = os.path.join(dest_path, f"result_{date_str}")
    os.makedirs(output_dir, exist_ok=True)
    
    filename = f"{sanitize_filename(ip)}_{hname}.txt"
    final_path = os.path.join(output_dir, filename)
    content = f"=== 设备 {ip} 执行结果 ===\n{output}"

    try:
        # 使用临时文件确保原子写入
        with tempfile.NamedTemporaryFile(
            mode='w', 
            encoding='utf-8',
            delete=False,
            dir=output_dir
        ) as tmp_file:
            tmp_file.write(content)
            tmp_path = tmp_file.name
            
        os.rename(tmp_path, final_path)
    except OSError as e:
        log_error(ip, f"File save failed: {str(e)}")

def log_error(ip: str, error: str) -> None:
    """安全日志记录（过滤敏感信息）"""
    timestamp = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    # 过滤敏感信息
    sanitized_error = re.sub(
        r'(password|secret)\s*=\s*\S+', 
        r'\1=***', 
        error, 
        flags=re.IGNORECASE
    )
    msg = f"{timestamp} | {ip} | {sanitized_error}"
    
    with write_lock:
        with open("error_log.txt", 'a', encoding='utf-8') as f:
            f.write(msg + '\n')
        print(f"{ip} [ERROR] {sanitized_error}")

def batch_execute(
    devices: List[Dict[str, str]], 
    config_set: bool, 
    max_workers: int = DEFAULT_THREADS, 
    destination: str = './'
) -> None:
    """批量执行入口（改进进度显示）"""
    with ThreadPoolExecutor(
        max_workers=max_workers, 
        initializer=thread_initializer
    ) as executor:
        try:
            futures = {
                executor.submit(execute_commands, dev, config_set): dev
                for dev in devices
            }
            
            success = 0
            with tqdm(
                total=len(devices),
                desc="执行进度",
                unit="台",
                dynamic_ncols=True
            ) as pbar:
                for future in as_completed(futures):
                    dev = futures[future]
                    try:
                        result = future.result()
                        if result is not None:
                            save_result(
                                dev['host'], 
                                dev.get('prompt', 'unknown'),
                                result,
                                destination
                            )
                            success += 1
                    except Exception as e:
                        log_error(dev['host'], str(e))
                    finally:
                        pbar.update(1)
                        
            print(f"\n执行完成: 成功 {success} 台 | 失败 {len(devices)-success} 台")
            
        except KeyboardInterrupt:
            print("\n安全终止中...")
            executor.shutdown(wait=False, cancel_futures=True)
            sys.exit(1)

def parse_args() -> argparse.Namespace:
    """改进的命令行解析"""
    parser = argparse.ArgumentParser(
        description="网络设备批量配置工具 v2.5",
        add_help=False,
        formatter_class=argparse.RawTextHelpFormatter
    )
    
    parser.add_argument('-i', '--input', required=True, 
                       help='设备清单Excel文件路径')
    parser.add_argument('-t', '--threads', type=int, default=DEFAULT_THREADS,
                       help=f'并发线程数 (默认: {DEFAULT_THREADS}, 范围：1-100)')
    parser.add_argument('-cs', '--config_set', action='store_true',
                       help='使用配置模式发送命令')
    parser.add_argument('-d', '--destination', type=str, default='./',
                       help='结果保存路径 (默认: 当前目录)')
    parser.add_argument('-h', '--help', action='store_true',
                       help='显示帮助信息')

    help_text = """
使用方法:
  connexec -i <设备清单.xlsx> [-t 并发数]

参数说明:
  -i, --input        必需  Excel文件路径
  -t, --threads      可选  并发线程线程（最小值1，默认4）
  -cs, --config_set  可选  自动进入设备配置模式，并发送命令
  -d, --destination  可选  保存输出结果的目标目录路径

示例Excel格式:
+-------------+----------+------------+--------------+--------+----------+------------------------+
|    host     | username |  password  | device_type  | secret | readtime |      mult_command      |
+-------------+----------+------------+--------------+--------+----------+------------------------+
| 192.168.1.1 |  admin   | Cisco@123  |   cisco_ios  | enable |    15    | show version;show run  |
| 10.10.1.1   |  huawei  | HuaWei@123 |   huawei     |        |    10    | display version;dis cur|
+-------------+----------+------------+--------------+--------+----------+------------------------+

支持平台列表参考:
https://github.com/ktbyers/netmiko/blob/develop/PLATFORMS.md
"""
    
    if '--help' in sys.argv or '-h' in sys.argv:
        print(help_text)
        sys.exit(0)
        
    return parser.parse_args()

def main() -> None:
    """主入口函数"""
    # 编码调试代码
    try:
        "test".encode('idna')
    except LookupError:
        print("IDNA编码支持异常，请检查Python环境")
        sys.exit(1)
        
    args = parse_args()
    
    if not os.path.exists(args.input):
        print(f"错误: 文件不存在 [{args.input}]")
        sys.exit(1)
        
    if args.threads < 1 or args.threads > 8:
        print("错误: 线程数必须在1-8之间")
        sys.exit(1)

    try:
        devices = load_excel(args.input)
        print(f"成功加载设备: {len(devices)} 台")
        batch_execute(
            devices, 
            args.config_set, 
            args.threads, 
            args.destination
        )
    except KeyboardInterrupt:
        print("\n操作已中止")
        sys.exit(0)

if __name__ == "__main__":
    main()
