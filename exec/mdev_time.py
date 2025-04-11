#!/usr/bin/env python3
# -*- coding: utf-8 -*-

# 预加载解决编码问题
import encodings.idna  # noqa: F401

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
import encodings.idna
from tqdm import tqdm
from netmiko import NetmikoTimeoutException, NetmikoAuthenticationException

# 环境配置
os.environ["NO_COLOR"] = "1"
write_lock = Lock()
DEFAULT_THREADS = min(900, max(4, (os.cpu_count() or 4)))  # 线程上限提升至900

def thread_initializer() -> None:
    """增强型线程初始化"""
    import encodings.idna
    # 显式调用确保模块加载
    encodings.idna.__name__  # 防止被优化
    encodings.idna.__file__  # 触发实际导入

def sanitize_filename(name: str) -> str:
    """安全文件名生成（保留特殊字符）"""
    clean_name = re.sub(r'[\\/*?:"<>|]', '', name).strip()
    return clean_name[:60]  # 延长文件名限制

def validate_device_data(device: Dict[str, str], row_idx: int) -> None:
    """设备数据验证"""
    required = ['host', 'username', 'password', 'device_type']
    if missing := [f for f in required if not device.get(f)]:
        raise ValueError(f"Row {row_idx} 缺失字段: {', '.join(missing)}")

def load_excel(excel_file: str) -> List[Dict[str, str]]:
    """修复后的Excel加载函数"""
    devices = []
    wb = None
    try:
        # 正确加载方式
        wb = openpyxl.load_workbook(excel_file, read_only=True)
        sheet = wb.active
        
        # 验证表头
        headers = [str(cell.value).lower().strip() for cell in sheet[1]]
        required = ['host', 'username', 'password', 'device_type']
        if missing := [f for f in required if f not in headers]:
            raise ValueError(f"缺少必要列: {', '.join(missing)}")

        # 处理数据行
        for row_idx, row in enumerate(sheet.iter_rows(min_row=2, values_only=True), 2):
            try:
                device = {headers[i]: str(cell).strip() if cell else "" for i, cell in enumerate(row)}
                validate_device_data(device, row_idx)
                devices.append(device)
            except ValueError as e:
                print(f"行 {row_idx} 数据错误: {str(e)}")
                sys.exit(1)
                
        return devices
        
    except Exception as e:
        print(f"Excel处理失败: {str(e)}")
        sys.exit(1)
    finally:
        if wb:  # 显式关闭工作簿
            wb.close()

def connect_device(device: Dict[str, str]) -> Optional[netmiko.BaseConnection]:
    """增强型设备连接"""
    params = {
        'device_type': device['device_type'],
        'host': device['host'],
        'username': device['username'],
        'password': device['password'],
        'secret': device.get('secret', ''),
        'session_log': 'netmiko.log' if device.get('debug') else None,
        'read_timeout_override': int(device.get('readtime', 20)),
        'fast_cli': False
    }
    
    try:
        conn = netmiko.ConnectHandler(**params)
        if params['secret']:
            conn.enable()
        return conn
    except (NetmikoTimeoutException, NetmikoAuthenticationException) as e:
        log_error(device['host'], f"{e.__class__.__name__}: {str(e)}")
    except Exception as e:
        log_error(device['host'], f"连接异常: {str(e)}")
    return None

def execute_commands(device: Dict[str, str], config_set: bool) -> Optional[str]:
    """命令执行（含主机名解析）"""
    try:
        if not (cmds := [c.strip() for c in device.get('mult_command', '').split(';') if c.strip()]):
            print(f"{device['host']} [WARN] 无有效命令")
            return None
            
        if not (conn := connect_device(device)):
            return None
            
        with conn:
            # 获取设备主机名
            try:
                prompt = conn.find_prompt().strip()
                for pattern in [r'\S+?([\w-]+)[#>]', r'$$(.*?)$$']:
                    if match := re.search(pattern, prompt):
                        device['hostname'] = sanitize_filename(match.group(1))
                        break
                else:
                    device['hostname'] = 'unknown'
            except Exception:
                device['hostname'] = 'unknown'

            # 执行命令
            output = conn.send_config_set(cmds, cmd_verify=False) if config_set else conn.send_multiline(
                cmds, 
                expect_string=r'>' if 'panos' in device['device_type'] else None
            )
            
            return output
            
    except Exception as e:
        log_error(device['host'], f"执行异常: {str(e)}")
        return None

def save_result(ip: str, hostname: str, output: str, dest_path: str) -> None:
    """增强型结果保存"""
    date_str = datetime.datetime.now().strftime('%Y%m%d')
    output_dir = os.path.join(dest_path, f"result_{date_str}")
    os.makedirs(output_dir, exist_ok=True)
    
    filename = f"{sanitize_filename(ip)}_{hostname or 'unknown'}.txt"
    content = f"=== {ip} ({hostname}) 执行结果 ===\n{output}"
    
    try:
        with tempfile.NamedTemporaryFile(
            mode='w', 
            encoding='utf-8',
            delete=False, 
            dir=output_dir
        ) as tmp_file:
            tmp_file.write(content)
            tmp_path = tmp_file.name
            
        os.rename(tmp_path, os.path.join(output_dir, filename))
    except OSError as e:
        log_error(ip, f"文件保存失败: {str(e)}")

def log_error(ip: str, error: str) -> None:
    """安全日志记录"""
    sanitized = re.sub(r'(password|secret)\s*=\s*\S+', r'\1=***', error, flags=re.I)
    log_line = f"{datetime.datetime.now():%Y-%m-%d %H:%M:%S} | {ip} | {sanitized}"
    
    with write_lock:
        with open("error.log", 'a', encoding='utf-8') as f:
            f.write(log_line + '\n')
        print(f"{ip} [ERROR] {sanitized}")

def batch_execute(
    devices: List[Dict[str, str]],
    config_set: bool,
    max_workers: int = DEFAULT_THREADS,
    destination: str = './'
) -> None:
    """增强型批量执行"""
    with ThreadPoolExecutor(
        max_workers=max_workers,
        initializer=thread_initializer
    ) as executor:
        try:
            futures = {executor.submit(execute_commands, dev, config_set): dev for dev in devices}
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
                        if (result := future.result()) is not None:
                            save_result(
                                dev['host'],
                                dev.get('hostname', 'unknown'),
                                result,
                                destination
                            )
                            success += 1
                    except Exception as e:
                        log_error(dev['host'], str(e))
                    finally:
                        pbar.update(1)
                        
            print(f"\n完成: 成功 {success}/{len(devices)}")
            
        except KeyboardInterrupt:
            print("\n安全终止中...")
            executor.shutdown(wait=False, cancel_futures=True)
            sys.exit(1)

def parse_args() -> argparse.Namespace:
    """命令行解析（解除线程限制）"""
    parser = argparse.ArgumentParser(
        description="网络设备批量管理工具 v3.0",
        add_help=False,
        formatter_class=argparse.RawTextHelpFormatter
    )
    
    parser.add_argument('-i', '--input', required=True, help='设备清单Excel路径')
    parser.add_argument('-t', '--threads', type=int, default=DEFAULT_THREADS, 
                       help=f'并发线程数 (默认: {DEFAULT_THREADS})')
    parser.add_argument('-cs', '--config_set', action='store_true', 
                       help='使用配置模式发送命令')
    parser.add_argument('-d', '--destination', default='./', 
                       help='结果保存路径 (默认: 当前目录)')
    parser.add_argument('--debug', action='store_true', 
                       help='启用调试日志')
    
    if '--help' in sys.argv or '-h' in sys.argv:
        print("""
使用方法:
  connexec -i <设备清单.xlsx> [-t 并发数]

参数说明:
  -i, --input        必需  Excel文件路径
  -t, --threads      可选  并发线程线程（最小值1，默认4）
  -cs, --config_set  可选  自动进入设备配置模式，并发送命令
  -d, --destination  可选  保存输出结果的目标目录路径，默认: 当前目录

示例Excel格式:
+-------------+----------+------------+--------------+--------+----------+------------------------+
|    host     | username |  password  | device_type  | secret | readtime |      mult_command      |
+-------------+----------+------------+--------------+--------+----------+------------------------+
| 192.168.1.1 |  admin   | Cisco@123  |   cisco_ios  | enable |    15    | show version;show run  |
| 10.10.1.1   |  huawei  | HuaWei@123 |   huawei     |        |    10    | display version;dis cur|
+-------------+----------+------------+--------------+--------+----------+------------------------+

支持平台列表参考:
https://github.com/ktbyers/netmiko/blob/develop/PLATFORMS.md
""")
        sys.exit(0)
        
    return parser.parse_args()

def main() -> None:
    """主入口"""
    args = parse_args()
    
    if not os.path.exists(args.input):
        print(f"错误: 文件不存在 [{args.input}]")
        sys.exit(1)
        
    try:
        devices = load_excel(args.input)
        print(f"成功加载设备: {len(devices)} 台")
        batch_execute(devices, args.config_set, args.threads, args.destination)
    except KeyboardInterrupt:
        print("\n用户终止")
        sys.exit(0)

if __name__ == "__main__":
    main()
