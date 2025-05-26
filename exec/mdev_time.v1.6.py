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
import encodings.idna
from tqdm import tqdm
from netmiko import NetmikoTimeoutException, NetmikoAuthenticationException

# 环境配置
os.environ["NO_COLOR"] = "1"
write_lock = Lock()
DEFAULT_THREADS = min(900, max(4, (os.cpu_count() or 4)))

def thread_initializer() -> None:
    """线程初始化（解决编码问题）"""
    import encodings.idna
    encodings.idna.__name__  # 防止被优化

def sanitize_filename(name: str) -> str:
    """生成安全文件名"""
    return re.sub(r'[\\/*?:"<>|]', '', name).strip()[:60]

def validate_device_data(device: Dict[str, str], row_idx: int) -> None:
    """验证设备数据完整性"""
    required = ['host', 'username', 'password', 'device_type']
    if missing := [f for f in required if not device.get(f)]:
        raise ValueError(f"Row {row_idx} 缺失字段: {', '.join(missing)}")

def load_excel(excel_file: str, sheet_name: str = 'Sheet1') -> List[Dict[str, str]]:
    """加载Excel设备清单（线程安全+版本兼容）"""
    devices = []
    wb = None
    try:
        # 方式1：直接加载（兼容旧版）
        wb = openpyxl.load_workbook(excel_file, read_only=True)
        
        # 方式2：如果升级到openpyxl>=3.0可改用：
        # with openpyxl.load_workbook(excel_file, read_only=True) as wb:
        
        if sheet_name not in wb.sheetnames:
            raise ValueError(f"工作表 '{sheet_name}' 不存在")
        sheet = wb[sheet_name]
        
        headers = [str(cell.value).lower().strip() for cell in sheet[1]]
        required = ['host', 'username', 'password', 'device_type']
        if missing := [f for f in required if f not in headers]:
            raise ValueError(f"缺少必要列: {', '.join(missing)}")

        for row_idx, row in enumerate(sheet.iter_rows(min_row=2, values_only=True), 2):
            device = {headers[i]: str(cell).strip() if cell else "" for i, cell in enumerate(row)}
            validate_device_data(device, row_idx)
            devices.append(device)
            
        return devices
    except Exception as e:
        print(f"Excel处理失败: {str(e)}")
        sys.exit(1)
    finally:
        if wb:  # 确保资源释放
            wb.close()

def connect_device(device: Dict[str, str]) -> Optional[netmiko.BaseConnection]:
    """设备连接（自动生成独立日志文件）"""
    params = {
        'device_type': device['device_type'],
        'host': device['host'],
        'username': device['username'],
        'password': device['password'],
        'secret': device.get('secret', ''),
        'read_timeout_override': int(device.get('readtime', 20)),
        'fast_cli': False
    }

    # 动态生成设备专属日志路径
    if device.get('debug'):
        debug_dir = os.path.join("debug_logs", datetime.datetime.now().strftime('%Y%m%d'))
        os.makedirs(debug_dir, exist_ok=True)
        log_file = f"{sanitize_filename(device['host'])}_{uuid.uuid4().hex[:6]}.log"
        params['session_log'] = os.path.join(debug_dir, log_file)

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
    """执行命令并捕获输出"""
    try:
        cmds = [c.strip() for c in device.get('mult_command', '').split(';') if c.strip()]
        if not cmds:
            print(f"{device['host']} [WARN] 无有效命令")
            return None

        if not (conn := connect_device(device)):
            return None

        with conn:
            # 获取设备主机名
            prompt = conn.find_prompt().strip()
            device['hostname'] = (
                re.search(r'\S+?([\w-]+)[#>]', prompt).group(1) 
                if re.search(r'\S+?([\w-]+)[#>]', prompt) 
                else 'unknown'
            )

            # 执行命令
            return (
                conn.send_config_set(cmds, cmd_verify=False) 
                if config_set 
                else conn.send_multiline(cmds)
            )
    except Exception as e:
        log_error(device['host'], f"执行异常: {str(e)}")
        return None

def save_result(ip: str, hostname: str, output: str, dest_path: str) -> None:
    """保存执行结果（原子写入）"""
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
    """安全记录错误日志"""
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
    """批量执行（带优雅终止）"""
    try:
        with ThreadPoolExecutor(
            max_workers=max_workers,
            initializer=thread_initializer
        ) as executor:
            futures = {executor.submit(execute_commands, dev, config_set): dev for dev in devices}
            progress = tqdm(total=len(devices), desc="执行进度", unit="台")

            try:
                for future in as_completed(futures):
                    dev = futures[future]
                    try:
                        if (result := future.result()) is not None:
                            save_result(dev['host'], dev.get('hostname', 'unknown'), result, destination)
                    except Exception as e:
                        log_error(dev['host'], str(e))
                    finally:
                        progress.update(1)
                progress.close()
                print(f"\n完成: 成功 {sum(1 for f in futures if f.result() is not None)}/{len(devices)}")
            except KeyboardInterrupt:
                progress.close()
                executor.shutdown(wait=False, cancel_futures=True)
                raise
    except KeyboardInterrupt:
        sys.exit(0)

def parse_args() -> argparse.Namespace:
    """命令行参数解析"""
    parser = argparse.ArgumentParser(description="网络设备批量管理工具 v4.0", add_help=False, formatter_class=argparse.RawTextHelpFormatter)
    parser.add_argument('-i', '--input', required=True, help='设备清单Excel路径')
    parser.add_argument('-t', '--threads', type=int, default=DEFAULT_THREADS, help=f'并发线程数 (默认: {DEFAULT_THREADS})')
    parser.add_argument('-cs', '--config_set', action='store_true', help='使用配置模式发送命令')
    parser.add_argument('-d', '--destination', default='./', help='结果保存路径 (默认: 当前目录)')
    parser.add_argument('--debug', action='store_true', help='启用调试日志')
    parser.add_argument('-s', '--sheet', default='Sheet1', help='指定Excel工作表名称')
    if '--help' in sys.argv or '-h' in sys.argv:
        print("""
使用方法:
  connexec -i <设备清单.xlsx> [-t 并发数]

参数说明:
  -i, --input        必需  Excel文件路径
  -t, --threads      可选  并发线程线程（最小值1，默认4）
  -cs, --config_set  可选  自动进入设备配置模式，并发送命令
  -d, --destination  可选  保存输出结果的目标目录路径，默认: 当前目录
  -s, --sheet        可选  指定excel中的sheet名称，默认: Sheet1

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
        devices = load_excel(args.input, args.sheet)
        
        # 注入debug标志到所有设备
        if args.debug:
            for device in devices:
                device['debug'] = True
        
        print(f"成功加载设备: {len(devices)} 台 (工作表: {args.sheet})")
        batch_execute(devices, args.config_set, args.threads, args.destination)
    except KeyboardInterrupt:
        print("\n用户终止")
        sys.exit(0)

if __name__ == "__main__":
    main()
