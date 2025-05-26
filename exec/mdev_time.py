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
import time
from typing import List, Dict, Optional, Tuple
from concurrent.futures import ThreadPoolExecutor, as_completed
from threading import Lock
import encodings.idna
from tqdm import tqdm
from netmiko import NetmikoTimeoutException, NetmikoAuthenticationException

# 环境配置
os.environ["NO_COLOR"] = "1"
write_lock = Lock()
DEFAULT_THREADS = min(900, max(4, (os.cpu_count() or 4)))

# **设备类型映射和配置**
DEVICE_TYPE_MAPPING = {
    'huawei': 'huawei',
    'huawei_vrpv8': 'huawei_vrpv8',
    'cisco': 'cisco_ios',
    'cisco_ios': 'cisco_ios',
    'cisco_xe': 'cisco_xe',
    'cisco_asa': 'cisco_asa',
    'hp': 'hp_comware',
    'hp_comware': 'hp_comware',
    'hp_procurve': 'hp_procurve',
    'h3c': 'hp_comware',  # H3C使用comware协议
    'h3c_comware': 'hp_comware',
    'ruckus': 'ruckus_fastiron',
    'ruckus_icx': 'ruckus_fastiron',
    'ruckus_fastiron': 'ruckus_fastiron',
    'paloalto': 'paloalto_panos',
    'panos': 'paloalto_panos',
    'fortinet': 'fortinet',
    'fortigate': 'fortinet',
    'fortios': 'fortinet'
}

# **设备特定配置参数**
DEVICE_CONFIGS = {
    'huawei': {
        'timeout': 30,
        'banner_timeout': 15,
        'auth_timeout': 10,
        'fast_cli': False,
        'session_timeout': 60
    },
    'huawei_vrpv8': {
        'timeout': 30,
        'banner_timeout': 15,
        'auth_timeout': 10,
        'fast_cli': False,
        'session_timeout': 60
    },
    'cisco_ios': {
        'timeout': 25,
        'banner_timeout': 15,
        'auth_timeout': 10,
        'fast_cli': False,
        'session_timeout': 60
    },
    'cisco_xe': {
        'timeout': 25,
        'banner_timeout': 15,
        'auth_timeout': 10,
        'fast_cli': False,
        'session_timeout': 60
    },
    'cisco_asa': {
        'timeout': 30,
        'banner_timeout': 20,
        'auth_timeout': 15,
        'fast_cli': False,
        'session_timeout': 60
    },
    'hp_comware': {
        'timeout': 25,
        'banner_timeout': 15,
        'auth_timeout': 10,
        'fast_cli': False,
        'session_timeout': 60
    },
    'hp_procurve': {
        'timeout': 25,
        'banner_timeout': 15,
        'auth_timeout': 10,
        'fast_cli': False,
        'session_timeout': 60
    },
    'ruckus_fastiron': {
        'timeout': 30,
        'banner_timeout': 20,
        'auth_timeout': 15,
        'fast_cli': False,
        'session_timeout': 60
    },
    'paloalto_panos': {
        'timeout': 45,
        'banner_timeout': 30,
        'auth_timeout': 20,
        'fast_cli': False,
        'session_timeout': 120
    },
    'fortinet': {
        'timeout': 30,
        'banner_timeout': 20,
        'auth_timeout': 15,
        'fast_cli': False,
        'session_timeout': 60
    }
}

def thread_initializer() -> None:
    """线程初始化（解决编码问题）"""
    import encodings.idna
    encodings.idna.__name__  # 防止被优化

def sanitize_filename(name: str) -> str:
    """生成安全文件名"""
    return re.sub(r'[\\/*?:"<>|]', '', name).strip()[:60]

def normalize_device_type(device_type: str) -> str:
    """**标准化设备类型**"""
    normalized = device_type.lower().strip()
    return DEVICE_TYPE_MAPPING.get(normalized, normalized)

def validate_device_data(device: Dict[str, str], row_idx: int) -> None:
    """验证设备数据完整性"""
    required = ['host', 'username', 'password', 'device_type']
    if missing := [f for f in required if not device.get(f)]:
        raise ValueError(f"Row {row_idx} 缺失字段: {', '.join(missing)}")
    
    # **验证设备类型是否支持**
    normalized_type = normalize_device_type(device['device_type'])
    if normalized_type not in DEVICE_CONFIGS:
        print(f"[WARN] Row {row_idx}: 未知设备类型 '{device['device_type']}', 将使用默认配置")

def load_excel(excel_file: str, sheet_name: str = 'Sheet1') -> List[Dict[str, str]]:
    """加载Excel设备清单（线程安全+版本兼容）"""
    devices = []
    wb = None
    try:
        wb = openpyxl.load_workbook(excel_file, read_only=True)
        
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
            # **标准化设备类型**
            device['device_type'] = normalize_device_type(device['device_type'])
            devices.append(device)
            
        return devices
    except Exception as e:
        print(f"Excel处理失败: {str(e)}")
        sys.exit(1)
    finally:
        if wb:
            wb.close()

def get_device_config(device_type: str) -> Dict:
    """**获取设备特定配置**"""
    return DEVICE_CONFIGS.get(device_type, {
        'timeout': 30,
        'banner_timeout': 15,
        'auth_timeout': 10,
        'fast_cli': False,
        'session_timeout': 60
    })

def connect_device(device: Dict[str, str]) -> Optional[netmiko.BaseConnection]:
    """**优化的设备连接（支持多厂商适配）**"""
    device_type = device['device_type']
    device_config = get_device_config(device_type)
    
    params = {
        'device_type': device_type,
        'host': device['host'],
        'username': device['username'],
        'password': device['password'],
        'secret': device.get('secret', ''),
        'timeout': device_config['timeout'],
        'banner_timeout': device_config['banner_timeout'],
        'auth_timeout': device_config['auth_timeout'],
        'fast_cli': device_config['fast_cli'],
        'session_timeout': device_config['session_timeout'],
        'read_timeout_override': int(device.get('readtime', device_config['timeout']))
    }

    # **端口配置**
    if device.get('port'):
        params['port'] = int(device['port'])
    
    # **SSL配置（PaloAlto等）**
    if device_type in ['paloalto_panos']:
        params['use_keys'] = False
        params['allow_agent'] = False
    
    # **特殊认证配置**
    if device_type in ['fortinet']:
        # Fortinet可能需要特殊的SSH配置
        params['allow_agent'] = False
        params['use_keys'] = False

    # **动态生成设备专属日志路径**
    if device.get('debug'):
        debug_dir = os.path.join("debug_logs", datetime.datetime.now().strftime('%Y%m%d'))
        os.makedirs(debug_dir, exist_ok=True)
        log_file = f"{sanitize_filename(device['host'])}_{uuid.uuid4().hex[:6]}.log"
        params['session_log'] = os.path.join(debug_dir, log_file)

    # **多重连接尝试**
    max_retries = 2
    for attempt in range(max_retries + 1):
        try:
            conn = netmiko.ConnectHandler(**params)
            
            # **设备特定的后连接处理**
            if device_type in ['cisco_ios', 'cisco_xe', 'cisco_asa', 'hp_comware', 'ruckus_fastiron']:
                if params['secret']:
                    conn.enable()
            elif device_type in ['huawei', 'huawei_vrpv8']:
                # 华为设备可能需要特殊处理
                if params['secret']:
                    try:
                        conn.enable()
                    except:
                        # 有些华为设备不需要enable
                        pass
            elif device_type == 'paloalto_panos':
                # PAN-OS可能需要额外的初始化
                time.sleep(1)
            elif device_type == 'fortinet':
                # Fortinet可能需要特殊处理
                time.sleep(0.5)
            
            return conn
            
        except (NetmikoTimeoutException, NetmikoAuthenticationException) as e:
            if attempt < max_retries:
                print(f"[RETRY {attempt+1}] {device['host']}: {e.__class__.__name__}")
                time.sleep(2 ** attempt)  # 指数退避
                continue
            log_error(device['host'], f"{e.__class__.__name__}: {str(e)}")
        except Exception as e:
            if attempt < max_retries:
                print(f"[RETRY {attempt+1}] {device['host']}: Connection error")
                time.sleep(2 ** attempt)
                continue
            log_error(device['host'], f"连接异常: {str(e)}")
    
    return None

def execute_commands(device: Dict[str, str], config_set: bool) -> Optional[str]:
    """**执行命令并捕获输出（多厂商适配）**"""
    try:
        cmds = [c.strip() for c in device.get('mult_command', '').split(';') if c.strip()]
        if not cmds:
            print(f"{device['host']} [WARN] 无有效命令")
            return None

        if not (conn := connect_device(device)):
            return None

        with conn:
            # **获取设备主机名（多厂商适配）**
            device_type = device['device_type']
            try:
                if device_type in ['paloalto_panos']:
                    # PAN-OS使用不同的提示符
                    prompt = conn.find_prompt().strip()
                    hostname_match = re.search(r'(\S+?)[@#>]', prompt)
                elif device_type == 'fortinet':
                    # Fortinet设备提示符
                    prompt = conn.find_prompt().strip()
                    hostname_match = re.search(r'$(\S+?)$', prompt) or re.search(r'(\S+?)[#>]', prompt)
                else:
                    # 通用设备提示符
                    prompt = conn.find_prompt().strip()
                    hostname_match = re.search(r'\S*?([\w.-]+)[#>]', prompt)
                
                device['hostname'] = hostname_match.group(1) if hostname_match else 'unknown'
            except:
                device['hostname'] = 'unknown'

            # **执行命令（根据设备类型选择方法）**
            all_output = []
            
            if config_set:
                # 配置模式
                if device_type in ['paloalto_panos']:
                    # PAN-OS需要特殊的配置模式处理
                    for cmd in cmds:
                        output = conn.send_command(cmd, expect_string=r'[#>]')
                        all_output.append(f"Command: {cmd}\n{output}")
                else:
                    # 标准配置模式
                    output = conn.send_config_set(cmds, cmd_verify=False)
                    all_output.append(output)
            else:
                # 非配置模式
                for cmd in cmds:
                    if device_type in ['huawei', 'huawei_vrpv8']:
                        # 华为设备特殊处理
                        output = conn.send_command(cmd, delay_factor=2)
                    elif device_type in ['paloalto_panos']:
                        # PAN-OS特殊处理
                        output = conn.send_command(cmd, expect_string=r'[#>]', delay_factor=3)
                    elif device_type == 'fortinet':
                        # Fortinet特殊处理
                        output = conn.send_command(cmd, delay_factor=2)
                    elif device_type in ['ruckus_fastiron']:
                        # Ruckus ICX特殊处理
                        output = conn.send_command(cmd, delay_factor=2)
                    else:
                        # 标准处理
                        output = conn.send_command(cmd)
                    
                    all_output.append(f"Command: {cmd}\n{output}")

            return "\n\n".join(all_output)
            
    except Exception as e:
        log_error(device['host'], f"执行异常 ({device.get('device_type', 'unknown')}): {str(e)}")
        return None

def save_result(ip: str, hostname: str, output: str, dest_path: str, device_type: str = '') -> None:
    """保存执行结果（原子写入）"""
    date_str = datetime.datetime.now().strftime('%Y%m%d')
    output_dir = os.path.join(dest_path, f"result_{date_str}")
    os.makedirs(output_dir, exist_ok=True)

    filename = f"{sanitize_filename(ip)}_{hostname or 'unknown'}_{device_type}.txt"
    timestamp = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    content = f"=== {ip} ({hostname}) [{device_type}] 执行结果 ===\n时间: {timestamp}\n\n{output}"

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
    success_count = 0
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
                            save_result(
                                dev['host'], 
                                dev.get('hostname', 'unknown'), 
                                result, 
                                destination,
                                dev.get('device_type', 'unknown')
                            )
                            success_count += 1
                    except Exception as e:
                        log_error(dev['host'], str(e))
                    finally:
                        progress.update(1)
                progress.close()
                print(f"\n**完成**: 成功 {success_count}/{len(devices)} 台设备")
            except KeyboardInterrupt:
                progress.close()
                executor.shutdown(wait=False, cancel_futures=True)
                raise
    except KeyboardInterrupt:
        sys.exit(0)

def parse_args() -> argparse.Namespace:
    """命令行参数解析"""
    parser = argparse.ArgumentParser(
        description="网络设备批量管理工具 v4.1 (支持Huawei/Cisco/HP/H3C/Ruckus/PaloAlto/Fortinet)", 
        add_help=False, 
        formatter_class=argparse.RawTextHelpFormatter
    )
    parser.add_argument('-i', '--input', required=True, help='设备清单Excel路径')
    parser.add_argument('-t', '--threads', type=int, default=DEFAULT_THREADS, help=f'并发线程数 (默认: {DEFAULT_THREADS})')
    parser.add_argument('-cs', '--config_set', action='store_true', help='使用配置模式发送命令')
    parser.add_argument('-d', '--destination', default='./', help='结果保存路径 (默认: 当前目录)')
    parser.add_argument('--debug', action='store_true', help='启用调试日志')
    parser.add_argument('-s', '--sheet', default='Sheet1', help='指定Excel工作表名称')
    
    if '--help' in sys.argv or '-h' in sys.argv:
        print(f"""
**网络设备批量管理工具 v4.1**

**支持设备类型**:
- **Huawei**: huawei, huawei_vrpv8
- **Cisco**: cisco, cisco_ios, cisco_xe, cisco_asa  
- **HP**: hp, hp_comware, hp_procurve
- **H3C**: h3c, h3c_comware (使用comware协议)
- **Ruckus**: ruckus, ruckus_icx, ruckus_fastiron
- **PaloAlto**: paloalto, panos, paloalto_panos
- **Fortinet**: fortinet, fortigate, fortios

**使用方法**:
  connexec -i <设备清单.xlsx> [-t 并发数]

**参数说明**:
  -i, --input        必需  Excel文件路径
  -t, --threads      可选  并发线程数（默认{DEFAULT_THREADS}）
  -cs, --config_set  可选  使用配置模式发送命令
  -d, --destination  可选  结果保存路径，默认: 当前目录
  -s, --sheet        可选  Excel工作表名称，默认: Sheet1
  --debug            可选  启用详细调试日志

**Excel格式示例**:
+---------------+----------+-------------+----------------+--------+----------+-------------------------+
|     host      | username |  password   |  device_type   | secret | readtime |      mult_command       |
+---------------+----------+-------------+----------------+--------+----------+-------------------------+
| 192.168.1.1   |  admin   | Cisco@123   |   cisco_ios    | enable |    15    | show version;show run   |
| 192.168.1.2   |  admin   | HuaWei@123  |   huawei       |        |    20    | display version;dis cur |
| 192.168.1.3   |  admin   | Pa@123      |   paloalto     |        |    30    | show system info        |
| 192.168.1.4   |  admin   | Forti@123   |   fortinet     |        |    25    | get system status       |
+---------------+----------+-------------+----------------+--------+----------+-------------------------+
""")
        sys.exit(0)

    return parser.parse_args()

def main() -> None:
    """主入口"""
    args = parse_args()
    
    if not os.path.exists(args.input):
        print(f"**错误**: 文件不存在 [{args.input}]")
        sys.exit(1)
        
    try:
        devices = load_excel(args.input, args.sheet)
        
        # 注入debug标志到所有设备
        if args.debug:
            for device in devices:
                device['debug'] = True
        
        # **显示设备类型统计**
        device_types = {}
        for device in devices:
            device_type = device['device_type']
            device_types[device_type] = device_types.get(device_type, 0) + 1
        
        print(f"**成功加载设备**: {len(devices)} 台 (工作表: {args.sheet})")
        print("**设备类型分布**:")
        for dt, count in device_types.items():
            print(f"  - {dt}: {count} 台")
        
        batch_execute(devices, args.config_set, args.threads, args.destination)
    except KeyboardInterrupt:
        print("\n**用户终止**")
        sys.exit(0)

if __name__ == "__main__":
    main()
