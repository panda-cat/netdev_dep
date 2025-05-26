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
from typing import List, Dict, Optional, Tuple, Any
from concurrent.futures import ThreadPoolExecutor, as_completed
from threading import Lock
import encodings.idna
from tqdm import tqdm
from netmiko import NetmikoTimeoutException, NetmikoAuthenticationException
from netmiko.ssh_dispatcher import CLASS_MAPPER

# 环境配置
os.environ["NO_COLOR"] = "1"
write_lock = Lock()
DEFAULT_THREADS = min(900, max(4, (os.cpu_count() or 4)))

# **动态获取所有支持的设备类型**
SUPPORTED_DEVICE_TYPES = set(CLASS_MAPPER.keys())

# **设备类型别名映射（常用别名到标准类型）**
DEVICE_TYPE_ALIASES = {
    # Cisco设备
    'cisco': 'cisco_ios',
    'cisco_switch': 'cisco_ios', 
    'cisco_router': 'cisco_ios',
    'cisco_catalyst': 'cisco_ios',
    'nexus': 'cisco_nxos',
    'cisco_nexus': 'cisco_nxos',
    'asa': 'cisco_asa',
    'cisco_firewall': 'cisco_asa',
    'ios_xe': 'cisco_xe',
    'ios_xr': 'cisco_xr',
    'cisco_wlc': 'cisco_wlc_ssh',
    
    # Huawei设备
    'huawei': 'huawei',
    'huawei_switch': 'huawei',
    'huawei_router': 'huawei',
    'huawei_firewall': 'huawei',
    'vrp': 'huawei',
    'vrpv8': 'huawei_vrpv8',
    'huawei_vrp': 'huawei_vrpv8',
    
    # HP设备  
    'hp': 'hp_comware',
    'hp_switch': 'hp_comware',
    'comware': 'hp_comware',
    'procurve': 'hp_procurve',
    'hp_procurve_switch': 'hp_procurve',
    'aruba': 'aruba_os',
    'aruba_switch': 'aruba_os',
    
    # H3C设备
    'h3c': 'hp_comware',
    'h3c_switch': 'hp_comware',
    'h3c_router': 'hp_comware',
    'h3c_comware': 'hp_comware',
    
    # Juniper设备
    'juniper': 'juniper',
    'junos': 'juniper', 
    'juniper_switch': 'juniper',
    'juniper_router': 'juniper',
    'juniper_firewall': 'juniper_screenos',
    'srx': 'juniper_screenos',
    
    # Fortinet设备
    'fortinet': 'fortinet',
    'fortigate': 'fortinet',
    'fortios': 'fortinet',
    'fortinet_firewall': 'fortinet',
    
    # PaloAlto设备
    'paloalto': 'paloalto_panos',
    'panos': 'paloalto_panos',
    'pa': 'paloalto_panos',
    'paloalto_firewall': 'paloalto_panos',
    
    # Dell设备
    'dell': 'dell_force10',
    'dell_switch': 'dell_force10',
    'force10': 'dell_force10',
    'dell_powerconnect': 'dell_powerconnect',
    'dell_os6': 'dell_os6',
    'dell_os9': 'dell_os9',
    'dell_os10': 'dell_os10',
    
    # Extreme设备
    'extreme': 'extreme',
    'extreme_switch': 'extreme',
    'extreme_exos': 'extreme_exos',
    'exos': 'extreme_exos',
    'extreme_wing': 'extreme_wing',
    
    # Ruckus设备
    'ruckus': 'ruckus_fastiron',
    'ruckus_switch': 'ruckus_fastiron',
    'ruckus_icx': 'ruckus_fastiron',
    'fastiron': 'ruckus_fastiron',
    'brocade': 'ruckus_fastiron',
    
    # Mikrotik设备
    'mikrotik': 'mikrotik_routeros',
    'routeros': 'mikrotik_routeros',
    'mikrotik_router': 'mikrotik_routeros',
    
    # Alcatel设备
    'alcatel': 'alcatel_aos',
    'aos': 'alcatel_aos',
    'alcatel_switch': 'alcatel_aos',
    'nokia': 'nokia_sros',
    'sros': 'nokia_sros',
    
    # Avaya设备
    'avaya': 'avaya_ers',
    'ers': 'avaya_ers',
    'avaya_switch': 'avaya_ers',
    
    # Allied Telesis设备
    'allied_telesis': 'allied_telesis_awplus',
    'awplus': 'allied_telesis_awplus',
    'at': 'allied_telesis_awplus',
    
    # F5设备
    'f5': 'f5_tmsh',
    'bigip': 'f5_tmsh',
    'f5_ltm': 'f5_tmsh',
    
    # A10设备
    'a10': 'a10',
    'a10_acos': 'a10',
    
    # Linux设备
    'linux': 'linux',
    'ubuntu': 'linux',
    'centos': 'linux',
    'redhat': 'linux',
    'debian': 'linux',
    
    # Others
    'generic_termserver': 'generic_termserver',
    'terminal_server': 'generic_termserver'
}

# **设备厂商分类配置**
DEVICE_VENDOR_CONFIGS = {
    # Cisco厂商设备
    'cisco': {
        'timeout': 25,
        'banner_timeout': 15,
        'auth_timeout': 10,
        'fast_cli': False,
        'session_timeout': 60,
        'global_delay_factor': 1,
        'conn_timeout': 10
    },
    
    # Huawei厂商设备
    'huawei': {
        'timeout': 30,
        'banner_timeout': 20,
        'auth_timeout': 15,
        'fast_cli': False,
        'session_timeout': 90,
        'global_delay_factor': 2,
        'conn_timeout': 15
    },
    
    # Juniper厂商设备
    'juniper': {
        'timeout': 35,
        'banner_timeout': 25,
        'auth_timeout': 15,
        'fast_cli': False,
        'session_timeout': 120,
        'global_delay_factor': 2,
        'conn_timeout': 15
    },
    
    # HP/Aruba厂商设备
    'hp': {
        'timeout': 25,
        'banner_timeout': 15,
        'auth_timeout': 10,
        'fast_cli': False,
        'session_timeout': 60,
        'global_delay_factor': 1,
        'conn_timeout': 10
    },
    
    # Fortinet厂商设备
    'fortinet': {
        'timeout': 30,
        'banner_timeout': 20,
        'auth_timeout': 15,
        'fast_cli': False,
        'session_timeout': 60,
        'global_delay_factor': 2,
        'conn_timeout': 15,
        'use_keys': False,
        'allow_agent': False
    },
    
    # PaloAlto厂商设备
    'paloalto': {
        'timeout': 45,
        'banner_timeout': 30,
        'auth_timeout': 20,
        'fast_cli': False,
        'session_timeout': 120,
        'global_delay_factor': 3,
        'conn_timeout': 20,
        'use_keys': False,
        'allow_agent': False
    },
    
    # Dell厂商设备
    'dell': {
        'timeout': 30,
        'banner_timeout': 20,
        'auth_timeout': 15,
        'fast_cli': False,
        'session_timeout': 60,
        'global_delay_factor': 1,
        'conn_timeout': 10
    },
    
    # Extreme厂商设备
    'extreme': {
        'timeout': 30,
        'banner_timeout': 20,
        'auth_timeout': 15,
        'fast_cli': False,
        'session_timeout': 60,
        'global_delay_factor': 2,
        'conn_timeout': 15
    },
    
    # Ruckus/Brocade厂商设备
    'ruckus': {
        'timeout': 30,
        'banner_timeout': 20,
        'auth_timeout': 15,
        'fast_cli': False,
        'session_timeout': 60,
        'global_delay_factor': 2,
        'conn_timeout': 15
    },
    
    # 默认配置（未知厂商）
    'default': {
        'timeout': 30,
        'banner_timeout': 15,
        'auth_timeout': 10,
        'fast_cli': False,
        'session_timeout': 60,
        'global_delay_factor': 1,
        'conn_timeout': 10
    }
}

def thread_initializer() -> None:
    """线程初始化（解决编码问题）"""
    import encodings.idna
    encodings.idna.__name__

def sanitize_filename(name: str) -> str:
    """生成安全文件名"""
    return re.sub(r'[\\/*?:"<>|]', '', name).strip()[:60]

def normalize_device_type(device_type: str) -> str:
    """**智能设备类型标准化**"""
    original_type = device_type.lower().strip()
    
    # 直接匹配支持的设备类型
    if original_type in SUPPORTED_DEVICE_TYPES:
        return original_type
    
    # 通过别名映射
    if original_type in DEVICE_TYPE_ALIASES:
        mapped_type = DEVICE_TYPE_ALIASES[original_type]
        if mapped_type in SUPPORTED_DEVICE_TYPES:
            return mapped_type
    
    # 模糊匹配（部分字符串匹配）
    for supported_type in SUPPORTED_DEVICE_TYPES:
        if original_type in supported_type or supported_type in original_type:
            return supported_type
    
    # 如果都找不到，返回原始类型（会在后续给出警告）
    return original_type

def get_device_vendor(device_type: str) -> str:
    """**根据设备类型获取厂商**"""
    device_type = device_type.lower()
    
    vendor_mapping = {
        'cisco': ['cisco_', 'ios', 'nxos', 'asa', 'wlc', 'xe', 'xr'],
        'huawei': ['huawei', 'vrp'],
        'juniper': ['juniper', 'junos', 'screenos'],
        'hp': ['hp_', 'aruba', 'comware', 'procurve'],
        'fortinet': ['fortinet'],
        'paloalto': ['paloalto', 'panos'],
        'dell': ['dell_'],
        'extreme': ['extreme'],
        'ruckus': ['ruckus', 'brocade', 'fastiron'],
        'mikrotik': ['mikrotik'],
        'alcatel': ['alcatel', 'nokia'],
        'avaya': ['avaya'],
        'f5': ['f5_'],
        'a10': ['a10']
    }
    
    for vendor, patterns in vendor_mapping.items():
        if any(pattern in device_type for pattern in patterns):
            return vendor
    
    return 'default'

def get_device_config(device_type: str) -> Dict[str, Any]:
    """**获取设备特定配置**"""
    vendor = get_device_vendor(device_type)
    return DEVICE_VENDOR_CONFIGS.get(vendor, DEVICE_VENDOR_CONFIGS['default']).copy()

def validate_device_data(device: Dict[str, str], row_idx: int) -> None:
    """验证设备数据完整性"""
    required = ['host', 'username', 'password', 'device_type']
    if missing := [f for f in required if not device.get(f)]:
        raise ValueError(f"Row {row_idx} 缺失字段: {', '.join(missing)}")
    
    # **验证设备类型是否支持**
    normalized_type = normalize_device_type(device['device_type'])
    if normalized_type not in SUPPORTED_DEVICE_TYPES:
        print(f"[WARN] Row {row_idx}: 未知设备类型 '{device['device_type']}' -> '{normalized_type}', 将使用默认配置")

def load_excel(excel_file: str, sheet_name: str = 'Sheet1') -> List[Dict[str, str]]:
    """加载Excel设备清单"""
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
            if not any(row):  # 跳过空行
                continue
            device = {headers[i]: str(cell).strip() if cell else "" for i, cell in enumerate(row)}
            validate_device_data(device, row_idx)
            
            # **标准化设备类型**
            original_type = device['device_type']
            device['device_type'] = normalize_device_type(original_type)
            device['original_type'] = original_type  # 保留原始类型用于日志
            
            devices.append(device)
            
        return devices
    except Exception as e:
        print(f"Excel处理失败: {str(e)}")
        sys.exit(1)
    finally:
        if wb:
            wb.close()

def connect_device(device: Dict[str, str]) -> Optional[netmiko.BaseConnection]:
    """**通用设备连接（支持所有netmiko设备）**"""
    device_type = device['device_type']
    device_config = get_device_config(device_type)
    vendor = get_device_vendor(device_type)
    
    # **基础连接参数**
    params = {
        'device_type': device_type,
        'host': device['host'],
        'username': device['username'],
        'password': device['password'],
        'timeout': device_config['timeout'],
        'banner_timeout': device_config['banner_timeout'],
        'auth_timeout': device_config['auth_timeout'],
        'fast_cli': device_config['fast_cli'],
        'session_timeout': device_config['session_timeout'],
        'global_delay_factor': device_config['global_delay_factor'],
        'conn_timeout': device_config['conn_timeout'],
        'read_timeout_override': int(device.get('readtime', device_config['timeout']))
    }

    # **可选参数**
    if device.get('secret'):
        params['secret'] = device['secret']
    if device.get('port'):
        params['port'] = int(device['port'])
    
    # **厂商特定配置**
    if 'use_keys' in device_config:
        params['use_keys'] = device_config['use_keys']
    if 'allow_agent' in device_config:
        params['allow_agent'] = device_config['allow_agent']

    # **特殊协议配置**
    if device_type.endswith('_telnet'):
        # Telnet连接不需要SSH相关参数
        params.pop('use_keys', None)
        params.pop('allow_agent', None)
    elif device_type.endswith('_serial'):
        # 串口连接特殊配置
        if device.get('serial_settings'):
            params['serial_settings'] = device['serial_settings']

    # **调试日志配置**
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
            post_connection_setup(conn, device_type, vendor, device.get('secret'))
            
            return conn
            
        except (NetmikoTimeoutException, NetmikoAuthenticationException) as e:
            if attempt < max_retries:
                print(f"[RETRY {attempt+1}] {device['host']}: {e.__class__.__name__}")
                time.sleep(2 ** attempt)
                continue
            log_error(device['host'], f"{e.__class__.__name__}: {str(e)} (Type: {device.get('original_type', device_type)})")
        except Exception as e:
            if attempt < max_retries:
                print(f"[RETRY {attempt+1}] {device['host']}: Connection error")
                time.sleep(2 ** attempt)
                continue
            log_error(device['host'], f"连接异常: {str(e)} (Type: {device.get('original_type', device_type)})")
    
    return None

def post_connection_setup(conn: netmiko.BaseConnection, device_type: str, vendor: str, secret: Optional[str]) -> None:
    """**连接后设备特定设置**"""
    try:
        # Enable模式设备
        enable_required_vendors = ['cisco', 'hp', 'ruckus', 'extreme', 'dell']
        enable_required_types = ['cisco_', 'hp_', 'aruba_', 'ruckus_', 'extreme', 'dell_']
        
        if secret and (vendor in enable_required_vendors or any(t in device_type for t in enable_required_types)):
            try:
                conn.enable()
            except:
                # 某些设备可能不需要enable或已经在enable模式
                pass
        
        # **厂商特定初始化**
        if vendor == 'huawei':
            # 华为设备：可能需要关闭分页
            try:
                conn.send_command('screen-length 0 temporary', expect_string='>')
            except:
                pass
        elif vendor == 'paloalto':
            # PAN-OS：等待系统就绪
            time.sleep(2)
        elif vendor == 'fortinet':
            # Fortinet：配置终端长度
            try:
                conn.send_command('config system console\nset output standard\nend')
            except:
                pass
        elif vendor == 'juniper':
            # Juniper：设置终端长度
            try:
                conn.send_command('set cli screen-length 0')
            except:
                pass
        elif device_type in ['mikrotik_routeros']:
            # MikroTik：特殊处理
            time.sleep(1)
        elif device_type.startswith('linux'):
            # Linux设备：设置TERM
            try:
                conn.send_command('export TERM=vt100')
            except:
                pass
                
    except Exception as e:
        # 后连接设置失败不应该中断连接
        pass

def execute_commands(device: Dict[str, str], config_set: bool) -> Optional[str]:
    """**通用命令执行（适配所有设备类型）**"""
    try:
        cmds = [c.strip() for c in device.get('mult_command', '').split(';') if c.strip()]
        if not cmds:
            print(f"{device['host']} [WARN] 无有效命令")
            return None

        if not (conn := connect_device(device)):
            return None

        with conn:
            device_type = device['device_type']
            vendor = get_device_vendor(device_type)
            
            # **获取设备主机名**
            try:
                device['hostname'] = extract_hostname(conn, device_type, vendor)
            except:
                device['hostname'] = 'unknown'

            # **执行命令**
            all_output = []
            
            if config_set:
                # **配置模式执行**
                all_output.extend(execute_config_commands(conn, cmds, device_type, vendor))
            else:
                # **普通模式执行**
                all_output.extend(execute_show_commands(conn, cmds, device_type, vendor))

            return "\n\n".join(all_output)
            
    except Exception as e:
        log_error(device['host'], f"执行异常 ({device.get('original_type', device_type)}): {str(e)}")
        return None

def extract_hostname(conn: netmiko.BaseConnection, device_type: str, vendor: str) -> str:
    """**提取设备主机名（多厂商适配）**"""
    try:
        prompt = conn.find_prompt().strip()
        
        # **厂商特定的主机名提取规则**
        hostname_patterns = {
            'paloalto': [r'(\S+?)[@#>]', r'$(\S+?)$', r'(\S+)[@#>]'],
            'fortinet': [r'$(\S+?)$', r'(\S+?)[#>]', r'(\S+)-'],
            'juniper': [r'(\S+?)[@#>]', r'(\S+?)%'],
            'mikrotik': [r'$$(\S+?)$$', r'(\S+?)>'],
            'linux': [r'(\S+?)[@#$]', r'(\S+?):'],
            'f5': [r'$(\S+?)$', r'(\S+?)#'],
            'default': [r'(\S*?)([\w.-]+)[#>@$]', r'(\S+?)[#>]', r'(\w+)']
        }
        
        patterns = hostname_patterns.get(vendor, hostname_patterns['default'])
        
        for pattern in patterns:
            match = re.search(pattern, prompt)
            if match:
                hostname = match.group(1)
                if hostname and len(hostname) > 1:  # 避免单字符主机名
                    return hostname
                    
        return 'unknown'
    except:
        return 'unknown'

def execute_config_commands(conn: netmiko.BaseConnection, cmds: List[str], device_type: str, vendor: str) -> List[str]:
    """**配置模式命令执行**"""
    outputs = []
    
    try:
        if vendor in ['paloalto', 'fortinet']:
            # 这些设备需要单独发送配置命令
            for cmd in cmds:
                output = conn.send_command(cmd, expect_string=r'[#>$]', delay_factor=2)
                outputs.append(f"Config Command: {cmd}\n{output}")
        else:
            # 标准配置模式
            output = conn.send_config_set(cmds, cmd_verify=False)
            outputs.append(output)
    except Exception as e:
        outputs.append(f"Config execution error: {str(e)}")
    
    return outputs

def execute_show_commands(conn: netmiko.BaseConnection, cmds: List[str], device_type: str, vendor: str) -> List[str]:
    """**查看命令执行**"""
    outputs = []
    
    # **厂商特定的延迟因子**
    delay_factors = {
        'huawei': 2,
        'paloalto': 3,
        'fortinet': 2,
        'juniper': 2,
        'ruckus': 2,
        'extreme': 2,
        'mikrotik': 3,
        'default': 1
    }
    
    delay_factor = delay_factors.get(vendor, 1)
    
    for cmd in cmds:
        try:
            # **设备特定的命令发送方式**
            if vendor == 'paloalto':
                output = conn.send_command(cmd, expect_string=r'[#>]', delay_factor=delay_factor, max_loops=200)
            elif vendor == 'fortinet':
                output = conn.send_command(cmd, delay_factor=delay_factor, max_loops=150)
            elif vendor == 'juniper':
                output = conn.send_command(cmd, delay_factor=delay_factor, strip_prompt=False)
            elif vendor == 'mikrotik':
                output = conn.send_command(cmd, expect_string=r'[>\]]\s*$', delay_factor=delay_factor)
            elif device_type.startswith('linux'):
                output = conn.send_command(cmd, expect_string=r'[#$]\s*$', delay_factor=delay_factor)
            else:
                # 标准命令发送
                output = conn.send_command(cmd, delay_factor=delay_factor)
            
            outputs.append(f"Command: {cmd}\n{output}")
            
        except Exception as e:
            outputs.append(f"Command: {cmd}\nError: {str(e)}")
    
    return outputs

def save_result(ip: str, hostname: str, output: str, dest_path: str, device_type: str = '', original_type: str = '') -> None:
    """保存执行结果"""
    date_str = datetime.datetime.now().strftime('%Y%m%d')
    output_dir = os.path.join(dest_path, f"result_{date_str}")
    os.makedirs(output_dir, exist_ok=True)

    vendor = get_device_vendor(device_type)
    filename = f"{sanitize_filename(ip)}_{hostname}_{vendor}_{device_type}.txt"
    timestamp = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    
    content = f"""=== 设备信息 ===
IP地址: {ip}
主机名: {hostname}
设备类型: {device_type}
原始类型: {original_type}
厂商: {vendor}
执行时间: {timestamp}

=== 执行结果 ===
{output}"""

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
    """批量执行"""
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
                                dev['device_type'],
                                dev.get('original_type', dev['device_type'])
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

def list_supported_devices() -> None:
    """**显示所有支持的设备类型**"""
    print(f"**支持的设备类型总数**: {len(SUPPORTED_DEVICE_TYPES)}\n")
    
    # 按厂商分组显示
    vendor_devices = {}
    for device_type in sorted(SUPPORTED_DEVICE_TYPES):
        vendor = get_device_vendor(device_type)
        if vendor not in vendor_devices:
            vendor_devices[vendor] = []
        vendor_devices[vendor].append(device_type)
    
    for vendor, devices in sorted(vendor_devices.items()):
        print(f"**{vendor.upper()}** ({len(devices)} 种):")
        for device in sorted(devices):
            print(f"  - {device}")
        print()

def parse_args() -> argparse.Namespace:
    """命令行参数解析"""
    parser = argparse.ArgumentParser(
        description="**网络设备批量管理工具 v5.0** - 支持所有netmiko设备", 
        add_help=False, 
        formatter_class=argparse.RawTextHelpFormatter
    )
    parser.add_argument('-i', '--input', help='设备清单Excel路径')
    parser.add_argument('-t', '--threads', type=int, default=DEFAULT_THREADS, help=f'并发线程数 (默认: {DEFAULT_THREADS})')
    parser.add_argument('-cs', '--config_set', action='store_true', help='使用配置模式发送命令')
    parser.add_argument('-d', '--destination', default='./', help='结果保存路径')
    parser.add_argument('--debug', action='store_true', help='启用调试日志')
    parser.add_argument('-s', '--sheet', default='Sheet1', help='Excel工作表名称')
    parser.add_argument('--list-devices', action='store_true', help='列出所有支持的设备类型')
    
    if '--help' in sys.argv or '-h' in sys.argv:
        print(f"""
**网络设备批量管理工具 v5.0 - 全设备支持版本**

**特性**:
- **支持 {len(SUPPORTED_DEVICE_TYPES)} 种设备类型** (所有netmiko支持的设备)
- **智能设备类型识别** (支持别名和模糊匹配)
- **厂商特定优化配置** (针对不同厂商调优)
- **自动重试机制** (连接失败自动重试)
- **并发执行** (多线程提高效率)

**使用方法**:
  python connexec.py -i <设备清单.xlsx> [选项]

**参数说明**:
  -i, --input        必需  Excel文件路径  
  -t, --threads      可选  并发线程数 (默认: {DEFAULT_THREADS})
  -cs, --config_set  可选  使用配置模式
  -d, --destination  可选  结果保存路径
  -s, --sheet        可选  Excel工作表名
  --debug            可选  启用详细日志
  --list-devices     可选  显示支持的设备类型

**Excel格式**:
| host        | username | password | device_type  | secret | port | mult_command         |
|-------------|----------|----------|--------------|--------|------|---------------------|
| 192.168.1.1 | admin    | pass123  | cisco_ios    | enable | 22   | show version        |
| 192.168.1.2 | admin    | pass123  | huawei       |        | 22   | display version     |
| 192.168.1.3 | admin    | pass123  | paloalto     |        | 22   | show system info    |

**示例**:
  # 基本使用
  python connexec.py -i devices.xlsx
  
  # 配置模式 + 调试
  python connexec.py -i devices.xlsx -cs --debug
  
  # 查看支持的设备类型
  python connexec.py --list-devices
""")
        sys.exit(0)

    return parser.parse_args()

def main() -> None:
    """主入口"""
    args = parse_args()
    
    # 显示支持的设备类型
    if args.list_devices:
        list_supported_devices()
        return
    
    # 检查必需参数
    if not args.input:
        print("**错误**: 必须指定输入文件 (-i)")
        sys.exit(1)
        
    if not os.path.exists(args.input):
        print(f"**错误**: 文件不存在 [{args.input}]")
        sys.exit(1)
        
    try:
        devices = load_excel(args.input, args.sheet)
        
        # 注入debug标志
        if args.debug:
            for device in devices:
                device['debug'] = True
        
        # **显示设备统计信息**
        device_stats = {}
        vendor_stats = {}
        
        for device in devices:
            device_type = device['device_type']
            original_type = device.get('original_type', device_type)
            vendor = get_device_vendor(device_type)
            
            # 设备类型统计
            key = f"{original_type} -> {device_type}" if original_type != device_type else device_type
            device_stats[key] = device_stats.get(key, 0) + 1
            
            # 厂商统计
            vendor_stats[vendor] = vendor_stats.get(vendor, 0) + 1
        
        print(f"**成功加载设备**: {len(devices)} 台 (工作表: {args.sheet})")
        print(f"**网络设备管理工具** - 支持 {len(SUPPORTED_DEVICE_TYPES)} 种设备类型")
        
        print("\n**厂商分布**:")
        for vendor, count in sorted(vendor_stats.items()):
            print(f"  - **{vendor.upper()}**: {count} 台")
        
        print("\n**设备类型详情**:")
        for device_type, count in sorted(device_stats.items()):
            print(f"  - {device_type}: {count} 台")
        
        print()
        batch_execute(devices, args.config_set, args.threads, args.destination)
        
    except KeyboardInterrupt:
        print("\n**用户终止**")
        sys.exit(0)

if __name__ == "__main__":
    main()
