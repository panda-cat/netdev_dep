#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import os
import re
import sys
import time
import argparse
import datetime
from typing import List, Dict, Optional, Any
from concurrent.futures import ThreadPoolExecutor, as_completed
from threading import Lock

import netmiko
import openpyxl
from tqdm import tqdm
from netmiko import (
    ConnectHandler,
    NetmikoTimeoutException,
    NetmikoAuthenticationException,
)
from netmiko.ssh_dispatcher import CLASS_MAPPER

# ---------------------------------------------------------------------------
# 环境配置
# ---------------------------------------------------------------------------
os.environ["NO_COLOR"] = "1"
write_lock = Lock()
DEFAULT_THREADS = min(900, max(4, (os.cpu_count() or 4)))

SUPPORTED_DEVICE_TYPES = set(CLASS_MAPPER.keys())

# ---------------------------------------------------------------------------
# 设备类型别名映射（节选，可按需补全）
# ---------------------------------------------------------------------------
DEVICE_TYPE_ALIASES = {
    'cisco': 'cisco_ios', 'cisco_switch': 'cisco_ios', 'nexus': 'cisco_nxos',
    'asa': 'cisco_asa', 'ios_xe': 'cisco_xe', 'ios_xr': 'cisco_xr',
    'huawei': 'huawei', 'vrp': 'huawei', 'vrpv8': 'huawei_vrpv8',
    'hp': 'hp_comware', 'comware': 'hp_comware', 'h3c': 'hp_comware',
    'aruba': 'aruba_os', 'juniper': 'juniper', 'junos': 'juniper',
    'srx': 'juniper_screenos', 'fortinet': 'fortinet', 'fortigate': 'fortinet',
    'paloalto': 'paloalto_panos', 'panos': 'paloalto_panos',
    'dell': 'dell_force10', 'mikrotik': 'mikrotik_routeros',
    'routeros': 'mikrotik_routeros', 'nokia': 'nokia_sros', 'sros': 'nokia_sros',
    'f5': 'f5_tmsh', 'linux': 'linux',
    'generic_termserver': 'generic_termserver', 'terminal_server': 'generic_termserver',
}


def resolve_device_type(raw: str) -> str:
    """把别名或原始字符串解析为 Netmiko 标准 device_type。"""
    if not raw:
        return 'generic_termserver'
    key = raw.strip().lower()
    if key in SUPPORTED_DEVICE_TYPES:
        return key
    if key in DEVICE_TYPE_ALIASES:
        return DEVICE_TYPE_ALIASES[key]
    # 未知类型回退，避免直接崩溃
    return 'generic_termserver'


# ---------------------------------------------------------------------------
# 厂商连接参数（按解析后的 device_type 前缀匹配）
# ---------------------------------------------------------------------------
VENDOR_TIMEOUTS = {
    'cisco': {'timeout': 25, 'banner_timeout': 15, 'auth_timeout': 10},
    'huawei': {'timeout': 30, 'banner_timeout': 15, 'auth_timeout': 12},
    'hp': {'timeout': 30, 'banner_timeout': 15, 'auth_timeout': 12},
    'juniper': {'timeout': 30, 'banner_timeout': 15, 'auth_timeout': 12},
    'default': {'timeout': 25, 'banner_timeout': 15, 'auth_timeout': 10},
}


def get_conn_extra(device_type: str) -> Dict[str, Any]:
    for prefix, cfg in VENDOR_TIMEOUTS.items():
        if prefix != 'default' and device_type.startswith(prefix):
            return dict(cfg)
    return dict(VENDOR_TIMEOUTS['default'])


# ---------------------------------------------------------------------------
# 命令解析：把 Excel 单元格里的多条命令拆成列表
# ---------------------------------------------------------------------------
def parse_commands(cell_value: Any) -> List[str]:
    """支持分号、换行混合分隔；去空白、去空行。"""
    if cell_value is None:
        return []
    text = str(cell_value)
    # 同时按换行和分号拆分
    parts = re.split(r'[;\n\r]+', text)
    return [p.strip() for p in parts if p.strip()]


# ---------------------------------------------------------------------------
# 分页处理函数
# ---------------------------------------------------------------------------
def send_command_with_pagination(
    conn: netmiko.BaseConnection,
    cmd: str,
    delay_factor: float = 1.0,
    max_loops: int = 2000,
    read_chunk_sleep: float = 0.08,
    pager_space_sleep: float = 0.18,
    stall_loops_before_return: int = 10,
) -> str:
    """发送单条命令并处理分页提示符。"""
    pager_regexes = [
        r'--?[Mm]ore--?',
        r'[Mm]ore\s*[:?]',
        r'\(space.*?to\s+continue\)',
        r'(?i)press\s+any\s+key',
        r'(?i)press\s+<?space>?',
        r'(?i)hit\s+any\s+key',
    ]
    pager_re = re.compile('|'.join(pager_regexes))
    output: List[str] = []
    seen_pager = False
    empty_reads = 0
    remote = getattr(conn, "remote_conn", None)

    def drain_available() -> str:
        buf = []
        while True:
            chunk = conn.read_channel()
            if chunk:
                buf.append(chunk)
                time.sleep(read_chunk_sleep * delay_factor)
                continue
            if remote is not None and remote.recv_ready():
                time.sleep(read_chunk_sleep * delay_factor)
                continue
            break
        return ''.join(buf)

    conn.write_channel(cmd + conn.RETURN)
    time.sleep(0.2 * delay_factor)

    first = drain_available()
    if first:
        output.append(first)

    for _ in range(max_loops):
        tail = ''.join(output)[-512:] if output else ''
        if pager_re.search(tail):
            seen_pager = True
            conn.write_channel(' ')
            time.sleep(pager_space_sleep * delay_factor)
            chunk = drain_available()
            if chunk:
                output.append(chunk)
                empty_reads = 0
                continue
            empty_reads += 1
            if empty_reads >= stall_loops_before_return:
                conn.write_channel(conn.RETURN)
                time.sleep(pager_space_sleep * delay_factor)
                chunk = drain_available()
                if chunk:
                    output.append(chunk)
                    empty_reads = 0
                    continue
        else:
            chunk = drain_available()
            if chunk:
                output.append(chunk)
                empty_reads = 0
                time.sleep(read_chunk_sleep * delay_factor)
                continue
            empty_reads += 1
            if seen_pager and empty_reads >= 4:
                break
            if not seen_pager and empty_reads >= 6:
                break

    # 清理 ANSI 与分页残留行
    text = ''.join(output)
    text = re.sub(r'\x1b\[[0-9;?]*[A-Za-z]', '', text)
    text = pager_re.sub('', text)
    return text


# ---------------------------------------------------------------------------
# 单台设备处理
# ---------------------------------------------------------------------------
def process_device(
    device: Dict[str, Any],
    global_commands: List[str],
    use_pagination: bool,
) -> Dict[str, Any]:
    """连接单台设备并执行其专属命令列表。"""
    host = device['host']
    device_type = device['device_type']

    # 行内命令优先；为空才用全局兜底
    commands = device.get('commands') or global_commands

    result: Dict[str, Any] = {
        'host': host,
        'device_type': device_type,
        'status': 'failed',
        'error': '',
        'outputs': [],  # [(cmd, output), ...]
    }

    if not commands:
        result['error'] = '无命令可执行（行内与全局均为空）'
        return result

    conn_params = {
        'device_type': device_type,
        'host': host,
        'username': device['username'],
        'password': device['password'],
        'port': device.get('port', 22),
        **get_conn_extra(device_type),
    }
    if device.get('secret'):
        conn_params['secret'] = device['secret']

    try:
        with ConnectHandler(**conn_params) as conn:
            if device.get('secret'):
                try:
                    conn.enable()
                except Exception:
                    pass  # 部分设备无需 enable

            for cmd in commands:
                try:
                    if use_pagination:
                        out = send_command_with_pagination(conn, cmd)
                    else:
                        out = conn.send_command(
                            cmd, read_timeout=60, strip_prompt=False,
                            strip_command=False,
                        )
                except Exception as e:
                    out = f'[命令执行异常] {e}'
                result['outputs'].append((cmd, out))

        result['status'] = 'success'

    except NetmikoAuthenticationException:
        result['error'] = '认证失败（用户名/密码错误）'
    except NetmikoTimeoutException:
        result['error'] = '连接超时（设备不可达或端口未开）'
    except Exception as e:
        result['error'] = f'未知错误: {e}'

    return result


# ---------------------------------------------------------------------------
# 读取 Excel 设备清单
# ---------------------------------------------------------------------------
def load_devices(path: str) -> List[Dict[str, Any]]:
    """
    读取设备清单。表头需包含（大小写不敏感）：
    host, device_type, username, password, [secret], [port], [commands]
    """
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb.active

    rows = ws.iter_rows(values_only=True)
    header = next(rows, None)
    if not header:
        raise ValueError('Excel 为空或缺少表头')

    col = {str(h).strip().lower(): i for i, h in enumerate(header) if h is not None}

    required = ['host', 'device_type', 'username', 'password']
    missing = [r for r in required if r not in col]
    if missing:
        raise ValueError(f'缺少必需列: {missing}')

    devices: List[Dict[str, Any]] = []
    for row in rows:
        if row is None or all(c is None for c in row):
            continue
        host = row[col['host']]
        if not host:
            continue

        device = {
            'host': str(host).strip(),
            'device_type': resolve_device_type(str(row[col['device_type']] or '')),
            'username': str(row[col['username']] or '').strip(),
            'password': str(row[col['password']] or ''),
            'secret': str(row[col['secret']]).strip() if 'secret' in col and row[col['secret']] else '',
            'port': int(row[col['port']]) if 'port' in col and row[col['port']] else 22,
            'commands': parse_commands(row[col['commands']]) if 'commands' in col else [],
        }
        devices.append(device)

    wb.close()
    return devices


# ---------------------------------------------------------------------------
# 写出结果
# ---------------------------------------------------------------------------
def write_results(results: List[Dict[str, Any]], out_path: str) -> None:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'results'
    ws.append(['host', 'device_type', 'status', 'command', 'output', 'error'])

    for r in results:
        if r['outputs']:
            for cmd, out in r['outputs']:
                ws.append([r['host'], r['device_type'], r['status'], cmd, out, r['error']])
        else:
            ws.append([r['host'], r['device_type'], r['status'], '', '', r['error']])

    wb.save(out_path)


# ---------------------------------------------------------------------------
# 主流程
# ---------------------------------------------------------------------------
def main() -> None:
    parser = argparse.ArgumentParser(
        description='批量采集网络设备命令输出（每台设备命令在 Excel 中定义）'
    )
    parser.add_argument('-i', '--input', required=True, help='设备清单 Excel 路径')
    parser.add_argument('-o', '--output', help='结果输出 Excel 路径')
    parser.add_argument(
        '-c', '--command', action='append', default=[],
        help='全局兜底命令（仅当某设备行未定义 commands 时使用，可重复指定）',
    )
    parser.add_argument(
        '-t', '--threads', type=int, default=DEFAULT_THREADS,
        help=f'并发线程数（默认 {DEFAULT_THREADS}）',
    )
    parser.add_argument(
        '--no-pagination', action='store_true',
        help='禁用自定义分页处理，改用 netmiko 默认 send_command',
    )
    args = parser.parse_args()

    if not os.path.isfile(args.input):
        print(f'输入文件不存在: {args.input}', file=sys.stderr)
        sys.exit(1)

    out_path = args.output or (
        f'result_{datetime.datetime.now():%Y%m%d_%H%M%S}.xlsx'
    )

    try:
        devices = load_devices(args.input)
    except Exception as e:
        print(f'读取设备清单失败: {e}', file=sys.stderr)
        sys.exit(1)

    if not devices:
        print('没有可处理的设备', file=sys.stderr)
        sys.exit(1)

    global_commands = [c.strip() for c in args.command if c.strip()]
    use_pagination = not args.no_pagination

    print(f'共 {len(devices)} 台设备，并发 {args.threads}，'
          f'分页处理: {"开启" if use_pagination else "关闭"}')

    results: List[Dict[str, Any]] = []
    with ThreadPoolExecutor(max_workers=args.threads) as pool:
        futures = {
            pool.submit(process_device, dev, global_commands, use_pagination): dev
            for dev in devices
        }
        for fut in tqdm(as_completed(futures), total=len(futures), desc='进度'):
            try:
                results.append(fut.result())
            except Exception as e:
                dev = futures[fut]
                results.append({
                    'host': dev['host'], 'device_type': dev['device_type'],
                    'status': 'failed', 'error': f'线程异常: {e}', 'outputs': [],
                })

    # 保持与输入顺序一致
    order = {d['host']: i for i, d in enumerate(devices)}
    results.sort(key=lambda r: order.get(r['host'], 1 << 30))

    with write_lock:
        write_results(results, out_path)

    ok = sum(1 for r in results if r['status'] == 'success')
    print(f'\n完成：成功 {ok} / 共 {len(results)}，结果已写入 {out_path}')


if __name__ == '__main__':
    main()
