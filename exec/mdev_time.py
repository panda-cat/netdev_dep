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

# 需要手写分页的 device_type（Netmiko 不会自动应答 --More--）
MANUAL_PAGER_TYPES = {'generic_termserver', 'terminal_server', 'generic'}

# ---------------------------------------------------------------------------
# 全局正则（大小写不敏感的 flag 统一放 re.compile 参数里）
# ---------------------------------------------------------------------------
PAGER_RE = re.compile(
    '|'.join([
        r'--?\s*more\s*--?',           # --More--, - More -
        r'more\s*[:?]',                # More:, More?
        r'\(space.*?to\s+continue\)',  # (space to continue)
        r'press\s+any\s+key',
        r'press\s+<?space>?',
        r'hit\s+any\s+key',
    ]),
    re.IGNORECASE,
)
ANSI_RE = re.compile(r'\x1b\[[0-9;?]*[A-Za-z]')

# ---------------------------------------------------------------------------
# 设备类型别名映射
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
    'generic_termserver': 'generic_termserver',
    'terminal_server': 'generic_termserver',
}


def resolve_device_type(raw: Any) -> str:
    """把别名或原始字符串解析为 Netmiko 标准 device_type。"""
    if raw is None:
        return 'generic_termserver'
    key = str(raw).strip().lower()
    if not key:
        return 'generic_termserver'
    if key in SUPPORTED_DEVICE_TYPES:
        return key
    if key in DEVICE_TYPE_ALIASES:
        return DEVICE_TYPE_ALIASES[key]
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
    parts = re.split(r'[;\n\r]+', text)
    return [p.strip() for p in parts if p.strip()]


# ---------------------------------------------------------------------------
# Excel 读取：每行一台设备，携带自己的命令列表
# ---------------------------------------------------------------------------
def load_excel(path: str, default_cmds: List[str]) -> List[Dict[str, Any]]:
    """读取 Excel，返回设备字典列表。

    必需列：host, device_type, username, password
    可选列：secret, port, commands
    commands 为空时回退到 default_cmds（命令行 -c）。
    """
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.active

    header = [str(c.value).strip().lower() if c.value is not None else ''
              for c in ws[1]]
    col = {name: idx for idx, name in enumerate(header)}

    required = ['host', 'device_type', 'username', 'password']
    missing = [r for r in required if r not in col]
    if missing:
        raise ValueError(f"Excel 缺少必需列: {', '.join(missing)}")

    def cell(row, name) -> Any:
        idx = col.get(name)
        if idx is None or idx >= len(row):
            return None
        return row[idx]

    def as_str(v: Any) -> str:
        """数值单元格统一转字符串，避免 int.strip() 报错。"""
        if v is None:
            return ''
        return str(v).strip()

    devices: List[Dict[str, Any]] = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        host = as_str(cell(row, 'host'))
        if not host:
            continue  # 跳过空行

        row_cmds = parse_commands(cell(row, 'commands'))
        if not row_cmds:
            row_cmds = list(default_cmds)  # 回退到全局 -c
        if not row_cmds:
            # 既无行内命令也无全局命令，跳过
            continue

        port_raw = as_str(cell(row, 'port'))
        try:
            port = int(port_raw) if port_raw else 22
        except ValueError:
            port = 22

        devices.append({
            'host': host,
            'device_type': resolve_device_type(cell(row, 'device_type')),
            'username': as_str(cell(row, 'username')),
            'password': as_str(cell(row, 'password')),
            'secret': as_str(cell(row, 'secret')),
            'port': port,
            'commands': row_cmds,
        })

    return devices


# ---------------------------------------------------------------------------
# 主机名提取（修复 LaTeX 残留的非法正则）
# ---------------------------------------------------------------------------
HOSTNAME_PATTERNS = [
    re.compile(r'[\r\n]([\w.\-]+)[#>]\s*$'),      # cisco/huawei/hp: name# 或 name>
    re.compile(r'[\r\n]<([\w.\-]+)>\s*$'),         # huawei: <name>
    re.compile(r'[\r\n]\[([\w.\-]+)\]\s*$'),       # huawei 系统视图: [name]
    re.compile(r'[\r\n](\S+?)@(\S+?)[>#%]\s*$'),   # juniper/linux: user@host>
]


def extract_hostname(prompt: str, fallback: str) -> str:
    """从设备提示符里提取主机名，失败则用 fallback（通常是 IP）。"""
    if not prompt:
        return fallback
    text = ANSI_RE.sub('', prompt)
    for pat in HOSTNAME_PATTERNS:
        m = pat.search(text)
        if m:
            # juniper/linux 模式取第 2 组（host），其余取第 1 组
            return m.group(2) if pat.groups >= 2 and m.lastindex and m.lastindex >= 2 else m.group(1)
    return fallback


# ---------------------------------------------------------------------------
# 终端服务器手写分页
# ---------------------------------------------------------------------------
def send_command_termserver(
    conn: netmiko.BaseConnection,
    cmd: str,
    delay_factor: float = 1.0,
    max_loops: int = 2000,
) -> str:
    """对终端服务器/generic 设备手动处理分页。"""
    remote = getattr(conn, "remote_conn", None)

    def drain() -> str:
        buf = []
        while True:
            chunk = conn.read_channel()
            if chunk:
                buf.append(chunk)
                time.sleep(0.08 * delay_factor)
                continue
            if remote is not None and remote.recv_ready():
                time.sleep(0.08 * delay_factor)
                continue
            break
        return ''.join(buf)

    output: List[str] = []
    conn.write_channel(cmd + conn.RETURN)
    time.sleep(0.3 * delay_factor)

    first = drain()
    if first:
        output.append(first)

    seen_pager = False
    empty_reads = 0

    for _ in range(max_loops):
        tail = ANSI_RE.sub('', ''.join(output))[-512:]

        if PAGER_RE.search(tail):
            seen_pager = True
            conn.write_channel(' ')
            time.sleep(0.18 * delay_factor)
            chunk = drain()
            if chunk:
                output.append(chunk)
                empty_reads = 0
                continue
            empty_reads += 1
        else:
            chunk = drain()
            if chunk:
                output.append(chunk)
                empty_reads = 0
                time.sleep(0.08 * delay_factor)
                continue
            empty_reads += 1
            if seen_pager and empty_reads >= 4:
                break
            if not seen_pager and empty_reads >= 6:
                break

    raw = ''.join(output)
    raw = ANSI_RE.sub('', raw)
    raw = PAGER_RE.sub('', raw)
    return raw


# ---------------------------------------------------------------------------
# 单台设备执行
# ---------------------------------------------------------------------------
def execute_device(device: Dict[str, Any], output_dir: str,
                   debug: bool = False) -> Dict[str, Any]:
    """连接单台设备，按其专属命令列表逐条执行，写入 txt。"""
    host = device['host']
    device_type = device['device_type']
    manual_pager = device_type in MANUAL_PAGER_TYPES

    conn_params = {
        'device_type': device_type,
        'host': host,
        'username': device['username'],
        'password': device['password'],
        'port': device['port'],
        **get_conn_extra(device_type),
    }
    if device.get('secret'):
        conn_params['secret'] = device['secret']

    result = {'host': host, 'ok': False, 'error': None, 'file': None}
    blocks: List[str] = []

    try:
        with ConnectHandler(**conn_params) as conn:
            if device.get('secret'):
                try:
                    conn.enable()
                except Exception:
                    pass  # 部分设备无需 enable

            prompt = ''
            try:
                prompt = conn.find_prompt()
            except Exception:
                pass
            hostname = extract_hostname(prompt, host)

            for cmd in device['commands']:
                if manual_pager:
                    out = send_command_termserver(conn, cmd)
                else:
                    out = conn.send_command(
                        cmd,
                        strip_prompt=False,
                        strip_command=False,
                        read_timeout=60,
                    )
                blocks.append(
                    f"{'=' * 60}\n[{hostname}] # {cmd}\n{'=' * 60}\n{out}\n"
                )

            result['ok'] = True
            result['hostname'] = hostname

    except NetmikoAuthenticationException:
        result['error'] = '认证失败（用户名/密码错误）'
    except NetmikoTimeoutException:
        result['error'] = '连接超时（无法到达或 SSH 未开启）'
    except Exception as e:
        result['error'] = f'{type(e).__name__}: {e}'

    # 写文件：成功写回显，失败写错误信息（便于排查）
    ts = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
    safe_host = re.sub(r'[^\w.\-]', '_', host)
    fname = f"{safe_host}_{ts}.txt"
    fpath = os.path.join(output_dir, fname)

    with write_lock:
        with open(fpath, 'w', encoding='utf-8') as f:
            if result['ok']:
                f.write('\n'.join(blocks))
            else:
                f.write(f"[{host}] 执行失败: {result['error']}\n")
                if debug and blocks:
                    f.write('\n--- 已捕获的部分输出 ---\n')
                    f.write('\n'.join(blocks))

    result['file'] = fpath
    return result


# ---------------------------------------------------------------------------
# 并发调度
# ---------------------------------------------------------------------------
def run_all(devices: List[Dict[str, Any]], output_dir: str,
            threads: int, debug: bool) -> List[Dict[str, Any]]:
    os.makedirs(output_dir, exist_ok=True)
    results: List[Dict[str, Any]] = []

    with ThreadPoolExecutor(max_workers=threads) as pool:
        futures = {
            pool.submit(execute_device, dev, output_dir, debug): dev
            for dev in devices
        }
        for fut in tqdm(as_completed(futures), total=len(futures),
                        desc='执行进度', ncols=80):
            try:
                results.append(fut.result())
            except Exception as e:
                dev = futures[fut]
                results.append({
                    'host': dev['host'], 'ok': False,
                    'error': f'线程异常: {e}', 'file': None,
                })

    return results


# ---------------------------------------------------------------------------
# 命令行入口
# ---------------------------------------------------------------------------
def main():
    parser = argparse.ArgumentParser(
        description='批量在网络设备上执行命令（每台设备命令由 Excel 行内定义）'
    )
    parser.add_argument('-i', '--input', required=True, help='设备清单 Excel 路径')
    parser.add_argument('-o', '--output', default='output',
                        help='输出目录（默认 output）')
    parser.add_argument('-c', '--commands', default='',
                        help='全局兜底命令（分号分隔），仅在某行 commands 为空时使用')
    parser.add_argument('-t', '--threads', type=int, default=DEFAULT_THREADS,
                        help=f'并发线程数（默认 {DEFAULT_THREADS}）')
    parser.add_argument('--debug', action='store_true',
                        help='失败时也写入已捕获的部分输出')
    args = parser.parse_args()

    if not os.path.isfile(args.input):
        print(f'错误：找不到输入文件 {args.input}', file=sys.stderr)
        sys.exit(1)

    default_cmds = parse_commands(args.commands)

    try:
        devices = load_excel(args.input, default_cmds)
    except Exception as e:
        print(f'读取 Excel 失败: {e}', file=sys.stderr)
        sys.exit(1)

    if not devices:
        print('没有可执行的设备（检查 host 列与 commands 列是否为空）', file=sys.stderr)
        sys.exit(1)

    print(f'共 {len(devices)} 台设备，并发 {args.threads} 线程开始执行...')
    results = run_all(devices, args.output, args.threads, args.debug)

    ok = sum(1 for r in results if r['ok'])
    fail = len(results) - ok
    print(f'\n完成：成功 {ok} 台，失败 {fail} 台。结果已保存到 {args.output}/')

    if fail:
        print('\n失败设备：')
        for r in results:
            if not r['ok']:
                print(f"  {r['host']}: {r['error']}")


if __name__ == '__main__':
    main()
