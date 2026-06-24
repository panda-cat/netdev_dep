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
import time
from typing import List, Dict, Optional
from concurrent.futures import ThreadPoolExecutor, as_completed
from threading import Lock
import encodings.idna
from tqdm import tqdm
from netmiko import NetmikoTimeoutException, NetmikoAuthenticationException

# ---------------------------------------------------------------------------
# 环境配置
# ---------------------------------------------------------------------------
os.environ["NO_COLOR"] = "1"
write_lock = Lock()
DEFAULT_THREADS = min(900, max(4, (os.cpu_count() or 4)))

# 需要手写分页的 device_type（Netmiko send_command 不会自动应答 --More--）
MANUAL_PAGER_TYPES = {'generic_termserver', 'terminal_server', 'generic'}

# pager 匹配：大小写不敏感，flag 放在 re.compile 参数里（避免 (?i) 中置报错）
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
# 工具函数
# ---------------------------------------------------------------------------
def thread_initializer() -> None:
    """线程初始化（解决编码问题）"""
    import encodings.idna
    encodings.idna.__name__  # 防止被优化


def sanitize_filename(name: str) -> str:
    """生成安全文件名"""
    return re.sub(r'[\\/*?:"<>|]', '', name).strip()[:60]


def log_error(host: str, msg: str) -> None:
    """统一错误日志（控制台 + error_log 文件）"""
    line = f"[{datetime.datetime.now():%Y-%m-%d %H:%M:%S}] {host} {msg}"
    print(line)
    with write_lock:
        with open("error_log.txt", "a", encoding="utf-8") as f:
            f.write(line + "\n")


def validate_device_data(device: Dict[str, str], row_idx: int) -> None:
    """验证设备数据完整性"""
    required = ['host', 'username', 'password', 'device_type']
    if missing := [f for f in required if not device.get(f)]:
        raise ValueError(f"Row {row_idx} 缺失字段: {', '.join(missing)}")


# ---------------------------------------------------------------------------
# Excel 加载
# ---------------------------------------------------------------------------
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
            device = {
                headers[i]: str(cell).strip() if cell else ""
                for i, cell in enumerate(row)
            }
            validate_device_data(device, row_idx)
            devices.append(device)

        return devices
    except Exception as e:
        print(f"Excel处理失败: {str(e)}")
        sys.exit(1)
    finally:
        if wb:
            wb.close()


# ---------------------------------------------------------------------------
# 设备连接
# ---------------------------------------------------------------------------
def connect_device(device: Dict[str, str]) -> Optional[netmiko.BaseConnection]:
    """设备连接（自动生成独立日志文件）"""
    params = {
        'device_type': device['device_type'],
        'host': device['host'],
        'username': device['username'],
        'password': device['password'],
        'secret': device.get('secret', ''),
        'read_timeout_override': int(device.get('readtime', 20)),
        'fast_cli': False,
    }
    if device.get('port'):
        params['port'] = int(device['port'])

    # 动态生成设备专属 debug 日志路径
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


# ---------------------------------------------------------------------------
# 终端服务器/generic 手写分页
# ---------------------------------------------------------------------------
def send_command_termserver(
    conn: netmiko.BaseConnection,
    cmd: str,
    delay_factor: float = 1.0,
    max_loops: int = 2000,
) -> str:
    """对终端服务器/generic 设备手动处理分页（send_command 不会自动应答 --More--）。"""
    remote = getattr(conn, "remote_conn", None)

    def drain() -> str:
        """持续读取当前可用字节，直到通道暂时无数据。"""
        buf = []
        while True:
            chunk = conn.read_channel()
            if chunk:
                buf.append(chunk)
                time.sleep(0.08 * delay_factor)
                continue
            # read_channel 非阻塞，单次可能只拿到部分，确认通道是否还有字节
            if remote is not None and remote.recv_ready():
                time.sleep(0.08 * delay_factor)
                continue
            break
        return ''.join(buf)

    output = []
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
            if empty_reads >= 10:  # 空格连续无效，尝试回车
                conn.write_channel(conn.RETURN)
                time.sleep(0.18 * delay_factor)
                chunk = drain()
                if chunk:
                    output.append(chunk)
                    empty_reads = 0
                    continue
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

    # 清理 ANSI 与 pager token 残留
    raw = ''.join(output)
    raw = ANSI_RE.sub('', raw)
    raw = PAGER_RE.sub('', raw)
    return raw


# ---------------------------------------------------------------------------
# 命令执行
# ---------------------------------------------------------------------------
def execute_commands(device: Dict[str, str], config_set: bool) -> Optional[str]:
    """执行命令并捕获输出（已识别厂商走 send_command，终端服务器走手写分页）。"""
    try:
        cmds = [c.strip() for c in device.get('mult_command', '').split(';') if c.strip()]
        if not cmds:
            print(f"{device['host']} [WARN] 无有效命令")
            return None

        if not (conn := connect_device(device)):
            return None

        dtype = device['device_type'].strip().lower()
        use_manual_pager = dtype in MANUAL_PAGER_TYPES

        with conn:
            # 获取设备主机名
            prompt = conn.find_prompt().strip()
            m = re.search(r'\S*?([\w.-]+)\s*[#>$\]]', prompt)
            device['hostname'] = m.group(1) if m else sanitize_filename(device['host'])

            blocks = []

            if config_set and not use_manual_pager:
                # 配置模式（仅已识别厂商支持 send_config_set）
                out = conn.send_config_set(cmds)
                blocks.append(out)
            else:
                for cmd in cmds:
                    if use_manual_pager:
                        out = send_command_termserver(conn, cmd)
                    else:
                        out = conn.send_command(
                            cmd,
                            read_timeout=int(device.get('readtime', 20)),
                            strip_prompt=False,
                            strip_command=False,
                        )
                    blocks.append(f"{'=' * 60}\n命令: {cmd}\n{'-' * 60}\n{out}")

            return "\n\n".join(blocks)

    except Exception as e:
        log_error(device['host'], f"[命令执行异常] {str(e)}")
        return None


# ---------------------------------------------------------------------------
# 结果保存（txt 文本）
# ---------------------------------------------------------------------------
def save_output(device: Dict[str, str], content: str, output_dir: str) -> None:
    """每台设备保存为独立 txt 文件。"""
    if not content:
        return
    host = device.get('hostname') or sanitize_filename(device['host'])
    ts = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
    filename = f"{sanitize_filename(host)}_{sanitize_filename(device['host'])}_{ts}.txt"
    filepath = os.path.join(output_dir, filename)

    header = (
        f"设备: {device['host']}\n"
        f"主机名: {device.get('hostname', '')}\n"
        f"类型: {device['device_type']}\n"
        f"时间: {datetime.datetime.now():%Y-%m-%d %H:%M:%S}\n"
        f"{'#' * 60}\n\n"
    )

    with write_lock:
        with open(filepath, "w", encoding="utf-8") as f:
            f.write(header + content + "\n")
    print(f"{device['host']} [OK] 结果已保存 -> {filepath}")


# ---------------------------------------------------------------------------
# 单设备任务
# ---------------------------------------------------------------------------
def process_device(device: Dict[str, str], config_set: bool, output_dir: str) -> bool:
    """单台设备完整流程：连接 -> 执行 -> 保存。"""
    output = execute_commands(device, config_set)
    if output is None:
        return False
    save_output(device, output, output_dir)
    return True


# ---------------------------------------------------------------------------
# 主流程
# ---------------------------------------------------------------------------
def main() -> None:
    parser = argparse.ArgumentParser(description="批量网络设备命令执行工具（每台设备命令来自 Excel）")
    parser.add_argument('-i', '--input', required=True, help="Excel 设备清单文件")
    parser.add_argument('-s', '--sheet', default='Sheet1', help="工作表名（默认 Sheet1）")
    parser.add_argument('-o', '--output', default='output', help="结果输出目录（默认 output）")
    parser.add_argument('-t', '--threads', type=int, default=DEFAULT_THREADS, help="并发线程数")
    parser.add_argument('--config', action='store_true', help="以配置模式下发命令（send_config_set）")
    args = parser.parse_args()

    if not os.path.isfile(args.input):
        print(f"输入文件不存在: {args.input}")
        sys.exit(1)

    os.makedirs(args.output, exist_ok=True)

    devices = load_excel(args.input, args.sheet)
    print(f"共加载 {len(devices)} 台设备，线程数 {args.threads}")

    success = 0
    with ThreadPoolExecutor(max_workers=args.threads, initializer=thread_initializer) as executor:
        futures = {
            executor.submit(process_device, dev, args.config, args.output): dev
            for dev in devices
        }
        for future in tqdm(as_completed(futures), total=len(futures), desc="执行进度"):
            dev = futures[future]
            try:
                if future.result():
                    success += 1
            except Exception as e:
                log_error(dev['host'], f"任务异常: {str(e)}")

    print(f"\n完成：成功 {success} / 总计 {len(devices)}")


if __name__ == "__main__":
    main()
