"""通用工具：日志、时间、文件名清洗、路径处理。"""

from __future__ import annotations

import logging
import os
import re
import sys
import unicodedata
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import List, Optional

ILLEGAL_CHARS_RE = re.compile(r'[<>:"/\\|?*\x00-\x1f]')
RESERVED_NAMES = {"CON", "PRN", "AUX", "NUL"} | {f"COM{i}" for i in range(1, 10)} | {f"LPT{i}" for i in range(1, 10)}


def force_utf8_console() -> None:
    """计划任务下 stdout 可能是 GBK，强制改成 UTF-8 避免中文乱码。"""
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8", errors="replace")  # type: ignore[union-attr]
        except Exception:
            pass


def setup_logging(log_dir: Path, level: str = "INFO", name: str = "maildl") -> logging.Logger:
    log_dir.mkdir(parents=True, exist_ok=True)
    logger = logging.getLogger(name)
    logger.setLevel(getattr(logging, str(level).upper(), logging.INFO))
    logger.propagate = False
    for handler in list(logger.handlers):
        logger.removeHandler(handler)
        try:
            handler.close()
        except Exception:
            pass

    fmt = logging.Formatter("%(asctime)s [%(levelname)s] %(message)s", "%Y-%m-%d %H:%M:%S")
    file_handler = logging.FileHandler(log_dir / f"run-{datetime.now():%Y%m%d}.log", encoding="utf-8")
    file_handler.setFormatter(fmt)
    stream_handler = logging.StreamHandler(sys.stdout)
    stream_handler.setFormatter(fmt)
    logger.addHandler(file_handler)
    logger.addHandler(stream_handler)
    return logger


def now_utc_naive() -> datetime:
    return datetime.now(timezone.utc).replace(tzinfo=None)


def to_utc_naive(dt: Optional[datetime]) -> Optional[datetime]:
    """Outlook 返回的 pywintypes.datetime 通常带 UTC 时区，统一转成 naive UTC 便于比较。"""
    if dt is None:
        return None
    if dt.tzinfo is None:
        return dt
    return dt.astimezone(timezone.utc).replace(tzinfo=None)


def local_day_start_utc(offset_days: int = 0, start_hour: int = 0) -> datetime:
    """本地时区第 offset_days 天的 start_hour:00:00，换算成 naive UTC。

    offset_days=0 -> 今天；1 -> 明天。start_hour 允许把"一天"的起点
    从 00:00 挪到别处（例如 8 表示 08:00 ~ 次日 08:00 算一天）。

    必须按本地日期算，直接用 UTC 会把东八区当天 00:00~08:00 的邮件算成前一天。
    """
    now_local = datetime.now().astimezone()
    midnight = (now_local + timedelta(days=offset_days)).replace(
        hour=0, minute=0, second=0, microsecond=0
    )
    return (midnight + timedelta(hours=start_hour)).astimezone(timezone.utc).replace(tzinfo=None)


def local_midnight_utc(offset_days: int = 0) -> datetime:
    """本地时区第 offset_days 天的 00:00:00，换算成 naive UTC（保留兼容）。"""
    return local_day_start_utc(offset_days, 0)


def today_window_utc(start_hour: int = 0):
    """返回 (本地今天 start_hour:00 UTC, 到次日同一时刻前 1 秒 UTC)。

    默认 start_hour=0 即 00:00:00 ~ 23:59:59，与直观的"仅今天"一致；
    配成 8 则变成 08:00:00 ~ 次日 07:59:59。
    """
    return (
        local_day_start_utc(0, start_hour),
        local_day_start_utc(1, start_hour) - timedelta(seconds=1),
    )


def local_date_window_utc(year: int, month: int, day: int, start_hour: int = 0):
    """指定本地日期的 (起点 UTC, 终点 UTC)，用于"补抓某一天"的邮件。"""
    start_local = datetime(year, month, day) + timedelta(hours=start_hour)
    start_utc = start_local.astimezone(timezone.utc).replace(tzinfo=None)
    end_utc = (start_local + timedelta(days=1) - timedelta(seconds=1)) \
        .astimezone(timezone.utc).replace(tzinfo=None)
    return start_utc, end_utc


def utc_to_local_naive(dt: datetime) -> datetime:
    if dt.tzinfo is None:
        dt = dt.replace(tzinfo=timezone.utc)
    return dt.astimezone().replace(tzinfo=None)


def sanitize_component(name: str, max_len: int = 80) -> str:
    """清洗路径中的单个目录名/文件名。"""
    text = unicodedata.normalize("NFC", str(name or "")).strip()
    text = ILLEGAL_CHARS_RE.sub("_", text)
    text = re.sub(r"\s+", " ", text).strip(" .")
    if text.split(".")[0].upper() in RESERVED_NAMES:
        text = "_" + text
    if not text:
        text = "unnamed"
    if len(text) > max_len:
        text = text[:max_len].rstrip(" .") or "unnamed"
    return text


def sanitize_filename(name: str, max_len: int = 120) -> str:
    """清洗附件名，尽量保留扩展名。"""
    text = unicodedata.normalize("NFC", str(name or "")).strip()
    text = ILLEGAL_CHARS_RE.sub("_", text)
    text = re.sub(r"\s+", " ", text).strip(" .")
    if text.split(".")[0].upper() in RESERVED_NAMES:
        text = "_" + text
    if not text:
        return "unnamed"
    if len(text) > max_len:
        root, ext = os.path.splitext(text)
        if len(ext) > 12:  # 不像是真扩展名，整体截断
            return text[:max_len]
        keep = max(1, max_len - len(ext))
        text = root[:keep].rstrip(" .") + ext
    return text


def clamp_path_length(path: Path, limit: int = 240) -> Path:
    """Windows 传统 API 有 260 字符限制，超长时截断文件名。"""
    text = str(path)
    if len(text) <= limit:
        return path
    over = len(text) - limit
    stem, suffix = path.stem, path.suffix
    new_stem = stem[: max(1, len(stem) - over - 1)].rstrip(" .") or "f"
    return path.parent / (new_stem + suffix)


def unique_path(path: Path, mode: str = "overwrite") -> Optional[Path]:
    """按配置处理同名文件，返回 None 表示跳过。默认直接覆盖。"""
    if not path.exists():
        return path
    if mode == "overwrite":
        return path
    if mode == "skip":
        return None
    stem, suffix = path.stem, path.suffix
    for n in range(1, 1000):
        candidate = path.parent / f"{stem}_{n}{suffix}"
        if not candidate.exists():
            return candidate
    return path.parent / f"{stem}_{os.getpid()}{suffix}"


def file_hash(path: Path, chunk: int = 1024 * 1024) -> str:
    import hashlib
    h = hashlib.sha256()
    with open(path, "rb") as fh:
        while True:
            block = fh.read(chunk)
            if not block:
                break
            h.update(block)
    return h.hexdigest()


def same_content(a: Path, b: Path) -> bool:
    """内容是否完全相同（先比大小，再比 sha256）。"""
    try:
        if a.stat().st_size != b.stat().st_size:
            return False
        return file_hash(a) == file_hash(b)
    except OSError:
        return False


def resolve_collision(target: Path, incoming: Path) -> Path:
    """决定 incoming 应落到哪个名字。

    目标不存在        -> 直接用目标名
    内容相同          -> 覆盖目标名（重复下载的同一份报告，不应产生 aaaa_1）
    内容不同          -> 另存为 名字_N（不同报告撞名，绝不静默丢文件）
    """
    if not target.exists():
        return target
    if same_content(incoming, target):
        return target
    stem, suffix = target.stem, target.suffix
    for n in range(1, 1000):
        candidate = target.with_name(f"{stem}_{n}{suffix}")
        if not candidate.exists() or same_content(incoming, candidate):
            return candidate
    return target.with_name(f"{stem}_{os.getpid()}{suffix}")


def normalize_addr(addr: str) -> str:
    return str(addr or "").strip().lower()


def sender_matches(address: str, patterns: List[str]) -> bool:
    """发件人匹配：支持精确地址与 *@domain.com 域名通配。"""
    addr = normalize_addr(address)
    if not addr:
        return False
    for pattern in patterns:
        rule = normalize_addr(pattern)
        if not rule:
            continue
        if rule.startswith("*@"):
            if addr.endswith(rule[1:]):
                return True
        elif "*" in rule:
            regex = re.escape(rule).replace(r"\*", ".*")
            if re.fullmatch(regex, addr):
                return True
        elif addr == rule:
            return True
    return False
