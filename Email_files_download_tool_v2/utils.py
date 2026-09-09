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


def local_midnight_utc(offset_days: int = 0) -> datetime:
    """本地时区第 offset_days 天的 00:00:00，换算成 naive UTC。

    offset_days=0 -> 今天零点；1 -> 明天零点。
    用于实现"只处理当天邮件"：必须按本地日期算，直接用 UTC 会把
    东八区当天 00:00~08:00 的邮件算成前一天。
    """
    now_local = datetime.now().astimezone()
    midnight = (now_local + timedelta(days=offset_days)).replace(
        hour=0, minute=0, second=0, microsecond=0
    )
    return midnight.astimezone(timezone.utc).replace(tzinfo=None)


def today_window_utc():
    """返回 (本地今天 00:00 UTC, 本地今天 23:59:59 UTC)。

    上界取当天 23:59:59（而不是次日零点），与用户看到的
    "仅今天 = 当天 00:00 ~ 23:59" 的直观理解保持一致。
    """
    return local_midnight_utc(0), local_midnight_utc(1) - timedelta(seconds=1)


def local_date_window_utc(year: int, month: int, day: int):
    """指定本地日期的 (当天 00:00 UTC, 当天 23:59:59 UTC)，用于"补抓某一天"的邮件。"""
    start_local = datetime(year, month, day)
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
