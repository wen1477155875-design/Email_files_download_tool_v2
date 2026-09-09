"""配置加载与校验。所有相对路径都相对配置文件所在目录解析。"""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import List

import yaml


class ConfigError(RuntimeError):
    pass


def _resolve(base: Path, value: str) -> Path:
    p = Path(str(value))
    return p if p.is_absolute() else (base / p)


@dataclass
class OutlookSection:
    mailbox: str = ""
    folder: str = "收件箱"
    only_mail_items: bool = True
    timeout_seconds: int = 120


@dataclass
class ScanSection:
    lookback_days: int = 7
    max_messages_per_run: int = 500
    # 只处理"本地日期 == 今天"的邮件（收件当天就下载，错过就不管）
    only_today: bool = False


@dataclass
class DownloadSection:
    target_dir: Path = Path("downloads")
    subfolder_template: str = "{sender}/{date}"
    skip_inline_images: bool = True
    save_embedded_msg: bool = False
    on_duplicate: str = "rename"
    min_size_bytes: int = 0
    max_size_bytes: int = 100 * 1024 * 1024
    allowed_extensions: List[str] = field(default_factory=list)


@dataclass
class ExtractSection:
    """下载完成后，把 zip 附件解压到另一个目录。"""
    enabled: bool = False
    target_dir: Path = Path("extracted")
    subfolder_template: str = "{sender}/{date}/{archive}"
    # 只解压这些扩展名（小写，不含点）
    extensions: List[str] = field(default_factory=lambda: ["zip"])
    on_duplicate: str = "rename"
    # 解压成功后删除原压缩包（省空间，但不可逆）
    delete_archive: bool = False
    # 正文里解析 "Report Job Description：xxxx"，把解压出的唯一文档重命名为 xxxx
    rename_by_job_desc: bool = True
    max_files: int = 2000
    max_total_bytes: int = 2 * 1024 * 1024 * 1024
    max_ratio: float = 200.0


@dataclass
class StateSection:
    file: Path = Path("data/state.json")
    retention_days: int = 180


@dataclass
class LoggingSection:
    dir: Path = Path("data/logs")
    level: str = "INFO"


@dataclass
class ScheduleSection:
    task_name: str = "EmailAttachmentDownloader"
    mode: str = "daily"
    time: str = "09:00"
    interval_minutes: int = 30
    start_when_available: bool = True
    execution_time_limit: str = "PT1H"


@dataclass
class Config:
    base_dir: Path
    config_path: Path
    outlook: OutlookSection
    senders: List[str]
    scan: ScanSection
    download: DownloadSection
    extract: ExtractSection
    state: StateSection
    logging: LoggingSection
    schedule: ScheduleSection


def load_config(path: Path) -> Config:
    path = Path(path).resolve()
    if not path.exists():
        raise ConfigError(f"配置文件不存在：{path}")
    try:
        raw = yaml.safe_load(path.read_text(encoding="utf-8")) or {}
    except yaml.YAMLError as exc:
        raise ConfigError(f"配置文件 YAML 解析失败：{exc}") from exc
    if not isinstance(raw, dict):
        raise ConfigError("配置文件内容必须是一个键值映射")

    base = path.parent
    outlook_raw = raw.get("outlook") or {}
    scan_raw = raw.get("scan") or {}
    dl_raw = raw.get("download") or {}
    ex_raw = raw.get("extract") or {}
    state_raw = raw.get("state") or {}
    log_raw = raw.get("logging") or {}
    sch_raw = raw.get("schedule") or {}

    senders = [str(s).strip() for s in (raw.get("senders") or []) if str(s).strip()]
    if not senders:
        raise ConfigError("senders 至少配置一个发件人，否则不会下载任何附件")

    on_duplicate = str(dl_raw.get("on_duplicate", "rename")).lower()
    if on_duplicate not in ("rename", "overwrite", "skip"):
        raise ConfigError("download.on_duplicate 只能是 rename / overwrite / skip")

    ex_dup = str(ex_raw.get("on_duplicate", "rename")).lower()
    if ex_dup not in ("rename", "overwrite", "skip"):
        raise ConfigError("extract.on_duplicate 只能是 rename / overwrite / skip")

    mode = str(sch_raw.get("mode", "daily")).lower()
    if mode not in ("daily", "minutes"):
        raise ConfigError("schedule.mode 只能是 daily / minutes")

    extentions = [str(e).strip().lower().lstrip(".") for e in (dl_raw.get("allowed_extensions") or []) if str(e).strip()]

    return Config(
        base_dir=base,
        config_path=path,
        outlook=OutlookSection(
            mailbox=str(outlook_raw.get("mailbox", "") or "").strip(),
            folder=str(outlook_raw.get("folder", "收件箱") or "收件箱"),
            only_mail_items=bool(outlook_raw.get("only_mail_items", True)),
            timeout_seconds=int(outlook_raw.get("timeout_seconds", 120)),
        ),
        senders=senders,
        scan=ScanSection(
            lookback_days=int(scan_raw.get("lookback_days", 7)),
            max_messages_per_run=int(scan_raw.get("max_messages_per_run", 500)),
            only_today=bool(scan_raw.get("only_today", False)),
        ),
        download=DownloadSection(
            target_dir=_resolve(base, dl_raw.get("target_dir", "downloads")),
            subfolder_template=str(dl_raw.get("subfolder_template", "{sender}/{date}")),
            skip_inline_images=bool(dl_raw.get("skip_inline_images", True)),
            save_embedded_msg=bool(dl_raw.get("save_embedded_msg", False)),
            on_duplicate=on_duplicate,
            min_size_bytes=int(dl_raw.get("min_size_bytes", 0)),
            max_size_bytes=int(dl_raw.get("max_size_bytes", 100 * 1024 * 1024)),
            allowed_extensions=extentions,
        ),
        extract=ExtractSection(
            enabled=bool(ex_raw.get("enabled", False)),
            target_dir=_resolve(base, ex_raw.get("target_dir", "extracted")),
            subfolder_template=str(ex_raw.get("subfolder_template", "{sender}/{date}/{archive}")),
            extensions=[str(e).strip().lower().lstrip(".")
                        for e in (ex_raw.get("extensions") or ["zip"]) if str(e).strip()] or ["zip"],
            on_duplicate=str(ex_raw.get("on_duplicate", "overwrite")).lower(),
            delete_archive=bool(ex_raw.get("delete_archive", False)),
            rename_by_job_desc=bool(ex_raw.get("rename_by_job_desc", True)),
            max_files=int(ex_raw.get("max_files", 2000)),
            max_total_bytes=int(ex_raw.get("max_total_bytes", 2 * 1024 * 1024 * 1024)),
            max_ratio=float(ex_raw.get("max_ratio", 200.0)),
        ),
        state=StateSection(
            file=_resolve(base, state_raw.get("file", "data/state.json")),
            retention_days=int(state_raw.get("retention_days", 180)),
        ),
        logging=LoggingSection(
            dir=_resolve(base, log_raw.get("dir", "data/logs")),
            level=str(log_raw.get("level", "INFO")),
        ),
        schedule=ScheduleSection(
            task_name=str(sch_raw.get("task_name", "EmailAttachmentDownloader")),
            mode=mode,
            time=str(sch_raw.get("time", "09:00")),
            interval_minutes=int(sch_raw.get("interval_minutes", 30)),
            start_when_available=bool(sch_raw.get("start_when_available", True)),
            execution_time_limit=str(sch_raw.get("execution_time_limit", "PT1H")),
        ),
    )
