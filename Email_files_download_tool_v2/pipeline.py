"""主流程：扫描 -> 过滤 -> 去重 -> 下载 -> 记账。"""

from __future__ import annotations

import hashlib
import re
from datetime import datetime, timedelta
from pathlib import Path
from typing import Dict, Optional

from config_loader import Config
from extractor import build_extract_dir, extract_archive, is_archive
from outlook_client import OutlookSession
from state_store import RunLock, StateStore
from utils import (
    clamp_path_length,
    now_utc_naive,
    resolve_collision,
    same_content,
    sanitize_component,
    sanitize_filename,
    sender_matches,
    setup_logging,
    today_window_utc,
    local_date_window_utc,
    unique_path,
    utc_to_local_naive,
)


def message_key(session: OutlookSession, entry_id: str, item) -> str:
    """优先用 Internet Message ID 去重（邮件被移动也不会变），取不到再退回 EntryID。"""
    basis = session.internet_message_id(item) or entry_id
    return hashlib.sha1(str(basis).encode("utf-8")).hexdigest()[:20]


def build_target_dir(cfg: Config, sender: str, received_utc, subject: str) -> Path:
    local = utc_to_local_naive(received_utc)
    local_part, _, domain = sender.partition("@")
    mapping = {
        "{sender}": sanitize_component(sender or "unknown"),
        "{sender_local}": sanitize_component(local_part or "unknown"),
        "{sender_domain}": sanitize_component(domain or "unknown"),
        "{subject}": sanitize_component(subject or "no-subject"),
        "{date}": local.strftime("%Y-%m-%d"),
        "{year}": local.strftime("%Y"),
        "{month}": local.strftime("%m"),
        "{day}": local.strftime("%d"),
    }
    parts = []
    for raw_part in re.split(r"[\\/]+", cfg.download.subfolder_template or ""):
        part = raw_part
        for token, value in mapping.items():
            part = part.replace(token, value)
        part = sanitize_component(part, max_len=120)
        if part:
            parts.append(part)
    return Path(cfg.download.target_dir, *parts)


def _extension_allowed(name: str, allowed) -> bool:
    if not allowed:
        return True
    suffix = Path(name).suffix.lower().lstrip(".")
    return suffix in allowed


# 邮件正文里的任务描述行，全角/半角冒号都支持
JOB_DESC_RE = re.compile(r"Report\s*Job\s*Description\s*[：:]\s*(\S[^\r\n]*)", re.IGNORECASE)


def job_description(body: str) -> str:
    """从正文提取 Report Job Description 的值；找不到返回空串。"""
    match = JOB_DESC_RE.search(body or "")
    return match.group(1).strip() if match else ""


def _desc_stem(desc: str) -> str:
    """最终文件名主干（不含扩展名）。

    描述里自带扩展名（如 bbbb.csv）时剥离掉——真正的扩展名由解压出的文件决定，
    否则会出现 bbbb.csv.pdf 这种名字。
    """
    stem = sanitize_filename(desc)
    m = re.search(r"\.[A-Za-z][A-Za-z0-9]{0,5}$", stem)
    if m:
        stem = stem[: m.start()]
    return stem or "unnamed"


def _desc_target(src: Path, desc: str, on_duplicate: str) -> Optional[Path]:
    """按正文描述生成重命名目标路径（压缩包与解压文件共用同一规则）。"""
    stem = _desc_stem(desc)
    suffix = src.suffix
    if not stem.lower().endswith(suffix.lower()):
        stem += suffix
    return unique_path(clamp_path_length(src.parent / stem), on_duplicate)


def _safe_desc_target(src: Path, desc: str, logger) -> Optional[Path]:
    """按描述生成目标路径，并处理撞名。

    规则：目标同名时，**内容相同才覆盖**（重复下载的同一份报告，不会出现 bbbb_1），
    **内容不同则保留两份并加序号**（不同报告撞名，绝不静默丢文件）。
    """
    base = _desc_target(src, desc, "overwrite")
    if base is None or base == src or not base.exists():
        return base
    final = resolve_collision(base, src)
    if final == base:
        logger.info("目标同名且内容一致，覆盖：%s", base.name)
    else:
        logger.warning("目标同名但内容不同，保留两份：%s（另存为 %s）",
                       base.name, final.name)
    return final


def _rename_by_job_desc(paths, desc: str, logger) -> None:
    """zip 只解出一个文件时，把文件重命名为正文里的 Report Job Description。"""
    if len(paths) != 1:
        logger.info("解压出 %d 个文件（非单个），跳过按正文重命名", len(paths))
        return
    src = Path(paths[0])
    target = _safe_desc_target(src, desc, logger)
    if target is None:
        logger.info("跳过重命名（目标同名文件已存在）：%s", src.name)
        return
    if target != src:
        # replace 在 Windows 上可覆盖已存在的目标文件（rename 会报 FileExistsError）
        src.replace(target)
        logger.info("已按 Report Job Description 重命名：%s -> %s", src.name, target.name)


def run_once(
    cfg: Config,
    dry_run: bool = False,
    lookback_days: Optional[int] = None,
    logger=None,
    only_today: Optional[bool] = None,
    specific_date: Optional[str] = None,
) -> Dict[str, int]:
    logger = logger or setup_logging(cfg.logging.dir, cfg.logging.level)
    stats = {
        "messages_matched": 0,
        "attachments_seen": 0,
        "downloaded": 0,
        "skipped_already": 0,
        "skipped_filter": 0,
        "skipped_duplicate": 0,
        "failed": 0,
        "bytes": 0,
        "extracted": 0,
        "extract_files": 0,
        "extract_failed": 0,
    }

    lock = RunLock(cfg.base_dir / "data" / "run.lock")
    if not lock.acquire():
        logger.warning("上一次运行仍在进行中（锁文件：%s），本次跳过", lock.path)
        return stats

    state = StateStore.load(cfg.state.file, cfg.state.retention_days)

    today_only = cfg.scan.only_today if only_today is None else only_today
    days = lookback_days if lookback_days is not None else cfg.scan.lookback_days
    cutoff = now_utc_naive() - timedelta(days=days)
    until: Optional[datetime] = None
    if specific_date:
        # 指定日期优先级最高：补抓某一天的邮件（UI / 命令行手动触发用）
        try:
            y, m, d = (int(x) for x in str(specific_date).split("-"))
        except ValueError as exc:
            raise ValueError(f"指定日期格式应为 YYYY-MM-DD，当前是 {specific_date!r}") from exc
        cutoff, until = local_date_window_utc(y, m, d, cfg.scan.day_start_hour)
        window_desc = (
            f"指定日期 {specific_date}（本地时间 "
            f"{utc_to_local_naive(cutoff):%Y-%m-%d %H:%M:%S} ~ {utc_to_local_naive(until):%Y-%m-%d %H:%M:%S}）"
        )
    elif today_only:
        # 取两个下界里更晚的那个：既满足"只看今天"，也不会因为 lookback_days
        # 配得很大而把今天之前的邮件又扫进来。
        today_start, today_end = today_window_utc(cfg.scan.day_start_hour)
        cutoff = max(cutoff, today_start)
        until = today_end
        window_desc = (
            f"仅今天（本地时间 "
            f"{utc_to_local_naive(today_start):%Y-%m-%d %H:%M:%S} ~ {utc_to_local_naive(today_end):%Y-%m-%d %H:%M:%S}）"
        )
    else:
        window_desc = f"回溯 {days} 天（截止 {cutoff} UTC）"

    try:
        logger.info("=" * 60)
        logger.info("开始运行 | dry_run=%s | %s", dry_run, window_desc)
        with OutlookSession.connect(timeout=cfg.outlook.timeout_seconds) as session:
            folder = session.resolve_folder(cfg.outlook.mailbox, cfg.outlook.folder)
            try:
                logger.info("目标文件夹：%s（共 %s 封）", folder.Name, folder.Items.Count)
            except Exception:
                logger.info("目标文件夹已定位")

            for msg in session.iter_messages(
                folder,
                cutoff,
                max_items=cfg.scan.max_messages_per_run,
                only_mail=cfg.outlook.only_mail_items,
                senders=cfg.senders,
                until_utc=until,
            ):
                stats["messages_matched"] += 1
                mkey = message_key(session, msg.entry_id, msg.item)
                if not sender_matches(msg.sender, cfg.senders):
                    continue

                failed_here = 0
                saved_here = 0
                for att in session.iter_attachments(
                    msg,
                    skip_inline=cfg.download.skip_inline_images,
                    save_embedded=cfg.download.save_embedded_msg,
                ):
                    stats["attachments_seen"] += 1
                    akey = f"{mkey}:{att.index}"
                    if state.has_attachment(akey):
                        # 记账过不代表文件还在：可能被人工移动/删除。
                        # 文件还在 -> 跳过；文件丢了 -> 当作没下过，重新下载。
                        rec = state.get_attachment(akey) or {}
                        rec_path = Path(rec.get("file", ""))
                        if rec_path.name and rec_path.exists():
                            stats["skipped_already"] += 1
                            continue
                        logger.info("状态有记录但文件已不在磁盘，将重新下载：%s", att.name)

                    if not _extension_allowed(att.name, cfg.download.allowed_extensions):
                        logger.info("跳过（扩展名不在白名单）：%s", att.name)
                        stats["skipped_filter"] += 1
                        continue
                    if att.size and cfg.download.min_size_bytes and att.size < cfg.download.min_size_bytes:
                        stats["skipped_filter"] += 1
                        continue
                    if att.size and cfg.download.max_size_bytes and att.size > cfg.download.max_size_bytes:
                        logger.info("跳过（超过大小上限 %d 字节）：%s", cfg.download.max_size_bytes, att.name)
                        stats["skipped_filter"] += 1
                        continue

                    target_dir = build_target_dir(cfg, msg.sender, msg.received_utc, msg.subject)
                    final_path = unique_path(
                        clamp_path_length(target_dir / sanitize_filename(att.name)),
                        cfg.download.on_duplicate,
                    )
                    if final_path is None:
                        logger.info("跳过（目标已存在同名文件）：%s", target_dir / att.name)
                        stats["skipped_duplicate"] += 1
                        continue

                    # 解压去向提前算好，dry-run 也要能打印出来
                    dest_dir = None
                    if cfg.extract.enabled and is_archive(att.name, cfg.extract.extensions):
                        dest_dir = build_extract_dir(
                            cfg.extract.subfolder_template,
                            cfg.extract.target_dir,
                            msg.sender,
                            msg.received_utc,
                            msg.subject,
                            Path(att.name).stem,
                        )

                    if dry_run:
                        logger.info("[DRY-RUN] 将下载 %s <- %s", final_path, msg.subject)
                        if dest_dir is not None:
                            logger.info("[DRY-RUN] 将解压 -> %s", dest_dir)
                        saved_here += 1
                        continue

                    try:
                        size = session.save_attachment(att, final_path)
                    except Exception as exc:
                        failed_here += 1
                        stats["failed"] += 1
                        logger.error("下载失败 %s（来自 %.60s）：%s", att.name, msg.subject, exc)
                        continue

                    stats["downloaded"] += 1
                    stats["bytes"] += size
                    saved_here += 1
                    state.mark_attachment(akey, {
                        "file": str(final_path),
                        "name": att.name,
                        "size": size,
                        "sender": msg.sender,
                        "subject": msg.subject,
                        "received_utc": msg.received_utc.isoformat() + "Z",
                    })
                    logger.info("已保存 %s（%d 字节）", final_path, size)

                    desc = ""
                    if cfg.extract.rename_by_job_desc:
                        desc = job_description(getattr(msg, "body", ""))

                    if dest_dir is not None:
                        outcome = extract_archive(
                            final_path,
                            dest_dir,
                            overwrite=cfg.extract.on_duplicate,
                            max_files=cfg.extract.max_files,
                            max_total_bytes=cfg.extract.max_total_bytes,
                            max_ratio=cfg.extract.max_ratio,
                            # 提前把最终名交给解压层：写盘时就按这个名字落，
                            # 避免"先按 zip 内原名覆盖掉上一份、事后才发现撞名"导致丢文件
                            final_stem=_desc_stem(desc) if desc else None,
                            logger=logger,
                        )
                        if outcome["ok"]:
                            stats["extracted"] += 1
                            stats["extract_files"] += int(outcome["files"])
                            logger.info("已解压 %s 个文件 -> %s", outcome["files"], dest_dir)
                            if desc:
                                _rename_by_job_desc(
                                    outcome.get("paths") or [], desc, logger,
                                )
                            if cfg.extract.delete_archive:
                                try:
                                    final_path.unlink()
                                    logger.info("已删除原压缩包 %s", final_path)
                                except Exception as exc:
                                    logger.warning("删除原压缩包失败 %s：%s", final_path, exc)
                        else:
                            stats["extract_failed"] += 1
                            logger.error("解压失败 %s：%s", final_path, outcome["message"])

                    # 压缩包本身也按同一规则重命名（如 bbbb.csv -> bbbb.zip），
                    # 并同步更新记账路径，避免下次运行被判"文件丢失"而重复下载。
                    if desc and final_path.exists():
                        arc_target = _safe_desc_target(final_path, desc, logger)
                        if arc_target is not None and arc_target != final_path:
                            try:
                                final_path.replace(arc_target)
                                logger.info("已按 Report Job Description 重命名压缩包：%s -> %s",
                                            final_path.name, arc_target.name)
                                final_path = arc_target
                                rec = state.get_attachment(akey) or {}
                                if rec:
                                    rec = dict(rec)
                                    rec["file"] = str(arc_target)
                                    state.mark_attachment(akey, rec)
                            except Exception as exc:
                                logger.warning("重命名压缩包失败 %s：%s", final_path, exc)

                if saved_here or failed_here:
                    if failed_here == 0 and not dry_run:
                        state.mark_message(mkey, {
                            "subject": msg.subject,
                            "sender": msg.sender,
                            "received_utc": msg.received_utc.isoformat() + "Z",
                            "attachments": saved_here,
                        })

    except Exception as exc:
        stats["failed"] += 1
        logger.exception("运行失败：%s", exc)
    finally:
        if not dry_run:
            removed = state.prune()
            if removed:
                logger.info("清理过期状态记录 %d 条", removed)
            try:
                state.save()
            except Exception as exc:
                logger.error("状态文件保存失败：%s", exc)
        lock.release()

    logger.info(
        "运行结束 | 命中邮件 %d 封 | 附件 %d 个 | 下载 %d | 已存在 %d | 过滤 %d | 同名跳过 %d | 失败 %d | %.2f MB",
        stats["messages_matched"], stats["attachments_seen"], stats["downloaded"],
        stats["skipped_already"], stats["skipped_filter"], stats["skipped_duplicate"],
        stats["failed"], stats["bytes"] / 1024 / 1024,
    )
    if cfg.extract.enabled:
        logger.info(
            "解压结果 | 成功 %d 个压缩包 | 解压出 %d 个文件 | 失败 %d",
            stats["extracted"], stats["extract_files"], stats["extract_failed"],
        )
    return stats
