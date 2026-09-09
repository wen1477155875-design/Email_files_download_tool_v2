"""邮件附件自动下载工具 - 命令行入口。

用法：
  python main.py check                 检查 Outlook 连接与配置是否正确
  python main.py run [--dry-run]       执行一次下载
  python main.py task install          注册 Windows 计划任务
  python main.py task status | uninstall | run-now
  python main.py state show | reset    查看/清空已处理记录
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

from utils import force_utf8_console

BASE_DIR = Path(__file__).resolve().parent
DEFAULT_CONFIG = BASE_DIR / "config.yaml"


def _load(args):
    from config_loader import load_config
    return load_config(Path(args.config) if args.config else DEFAULT_CONFIG)


def cmd_run(args) -> int:
    from config_loader import ConfigError
    from pipeline import run_once
    from utils import setup_logging
    try:
        cfg = _load(args)
    except ConfigError as exc:
        print(f"[配置错误] {exc}")
        return 1
    logger = setup_logging(cfg.logging.dir, cfg.logging.level)
    stats = run_once(
        cfg,
        dry_run=args.dry_run,
        lookback_days=args.lookback_days,
        logger=logger,
        only_today=args.only_today,
    )
    return 1 if stats.get("failed") else 0


def cmd_check(args) -> int:
    from config_loader import ConfigError
    try:
        cfg = _load(args)
    except ConfigError as exc:
        print(f"[配置错误] {exc}")
        return 1

    print("== 基本环境 ==")
    print(f"Python        : {sys.executable}")
    try:
        import win32com  # type: ignore
        print(f"pywin32       : {win32com.__version__ if hasattr(win32com, '__version__') else '已安装'}")
    except Exception as exc:
        print(f"pywin32       : 未安装或不可用 -> {exc}")
        print("  请先执行：pip install -r requirements.txt")
        return 1

    print("\n== 配置解析 ==")
    print(f"配置文件      : {cfg.config_path}")
    print(f"邮箱          : {cfg.outlook.mailbox or '(默认邮箱)'}")
    print(f"扫描文件夹    : {cfg.outlook.folder}")
    print(f"关注发件人    : {', '.join(cfg.senders)}")
    print(f"时间范围      : {'仅今天' if cfg.scan.only_today else f'最近 {cfg.scan.lookback_days} 天'}")
    print(f"同类邮件上限  : {cfg.scan.max_messages_per_run} 封/次")
    print(f"保存目录      : {cfg.download.target_dir}")
    print(f"子目录模板    : {cfg.download.subfolder_template}")
    if cfg.extract.enabled:
        print(f"解压到        : {cfg.extract.target_dir}")
        print(f"解压子目录    : {cfg.extract.subfolder_template}")
        print(f"解压类型      : {', '.join(cfg.extract.extensions)}"
              f"{'（解压后删除原包）' if cfg.extract.delete_archive else ''}")
    else:
        print("解压          : 未启用")
    print(f"状态文件      : {cfg.state.file}")

    print("\n== Outlook 连接 ==")
    try:
        from outlook_client import OutlookSession
        with OutlookSession.connect(timeout=cfg.outlook.timeout_seconds) as session:
            print("连接          : 成功")
            mailboxes = session.list_mailboxes()
            print(f"已配置账号    : {mailboxes if mailboxes else '(一个都没有，经典版 Outlook 还没配邮箱)'}")
            if cfg.outlook.mailbox and mailboxes:
                hit = cfg.outlook.mailbox.strip().lower() in [m.strip().lower() for m in mailboxes]
                print(f"配置匹配      : {'命中' if hit else '未命中！config 里的 outlook.mailbox 不在上面的账号列表里'}")
            folder = session.resolve_folder(cfg.outlook.mailbox, cfg.outlook.folder)
            print(f"目标文件夹    : {folder.Name}")
            try:
                print(f"文件夹邮件数  : {folder.Items.Count}")
            except Exception as exc:
                print(f"文件夹邮件数  : 读取失败 -> {exc}")
    except Exception as exc:
        print(f"连接          : 失败 -> {exc}")
        return 1

    print("\n检查通过，可以执行：python main.py run --dry-run")
    return 0


def cmd_task(args) -> int:
    from config_loader import ConfigError
    try:
        cfg = _load(args)
    except ConfigError as exc:
        print(f"[配置错误] {exc}")
        return 1

    from scheduler import SchedulerError, install_task, run_task_now, task_status, uninstall_task
    name = cfg.schedule.task_name
    try:
        if args.action == "install":
            install_task(cfg, logger=None)
            print(f"\n查看任务：schtasks /Query /TN \"{name}\" /V /FO LIST")
            print("提示：Outlook COM 需要用户已登录，任务类型默认是「只在用户登录时运行」。")
        elif args.action == "uninstall":
            uninstall_task(name)
        elif args.action == "status":
            task_status(name)
        elif args.action == "run-now":
            run_task_now(name)
    except SchedulerError as exc:
        print(f"[计划任务错误] {exc}")
        return 1
    return 0


def cmd_state(args) -> int:
    from config_loader import ConfigError
    try:
        cfg = _load(args)
    except ConfigError as exc:
        print(f"[配置错误] {exc}")
        return 1

    from state_store import StateStore
    if args.action == "reset":
        path = cfg.state.file
        if path.exists():
            path.unlink()
            print(f"已清空状态文件：{path}")
        else:
            print(f"状态文件不存在：{path}")
        return 0

    state = StateStore.load(cfg.state.file, cfg.state.retention_days)
    counts = state.counts()
    print(f"状态文件      : {cfg.state.file}")
    print(f"已处理邮件    : {counts['messages']}")
    print(f"已下载附件    : {counts['attachments']}")
    return 0


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="main.py",
        description="定时下载 Outlook 中指定发件人的邮件附件",
    )
    parser.add_argument("-c", "--config", default=None, help="配置文件路径，默认 ./config.yaml")
    sub = parser.add_subparsers(dest="command", required=True)

    p_run = sub.add_parser("run", help="执行一次下载")
    p_run.add_argument("--dry-run", action="store_true", help="只打印将要下载的内容，不落盘")
    p_run.add_argument("--lookback-days", type=int, default=None, help="临时覆盖配置里的回溯天数")
    p_run.add_argument(
        "--only-today",
        action=argparse.BooleanOptionalAction,
        default=None,
        help="临时覆盖 only_today：--only-today 只看当天，--no-only-today 按回溯天数",
    )
    p_run.set_defaults(func=cmd_run)

    p_check = sub.add_parser("check", help="检查环境与 Outlook 连接")
    p_check.set_defaults(func=cmd_check)

    p_task = sub.add_parser("task", help="管理 Windows 计划任务")
    p_task.add_argument("action", choices=["install", "uninstall", "status", "run-now"])
    p_task.set_defaults(func=cmd_task)

    p_state = sub.add_parser("state", help="查看或清空已处理记录")
    p_state.add_argument("action", choices=["show", "reset"])
    p_state.set_defaults(func=cmd_state)

    return parser


def main(argv=None) -> int:
    force_utf8_console()
    parser = build_parser()
    args = parser.parse_args(argv)
    return args.func(args)


if __name__ == "__main__":
    sys.exit(main())
