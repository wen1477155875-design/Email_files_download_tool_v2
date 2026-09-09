"""邮件附件自动下载工具 - 图形界面（tkinter，Python 自带，无新增依赖）。

用法：
    .venv\\Scripts\\python.exe ui.py
    .venv\\Scripts\\pythonw.exe ui.py   （不弹控制台窗口）

说明：
- 「启动定时」= 把当前界面配置保存进 config.yaml，并注册 Windows 计划任务；
  「停止定时」= 删除该计划任务。界面关掉后定时仍然生效（任务归 Windows 管）。
- 「立即运行」= 马上跑一次主流程，可勾选 dry-run 或指定补抓某一天的邮件。
"""

from __future__ import annotations

import logging
import queue
import re
import sys
import threading
import traceback
from pathlib import Path
from tkinter import filedialog, messagebox, scrolledtext, ttk
import tkinter as tk

import yaml

BASE_DIR = Path(__file__).resolve().parent
CONFIG_PATH = BASE_DIR / "config.yaml"
TIME_RE = re.compile(r"^([01]?\d|2[0-3]):[0-5]\d$")
DATE_RE = re.compile(r"^\d{4}-\d{2}-\d{2}$")


# --------------------------------------------------------------------------
# 配置读写：写回时做"段落内精准替换"，保留 config.yaml 里的全部注释
# --------------------------------------------------------------------------

def _update_section(text: str, section: str, updates: dict) -> str:
    """在指定顶层段落里替换若干键值，其他行（含注释）原样保留。"""
    pattern = re.compile(
        rf"(?ms)^({section}):[^\n]*\n(.*?)(?=^[A-Za-z_][A-Za-z0-9_]*:|\Z)"
    )
    match = pattern.search(text)
    if not match:
        raise ValueError(f"配置文件里找不到段落 {section}:")
    body = match.group(2)
    for key, value in updates.items():
        key_re = re.compile(rf"(?m)^(\s+{re.escape(key)}:)[^\n]*$")
        if not key_re.search(body):
            raise ValueError(f"段落 {section}: 里找不到键 {key}")
        body = key_re.sub(lambda m: f"{m.group(1)} {value}", body, count=1)
    return text[: match.start(2)] + body + text[match.end(2):]


def _replace_senders(text: str, senders: list) -> str:
    block = "senders:\n" + "".join(f'  - "{s}"\n' for s in senders)
    new_text, n = re.subn(r"(?ms)^senders:\n(?:[ \t]+-[^\n]*\n)+", block, text, count=1)
    if n == 0:
        raise ValueError("找不到 senders: 段落")
    return new_text


def _quote(value) -> str:
    return '"' + str(value).strip().replace('"', "").replace("\\", "/") + '"'


def load_values(path: Path = CONFIG_PATH) -> dict:
    raw = yaml.safe_load(path.read_text(encoding="utf-8")) or {}
    outlook = raw.get("outlook") or {}
    scan = raw.get("scan") or {}
    download = raw.get("download") or {}
    extract = raw.get("extract") or {}
    schedule = raw.get("schedule") or {}
    return {
        "mailbox": str(outlook.get("mailbox", "") or ""),
        "folder": str(outlook.get("folder", "收件箱") or "收件箱"),
        "senders": [str(s).strip() for s in (raw.get("senders") or []) if str(s).strip()],
        "download_dir": str(download.get("target_dir", "")),
        "extract_enabled": bool(extract.get("enabled", False)),
        "extract_dir": str(extract.get("target_dir", "")),
        "scan_mode": "today" if scan.get("only_today") else "days",
        "lookback_days": int(scan.get("lookback_days", 7)),
        "daily_time": str(schedule.get("time", "09:00")),
    }


def save_values(values: dict, path: Path = CONFIG_PATH) -> None:
    """把界面值写回 config.yaml（保留注释）。"""
    text = path.read_text(encoding="utf-8")
    text = _update_section(text, "outlook", {
        "mailbox": _quote(values["mailbox"]),
        "folder": _quote(values["folder"]),
    })
    text = _replace_senders(text, values["senders"])
    text = _update_section(text, "scan", {
        "only_today": "true" if values["scan_mode"] == "today" else "false",
        "lookback_days": str(int(values["lookback_days"])),
    })
    text = _update_section(text, "download", {"target_dir": _quote(values["download_dir"])})
    text = _update_section(text, "extract", {
        "enabled": "true" if values["extract_enabled"] else "false",
        "target_dir": _quote(values["extract_dir"]),
    })
    text = _update_section(text, "schedule", {
        "mode": "daily",
        "time": _quote(values["daily_time"]),
    })
    path.write_text(text, encoding="utf-8")


def validate(values: dict) -> str:
    """校验界面输入，返回错误信息；None 表示通过。"""
    if not values["senders"]:
        return "发件人至少填一个"
    if not str(values["download_dir"]).strip():
        return "下载目录不能为空"
    if values["extract_enabled"] and not str(values["extract_dir"]).strip():
        return "已勾选解压，但解压目录为空"
    if values["scan_mode"] == "days" and not 1 <= int(values["lookback_days"]) <= 365:
        return "回溯天数应在 1~365 之间"
    if not TIME_RE.match(str(values["daily_time"]).strip()):
        return "每日运行时间格式应为 HH:MM，例如 10:00"
    return None


# --------------------------------------------------------------------------
# 日志桥：工作线程 -> 队列 -> UI 线程
# --------------------------------------------------------------------------

class QueueLogHandler(logging.Handler):
    def __init__(self, q: queue.Queue):
        super().__init__()
        self.q = q
        self.setFormatter(logging.Formatter("%(asctime)s [%(levelname)s] %(message)s",
                                            "%H:%M:%S"))

    def emit(self, record):
        try:
            self.q.put(self.format(record))
        except Exception:
            pass


# --------------------------------------------------------------------------
# 界面
# --------------------------------------------------------------------------

class App:
    def __init__(self, root: tk.Tk):
        self.root = root
        root.title("邮件附件自动下载工具")
        root.minsize(720, 640)

        self.log_q: queue.Queue = queue.Queue()
        self.busy = False

        self._build_vars()
        self._build_ui()
        self.load_into_ui()
        self.root.after(150, self._poll_log)
        self.refresh_status()

    # ---------- 控件 ----------

    def _build_vars(self):
        self.var_mailbox = tk.StringVar()
        self.var_folder = tk.StringVar(value="收件箱")
        self.var_download = tk.StringVar()
        self.var_extract_on = tk.BooleanVar(value=True)
        self.var_extract_dir = tk.StringVar()
        self.var_scan_mode = tk.StringVar(value="days")
        self.var_lookback = tk.IntVar(value=7)
        self.var_daily_time = tk.StringVar(value="09:00")
        self.var_dry_run = tk.BooleanVar(value=False)
        self.var_date = tk.StringVar()
        self.var_status = tk.StringVar(value="定时任务：查询中…")

    def _build_ui(self):
        pad = {"padx": 8, "pady": 4}
        outer = ttk.Frame(self.root)
        outer.pack(fill="both", expand=True, padx=10, pady=8)

        # -- 邮箱 --
        f1 = ttk.LabelFrame(outer, text="邮箱")
        f1.pack(fill="x", **pad)
        ttk.Label(f1, text="邮箱 (mailbox)：").grid(row=0, column=0, sticky="e", **pad)
        ttk.Entry(f1, textvariable=self.var_mailbox, width=34).grid(row=0, column=1, sticky="w", **pad)
        ttk.Label(f1, text="文件夹：").grid(row=0, column=2, sticky="e", **pad)
        ttk.Entry(f1, textvariable=self.var_folder, width=16).grid(row=0, column=3, sticky="w", **pad)
        ttk.Label(f1, text="留空 = 默认邮箱").grid(row=0, column=4, sticky="w")

        # -- 发件人 --
        f2 = ttk.LabelFrame(outer, text="关注发件人（每行一个，支持 *@域名.com 通配）")
        f2.pack(fill="x", **pad)
        self.txt_senders = scrolledtext.ScrolledText(f2, height=3, width=70, font=("Consolas", 10))
        self.txt_senders.pack(fill="x", **pad)

        # -- 路径 --
        f3 = ttk.LabelFrame(outer, text="下载与解压")
        f3.pack(fill="x", **pad)
        ttk.Label(f3, text="下载目录：").grid(row=0, column=0, sticky="e", **pad)
        ttk.Entry(f3, textvariable=self.var_download, width=52).grid(row=0, column=1, sticky="w", **pad)
        ttk.Button(f3, text="浏览…", command=lambda: self._browse(self.var_download)).grid(row=0, column=2, **pad)
        ttk.Checkbutton(f3, text="自动解压压缩包", variable=self.var_extract_on).grid(row=1, column=0, sticky="e", **pad)
        ttk.Entry(f3, textvariable=self.var_extract_dir, width=52).grid(row=1, column=1, sticky="w", **pad)
        ttk.Button(f3, text="浏览…", command=lambda: self._browse(self.var_extract_dir)).grid(row=1, column=2, **pad)

        # -- 扫描范围 --
        f4 = ttk.LabelFrame(outer, text="扫描范围（定时运行时生效）")
        f4.pack(fill="x", **pad)
        ttk.Radiobutton(f4, text="仅今天", value="today", variable=self.var_scan_mode).grid(row=0, column=0, **pad)
        ttk.Radiobutton(f4, text="最近", value="days", variable=self.var_scan_mode).grid(row=0, column=1, sticky="e")
        ttk.Spinbox(f4, from_=1, to=365, textvariable=self.var_lookback, width=5).grid(row=0, column=2)
        ttk.Label(f4, text="天").grid(row=0, column=3, sticky="w")

        # -- 定时 --
        f5 = ttk.LabelFrame(outer, text="定时方式（每天固定时间运行）")
        f5.pack(fill="x", **pad)
        ttk.Label(f5, text="每天").grid(row=0, column=0, **pad)
        ttk.Entry(f5, textvariable=self.var_daily_time, width=8).grid(row=0, column=1, sticky="w", **pad)
        ttk.Label(f5, text="运行 (HH:MM)").grid(row=0, column=2, sticky="w")

        # -- 操作 --
        f6 = ttk.LabelFrame(outer, text="控制")
        f6.pack(fill="x", **pad)
        self.btn_save = ttk.Button(f6, text="保存配置", command=self.on_save)
        self.btn_start = ttk.Button(f6, text="启动定时", command=self.on_start)
        self.btn_stop = ttk.Button(f6, text="停止定时", command=self.on_stop)
        self.btn_refresh = ttk.Button(f6, text="刷新状态", command=self.refresh_status)
        self.btn_save.grid(row=0, column=0, **pad)
        self.btn_start.grid(row=0, column=1, **pad)
        self.btn_stop.grid(row=0, column=2, **pad)
        self.btn_refresh.grid(row=0, column=3, **pad)
        ttk.Label(f6, textvariable=self.var_status).grid(row=0, column=4, sticky="w", padx=12)

        # -- 立即运行 --
        f7 = ttk.LabelFrame(outer, text="立即运行一次")
        f7.pack(fill="x", **pad)
        ttk.Label(f7, text="补抓指定日期 (YYYY-MM-DD，可空)：").grid(row=0, column=0, sticky="e", **pad)
        ttk.Entry(f7, textvariable=self.var_date, width=12).grid(row=0, column=1, sticky="w", **pad)
        ttk.Checkbutton(f7, text="只预览不下载 (dry-run)", variable=self.var_dry_run).grid(row=0, column=2, **pad)
        self.btn_run = ttk.Button(f7, text="立即运行", command=self.on_run_now)
        self.btn_run.grid(row=0, column=3, **pad)

        # -- 日志 --
        f8 = ttk.LabelFrame(outer, text="运行日志")
        f8.pack(fill="both", expand=True, **pad)
        self.txt_log = scrolledtext.ScrolledText(f8, height=12, font=("Consolas", 9), state="disabled")
        self.txt_log.pack(fill="both", expand=True, padx=4, pady=4)

    # ---------- 工具 ----------

    def _browse(self, var: tk.StringVar):
        chosen = filedialog.askdirectory(initialdir=var.get() or str(BASE_DIR))
        if chosen:
            var.set(chosen.replace("\\", "/"))

    def _collect(self) -> dict:
        senders = [ln.strip() for ln in self.txt_senders.get("1.0", "end").splitlines() if ln.strip()]
        return {
            "mailbox": self.var_mailbox.get().strip(),
            "folder": self.var_folder.get().strip() or "收件箱",
            "senders": senders,
            "download_dir": self.var_download.get().strip(),
            "extract_enabled": bool(self.var_extract_on.get()),
            "extract_dir": self.var_extract_dir.get().strip(),
            "scan_mode": self.var_scan_mode.get(),
            "lookback_days": int(self.var_lookback.get() or 7),
            "daily_time": self.var_daily_time.get().strip(),
        }

    def load_into_ui(self):
        v = load_values()
        self.var_mailbox.set(v["mailbox"])
        self.var_folder.set(v["folder"])
        self.txt_senders.delete("1.0", "end")
        self.txt_senders.insert("1.0", "\n".join(v["senders"]))
        self.var_download.set(v["download_dir"])
        self.var_extract_on.set(v["extract_enabled"])
        self.var_extract_dir.set(v["extract_dir"])
        self.var_scan_mode.set(v["scan_mode"])
        self.var_lookback.set(v["lookback_days"])
        self.var_daily_time.set(v["daily_time"])

    def _log(self, message: str):
        self.log_q.put(message)

    def _poll_log(self):
        try:
            while True:
                line = self.log_q.get_nowait()
                self.txt_log.configure(state="normal")
                self.txt_log.insert("end", line.rstrip("\n") + "\n")
                self.txt_log.see("end")
                self.txt_log.configure(state="disabled")
        except queue.Empty:
            pass
        self.root.after(150, self._poll_log)

    def _set_busy(self, busy: bool):
        self.busy = busy
        state = "disabled" if busy else "normal"
        for btn in (self.btn_save, self.btn_start, self.btn_stop,
                    self.btn_refresh, self.btn_run):
            btn.configure(state=state)

    def _spawn(self, work, on_done=None):
        """在工作线程里跑 work()，完成后回调 on_done(result, error)。"""

        def runner():
            result, error = None, None
            try:
                self._com_init()
                result = work()
            except Exception as exc:  # noqa: BLE001
                error = exc
                traceback.print_exc()
            finally:
                self._com_uninit()
                self.root.after(0, lambda: self._finish(on_done, result, error))

        threading.Thread(target=runner, daemon=True).start()

    @staticmethod
    def _com_init():
        try:
            import pythoncom
            pythoncom.CoInitialize()
        except Exception:
            pass

    @staticmethod
    def _com_uninit():
        try:
            import pythoncom
            pythoncom.CoUninitialize()
        except Exception:
            pass

    def _finish(self, on_done, result, error):
        self._set_busy(False)
        if error is not None:
            self._log(f"[错误] {error}")
            messagebox.showerror("出错了", str(error))
        elif on_done:
            on_done(result)

    def _saved_or_raise(self) -> bool:
        values = self._collect()
        problem = validate(values)
        if problem:
            messagebox.showwarning("填写不完整", problem)
            return False
        try:
            save_values(values)
        except Exception as exc:
            messagebox.showerror("保存失败", f"写 config.yaml 失败：\n{exc}")
            return False
        self._log("配置已保存到 config.yaml")
        return True

    # ---------- 按钮 ----------

    def on_save(self):
        self._saved_or_raise()

    def on_start(self):
        if not self._saved_or_raise():
            return
        self._set_busy(True)
        self._log("正在注册 Windows 计划任务…")

        def work():
            from config_loader import load_config
            from scheduler import install_task
            cfg = load_config(CONFIG_PATH)
            logger = self._make_logger()
            install_task(cfg, logger=logger)
            return cfg.schedule.task_name

        def done(task_name):
            self._log(f"定时任务已启动：{task_name}")
            self.refresh_status()

        self._spawn(work, done)

    def on_stop(self):
        self._set_busy(True)
        self._log("正在停止定时任务…")

        def work():
            from scheduler import uninstall_task
            from config_loader import load_config
            cfg = load_config(CONFIG_PATH)
            uninstall_task(cfg.schedule.task_name, logger=self._make_logger())
            return cfg.schedule.task_name

        def done(task_name):
            self._log(f"定时任务已停止：{task_name}")
            self.refresh_status()

        self._spawn(work, done)

    def on_run_now(self):
        if not self._saved_or_raise():
            return
        date_text = self.var_date.get().strip()
        if date_text and not DATE_RE.match(date_text):
            messagebox.showwarning("日期格式", "指定日期应为 YYYY-MM-DD，例如 2026-09-01")
            return
        if date_text:
            try:
                y, m, d = (int(x) for x in date_text.split("-"))
                import datetime as _dt
                _dt.date(y, m, d)
            except ValueError:
                messagebox.showwarning("日期无效", f"{date_text} 不是有效日期")
                return
        self._set_busy(True)
        dry = bool(self.var_dry_run.get())
        # 直接采用界面上的扫描范围，不依赖 config 是否同步（填了指定日期时以日期为准）
        scan_today = self.var_scan_mode.get() == "today"
        lookback = int(self.var_lookback.get() or 7)
        self._log(f"开始运行（dry_run={dry}，扫描范围="
                  + (f"指定日期 {date_text}" if date_text else ("仅今天" if scan_today else f"最近 {lookback} 天"))
                  + "）…")

        def work():
            from config_loader import load_config
            from pipeline import run_once
            from utils import setup_logging
            cfg = load_config(CONFIG_PATH)
            logger = setup_logging(cfg.logging.dir, cfg.logging.level)
            handler = QueueLogHandler(self.log_q)
            logger.addHandler(handler)
            try:
                return run_once(cfg, dry_run=dry, logger=logger,
                                specific_date=date_text or None,
                                only_today=None if date_text else scan_today,
                                lookback_days=lookback)
            finally:
                logger.removeHandler(handler)

        def done(stats):
            self._log(f"运行结束：下载 {stats.get('downloaded', 0)} 个附件，"
                      f"解压 {stats.get('extracted', 0)} 个压缩包")

        self._spawn(work, done)

    def refresh_status(self):
        def work():
            from scheduler import query_status
            from config_loader import load_config
            cfg = load_config(CONFIG_PATH)
            return query_status(cfg.schedule.task_name)

        def done(info):
            if info:
                self.var_status.set("定时任务：已启动")
                self._log("任务状态：\n" + info)
            else:
                self.var_status.set("定时任务：未启动")

        self._spawn(work, done)

    def _make_logger(self) -> logging.Logger:
        logger = logging.getLogger("maildl.ui")
        logger.setLevel(logging.INFO)
        logger.propagate = False
        for h in list(logger.handlers):
            logger.removeHandler(h)
        logger.addHandler(QueueLogHandler(self.log_q))
        return logger


def main() -> int:
    try:
        from utils import force_utf8_console
        force_utf8_console()
    except Exception:
        pass
    root = tk.Tk()
    try:
        ttk.Style().theme_use("vista")
    except Exception:
        pass
    App(root)
    root.mainloop()
    return 0


if __name__ == "__main__":
    sys.exit(main())
