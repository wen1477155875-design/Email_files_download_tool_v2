"""端到端冒烟测试：用假的 Outlook 数据验证下载 / 分目录 / 去重 / 记账链路。

不连接真实 Outlook，可随时运行：
    .venv/Scripts/python.exe smoke_test.py
"""

from __future__ import annotations

import io
import shutil
import sys
import tempfile
import zipfile
from datetime import datetime, timedelta, timezone
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

import pipeline  # noqa: E402
from config_loader import load_config  # noqa: E402
from extractor import extract_archive  # noqa: E402
from outlook_client import AttachmentInfo, MessageInfo, OutlookSession  # noqa: E402


def _make_zip(members) -> bytes:
    """members: [(成员名, 内容bytes)]，构造一个真实 zip。"""
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for name, data in members:
            zf.writestr(name, data)
    return buf.getvalue()


class FakeAttachment:
    def __init__(self, index: int, name: str, data: bytes, inline: bool = False):
        self.index = index
        self.name = name
        self.data = data
        self.inline = inline
        self.size = len(data)

    def SaveAsFile(self, path) -> None:  # noqa: N802  # 模仿 COM 的命名
        Path(path).write_bytes(self.data)


class FakeSession(OutlookSession):
    """只覆写需要联网的部分，save_attachment 仍走真实实现（.part + 原子改名）。"""

    mails: list = []

    def __init__(self) -> None:
        pass

    def __enter__(self):
        return self

    def __exit__(self, *exc_info):
        pass

    @classmethod
    def connect(cls, timeout: float = 0):
        return cls()

    def resolve_folder(self, mailbox: str = "", folder_path: str = ""):
        return "FAKE_FOLDER"

    def internet_message_id(self, item) -> str:
        return str(item)

    def iter_messages(self, folder, cutoff_utc, max_items=0, only_mail=True,
                      senders=None, until_utc=None):
        for mail in FakeSession.mails:
            if mail.received_utc < cutoff_utc:
                continue
            if until_utc is not None and mail.received_utc >= until_utc:
                continue
            yield mail

    def iter_attachments(self, message, skip_inline=True, save_embedded=False):
        for att in message.attachments:
            if att.inline and skip_inline:
                continue
            yield AttachmentInfo(index=att.index, name=att.name, size=att.size,
                                 inline=att.inline, embedded=False, obj=att)


def _make_mail(entry_id, sender, subject, received_utc, attachments, body=""):
    mail = MessageInfo(entry_id=entry_id, subject=subject, sender=sender,
                       received_utc=received_utc, item=entry_id, body=body)
    mail.attachments = attachments
    return mail


def main() -> int:
    tmp = Path(tempfile.mkdtemp(prefix="maildl_smoke_"))
    config_path = tmp / "config.yaml"
    # 注意：YAML 双引号里 \ 是转义符，Windows 路径一律用正斜杠写
    downloads = (tmp / "downloads").as_posix()
    extracted = (tmp / "extracted").as_posix()
    state_file = (tmp / "data" / "state.json").as_posix()
    log_dir = (tmp / "data" / "logs").as_posix()
    config_path.write_text(f"""
outlook:
  mailbox: ""
  folder: "收件箱"
  only_mail_items: true
  timeout_seconds: 5
senders:
  - "reports@example.com"
  - "*@partner.com"
scan:
  only_today: false
  lookback_days: 7
  max_messages_per_run: 500
download:
  target_dir: "{downloads}"
  subfolder_template: "{{sender}}/{{date}}"
  skip_inline_images: true
  save_embedded_msg: false
  on_duplicate: rename
  min_size_bytes: 0
  max_size_bytes: 104857600
  allowed_extensions: []
extract:
  enabled: true
  target_dir: "{extracted}"
  subfolder_template: "{{sender}}/{{date}}/{{archive}}"
  extensions: [zip]
  on_duplicate: rename
  delete_archive: false
  max_files: 100
  max_total_bytes: 104857600
  max_ratio: 200
state:
  file: "{state_file}"
  retention_days: 180
logging:
  dir: "{log_dir}"
  level: INFO
schedule:
  task_name: "SmokeTestTask"
  mode: daily
  time: "09:00"
""", encoding="utf-8")

    received = datetime.now(timezone.utc).replace(tzinfo=None) - timedelta(hours=2)
    old_received = datetime.now(timezone.utc).replace(tzinfo=None) - timedelta(days=2)
    FakeSession.mails = [
        _make_mail("E1", "reports@example.com", "月度报表", received, [
            FakeAttachment(1, "报表 2026-09.xlsx", b"x" * 2048),
            FakeAttachment(2, "logo.png", b"y" * 100, inline=True),
        ]),
        _make_mail("E2", "bill@partner.com", "对账单", received, [
            FakeAttachment(1, "对账单.pdf", b"z" * 512),
        ]),
        _make_mail("E3", "spam@other.com", "不该下载", received, [
            FakeAttachment(1, "virus.exe", b"v" * 10),
        ]),
        # 前天的邮件：only_today=true 时应被排除，only_today=false 时应当补下
        _make_mail("E4", "reports@example.com", "前天报表", old_received, [
            FakeAttachment(1, "前天报表.xlsx", b"o" * 1024),
        ]),
        # zip 附件：下载后应自动解压到 extracted 目录
        _make_mail("E5", "reports@example.com", "打包数据", received, [
            FakeAttachment(1, "数据打包.zip", _make_zip([
                ("明细/2026-09.csv", b"a,b,c\n1,2,3\n"),
                ("汇总.txt", "中文内容".encode("gbk")),
            ])),
        ]),
        # zip 里只有一个文档 + 正文带 Report Job Description → 解压后应重命名为描述值
        _make_mail("E6", "reports@example.com", "单文件打包", received, [
            FakeAttachment(1, "单文件.zip", _make_zip([
                ("数据/原件报告.pdf", b"pdf-bytes"),
            ])),
        ], body=("Dear team,\r\nReport Name：AOS月度\r\n"
                 "Report Job Description：AOS数据报告2026.csv\r\n"
                 "Job Submitted：2026-09-09 10:00\r\nRegards")),
    ]

    pipeline.OutlookSession = FakeSession  # type: ignore[assignment]

    try:
        cfg = load_config(config_path)
        print("--- 第 1 次运行（only_today=True，应跳过前天的邮件）---")
        first = pipeline.run_once(cfg, dry_run=False, only_today=True)

        print("\n--- 第 2 次运行（应全部命中去重）---")
        second = pipeline.run_once(cfg, dry_run=False, only_today=True)

        # 模拟文件被人工移走：状态里有记录，但磁盘上没了
        for p in (tmp / "downloads").rglob("对账单.pdf"):
            p.unlink()

        print("\n--- 第 3 次运行（文件丢失应补下 + 补下前天那封）---")
        third = pipeline.run_once(cfg, dry_run=False, only_today=False)

        print("\n--- 解压防护：zip slip 攻击包 ---")
        evil_dir = tmp / "evil"
        evil_zip = evil_dir / "evil.zip"
        evil_zip.parent.mkdir(parents=True, exist_ok=True)
        evil_zip.write_bytes(_make_zip([("../../逃逸.txt", b"pwned")]))
        evil_out = extract_archive(evil_zip, evil_dir / "out")
        escaped = not (evil_dir / "逃逸.txt").exists() and not (tmp / "逃逸.txt").exists()
        print(f"  结果：ok={evil_out['ok']} | {evil_out['message']}")

        print("\n--- 解压：剥离公共顶层目录 ---")
        nested_dir = tmp / "nested"
        nested_dir.mkdir(parents=True, exist_ok=True)
        nested_zip = nested_dir / "pack.zip"
        nested_zip.write_bytes(_make_zip([
            ("TEST/报告.pdf", b"pdf-bytes"),
            ("TEST/sub/数据.csv", b"1,2,3"),
        ]))
        nested_out = extract_archive(nested_zip, nested_dir / "out")
        flat_ok = (
            nested_out["ok"]
            and (nested_dir / "out" / "报告.pdf").exists()
            and (nested_dir / "out" / "sub" / "数据.csv").exists()
            and not (nested_dir / "out" / "TEST").exists()
        )
        print(f"  结果：ok={nested_out['ok']} | {nested_out['message']}")

        # E6：单文件 zip + 正文 Job Description（带 .csv 扩展名）→ 应剥离扩展名重命名
        renamed = any(p.name == "AOS数据报告2026.pdf"
                      for p in (tmp / "extracted").rglob("*"))
        original_left = any(p.name == "原件报告.pdf"
                            for p in (tmp / "extracted").rglob("*"))
        # 三行正文格式下精准提取（不受 Report Name / Job Submitted 干扰）
        desc3 = pipeline.job_description(
            "Report Name：甲\r\nReport Job Description：bbbb.csv\r\nJob Submitted：乙")
        desc3_ok = desc3 == "bbbb.csv"
        # E6：压缩包本身也按同一规则重命名（bbbb.csv -> bbbb.zip）
        archive_renamed = any(p.name == "AOS数据报告2026.zip"
                              for p in (tmp / "downloads").rglob("*"))
        archive_original_left = any(p.name == "单文件.zip"
                                    for p in (tmp / "downloads").rglob("*"))

        # 重复下载/解压同一份 zip 时，命名不能退化成 bbbb_1（必须覆盖而不是加后缀）
        class _NullLogger:
            def info(self, *a, **k): pass
            def warning(self, *a, **k): pass
            def error(self, *a, **k): pass

        dup_dir = Path(tempfile.mkdtemp(prefix="maildl_dup_"))
        out_dir = dup_dir / "out"

        def _pass(content: bytes) -> list:
            """模拟一次"下载 zip -> 解压 -> 按正文命名"，返回目录里的文件名。"""
            zp = dup_dir / "pack.zip"
            with zipfile.ZipFile(zp, "w") as zf:
                zf.writestr("原件.pdf", content)
            r = extract_archive(zp, out_dir)
            pipeline._rename_by_job_desc(r.get("paths") or [], "bbbb.csv", _NullLogger())
            return sorted(p.name for p in out_dir.rglob("*"))

        dup_names = _pass(b"x")           # 第一次
        dup_names = _pass(b"x")           # 第二次：内容相同 -> 覆盖，不留 bbbb_1
        dup_ok = dup_names == ["bbbb.pdf"]

        out_dir2 = dup_dir / "out2"       # 内容不同的同名报告 -> 必须保留两份
        def _pass2(content: bytes) -> list:
            zp = dup_dir / "pack2.zip"
            with zipfile.ZipFile(zp, "w") as zf:
                zf.writestr("原件.pdf", content)
            r = extract_archive(zp, out_dir2)
            pipeline._rename_by_job_desc(r.get("paths") or [], "bbbb.csv", _NullLogger())
            return sorted(p.name for p in out_dir2.rglob("*"))

        diff_names = _pass2(b"AAA")
        diff_names = _pass2(b"BBB")
        diff_ok = diff_names == ["bbbb.pdf", "bbbb_1.pdf"]
        print("  相同内容两次:", dup_names, "| 不同内容两次:", diff_names)

    finally:
        shutil.rmtree(tmp, ignore_errors=True)

    checks = [
        ("下载 4 个附件（跳过内嵌图 + 非目标发件人 + 非当天邮件）", first["downloaded"] == 4),
        ("没有失败项", first["failed"] == 0),
        ("zip 已解压 2 个", first["extracted"] == 2),
        ("解压出 3 个文件（含子目录内的）", first["extract_files"] == 3),
        ("第 2 次运行 0 新增下载", second["downloaded"] == 0),
        ("第 2 次运行 4 个命中去重", second["skipped_already"] == 4),
        ("关闭 only_today 后补下前天邮件 + 丢失文件重下", third["downloaded"] == 2),
        ("补下时未丢的附件不重复下载", third["skipped_already"] == 3),
        ("zip slip 攻击包被拦截", evil_out["ok"] is False),
        ("zip slip 未写出任何越界文件", escaped),
        ("zip 公共顶层目录被剥离，文件直接落在解压根目录", flat_ok),
        ("单文件 zip 已按正文 Job Description 重命名（扩展名 .csv 被剥离）", renamed),
        ("重命名后原文件名不再存在", not original_left),
        ("三行正文格式精准提取 Job Description", desc3_ok),
        ("压缩包已按同一规则重命名（单文件.zip -> AOS数据报告2026.zip）", archive_renamed),
        ("压缩包原文件名不再存在", not archive_original_left),
        ("重复解压同一 zip 命名不退化成 bbbb_1", dup_ok),
        ("内容不同的同名报告保留两份，不静默覆盖", diff_ok),
    ]

    print("\n=== 断言结果 ===")
    ok = True
    for label, passed in checks:
        print(f"{'PASS' if passed else 'FAIL'}  {label}")
        ok = ok and passed
    print("\n总体：", "PASS" if ok else "FAIL")
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
