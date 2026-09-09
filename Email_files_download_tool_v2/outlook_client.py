"""Outlook COM 封装层。

注意：Outlook 自动化必须在已登录的 Windows 交互会话中运行，
不能放在"不管用户是否登录都要运行"的计划任务里。
"""

from __future__ import annotations

import os
import re
import threading
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Iterator, List, Optional

from utils import sender_matches, to_utc_naive

_IMPORT_ERROR: Optional[BaseException] = None
try:
    import pythoncom  # type: ignore
    import win32com.client  # type: ignore
except Exception as exc:  # pragma: no cover
    _IMPORT_ERROR = exc


OL_FOLDER_INBOX = 6
OL_MAIL_ITEM = 43
OL_ATTACH_BY_VALUE = 1
OL_ATTACH_EMBEDDED_ITEM = 5

PR_SENDER_SMTP_ADDRESS = "http://schemas.microsoft.com/mapi/proptag/0x5F01001F"
PR_SENT_REPRESENTING_SMTP_ADDRESS = "http://schemas.microsoft.com/mapi/proptag/0x5D01001F"
PR_ATTACH_CONTENT_ID = "http://schemas.microsoft.com/mapi/proptag/0x3712001F"
PR_ATTACH_SIZE = "http://schemas.microsoft.com/mapi/proptag/0x0E200003"
PR_INTERNET_MESSAGE_ID = "http://schemas.microsoft.com/mapi/proptag/0x1035001F"

INBOX_NAMES = ("收件箱", "Inbox")


class OutlookError(RuntimeError):
    pass


class OutlookTimeout(OutlookError):
    pass


@dataclass
class MessageInfo:
    entry_id: str
    subject: str
    sender: str
    received_utc: datetime
    item: object = field(repr=False, default=None)
    body: str = ""


@dataclass
class AttachmentInfo:
    index: int
    name: str
    size: int
    inline: bool
    embedded: bool
    obj: object = field(repr=False, default=None)


def _prop(target, tag: str):
    """读取 MAPI 属性，失败返回 None。"""
    try:
        return target.PropertyAccessor.GetProperty(tag)
    except Exception:
        return None


_PROBE_CODE = r"""
import sys
try:
    import pythoncom, win32com.client
    pythoncom.CoInitialize()
    app = win32com.client.Dispatch("Outlook.Application")
    _ = app.Session.Folders.Count
    print("OK")
    sys.exit(0)
except Exception as exc:
    sys.stderr.write("ERR: %s\n" % exc)
    sys.exit(1)
"""


def _probe_outlook(timeout: float) -> None:
    """在子进程里探测 Outlook 能否连接。子进程可以被 kill，所以主线程不会卡死。"""
    import subprocess
    import sys

    try:
        proc = subprocess.run(
            [sys.executable, "-c", _PROBE_CODE],
            capture_output=True,
            timeout=max(5.0, float(timeout)),
            text=True,
            encoding="utf-8",
            errors="replace",
        )
    except subprocess.TimeoutExpired:
        raise OutlookTimeout(
            f"连接 Outlook 超过 {timeout:.0f} 秒仍未响应。"
            "常见原因：Outlook 首次启动正在弹配置向导、正在要求输入密码、"
            "或装的是不支持 COM 的「Outlook (new)」。请手动打开一次 Outlook 并确认能正常收信。"
        )
    except Exception as exc:  # noqa: BLE001
        raise OutlookError(f"无法启动 Outlook 探测进程：{exc}") from exc

    if proc.returncode == 0:
        return
    detail = (proc.stderr or proc.stdout or "").strip()
    raise OutlookError(
        "连接 Outlook 失败。"
        f"子进程输出：{detail or '(无)'}\n"
        "请确认：安装的是经典桌面版 Outlook（不是 Windows 11 自带的「Outlook (new)」）、"
        "已配置邮箱账号、且当前处于已登录的 Windows 会话。"
    )


class OutlookSession:
    def __init__(self) -> None:
        if _IMPORT_ERROR is not None:
            raise OutlookError(
                f"缺少 pywin32：{_IMPORT_ERROR}。请先执行 pip install -r requirements.txt"
            )
        self._coinitialized = False
        try:
            pythoncom.CoInitialize()
            self._coinitialized = True
        except Exception:
            self._coinitialized = False
        try:
            self.app = win32com.client.Dispatch("Outlook.Application")
            try:
                # Session 比 GetNamespace("MAPI") 更直接，优先使用
                self.namespace = self.app.Session
            except Exception:
                self.namespace = self.app.GetNamespace("MAPI")
        except Exception as exc:
            raise OutlookError(
                f"无法连接 Outlook：{exc}。"
                "请确认已安装桌面版 Outlook（注意：不是 Windows 11 自带的「Outlook (new)」）、"
                "已配置账号，且当前处于已登录的 Windows 会话。"
            ) from exc

    @classmethod
    def connect(cls, timeout: float = 120.0) -> "OutlookSession":
        """带超时的连接。

        COM 的 Dispatch 是同步阻塞的：Outlook 弹出配置文件向导、要求输入密码或卡死时，
        计划任务里的进程会永远挂着，后续运行全被运行锁挡住。

        实现：先用**子进程**做一次探针（子进程可以被 kill，因此不会卡住主线程），
        探针通过后再在主线程建立正式的 COM 连接。

        ⚠️ 不能把 COM 对象的创建放进线程、再拿回主线程使用 —— COM 是套间线程模型，
        跨线程使用接口指针会失败，表现为 Dispatch/GetNamespace 成功、
        但后续调用（如 namespace.Folders）报莫名其妙的错。
        """
        _probe_outlook(timeout)
        return cls()

    def __enter__(self) -> "OutlookSession":
        return self

    def __exit__(self, *exc_info) -> None:
        self.close()

    def close(self) -> None:
        try:
            self.namespace = None
            self.app = None
        finally:
            if self._coinitialized:
                try:
                    pythoncom.CoUninitialize()
                except Exception:
                    pass
                self._coinitialized = False

    def list_mailboxes(self) -> List[str]:
        """优先返回账号的 SMTP 地址（填进 config 的就是这个），取不到才退回文件夹显示名。"""
        names: List[str] = []
        try:
            accounts = self.namespace.Accounts
            for i in range(1, accounts.Count + 1):
                try:
                    smtp = str(accounts.Item(i).SmtpAddress or "").strip()
                except Exception:
                    continue
                if smtp:
                    names.append(smtp)
        except Exception:
            pass
        if names:
            return names

        try:
            folders = self.namespace.Folders
            for i in range(1, folders.Count + 1):
                try:
                    names.append(str(folders.Item(i).Name))
                except Exception:
                    continue
        except Exception:
            pass
        return names

    def _find_store_root(self, mailbox: str):
        if not mailbox:
            return None
        target = mailbox.strip().lower()

        # 1) 最可靠：遍历账号，按 SMTP 地址匹配，再取其投递存储的根目录
        try:
            accounts = self.namespace.Accounts
            for i in range(1, accounts.Count + 1):
                try:
                    account = accounts.Item(i)
                    smtp = str(account.SmtpAddress or "").strip().lower()
                except Exception:
                    continue
                if smtp == target:
                    try:
                        return account.DeliveryStore.GetRootFolder()
                    except Exception:
                        try:
                            return account.Session.GetDefaultFolder(OL_FOLDER_INBOX)
                        except Exception:
                            pass
        except Exception:
            pass

        # 2) 回退：按顶层文件夹显示名匹配（精确 -> 包含）
        folders = self.namespace.Folders
        fallback = None
        for i in range(1, folders.Count + 1):
            folder = folders.Item(i)
            try:
                name = str(folder.Name).strip().lower()
            except Exception:
                continue
            if name == target:
                return folder
            if target in name and fallback is None:
                fallback = folder
        return fallback

    def _default_inbox(self, root=None):
        if root is None:
            return self.namespace.GetDefaultFolder(OL_FOLDER_INBOX)
        try:
            return root.GetDefaultFolder(OL_FOLDER_INBOX)
        except Exception:
            pass
        for name in INBOX_NAMES:
            try:
                return root.Folders(name)
            except Exception:
                continue
        raise OutlookError("无法在指定邮箱中定位收件箱，请在配置里用 folder 显式写路径")

    def resolve_folder(self, mailbox: str = "", folder_path: str = "收件箱"):
        root = None
        if mailbox:
            root = self._find_store_root(mailbox)
            if root is None:
                raise OutlookError(f"找不到名为 {mailbox!r} 的邮箱。当前可用：{self.list_mailboxes()}")

        parts = [p.strip() for p in re.split(r"[\\/]+", folder_path or "") if p.strip()]
        if not parts:
            return self._default_inbox(root)

        if parts[0].lower() in ("收件箱", "inbox"):
            folder = self._default_inbox(root)
            parts = parts[1:]
        else:
            folder = root if root is not None else self.namespace

        for part in parts:
            try:
                folder = folder.Folders(part)
            except Exception as exc:
                raise OutlookError(f"找不到子文件夹 {part!r}（配置路径：{folder_path}）") from exc
        return folder

    def sender_smtp(self, item) -> str:
        """Exchange 账号下 SenderEmailAddress 可能是 X500 地址，需要解析成 SMTP。"""
        try:
            mail_type = item.SenderEmailType
        except Exception:
            mail_type = None

        if mail_type == "EX":
            try:
                user = item.Sender.GetExchangeUser()
                if user and user.PrimarySmtpAddress:
                    return str(user.PrimarySmtpAddress)
            except Exception:
                pass
            try:
                dist = item.Sender.GetExchangeDistributionList()
                if dist and dist.PrimarySmtpAddress:
                    return str(dist.PrimarySmtpAddress)
            except Exception:
                pass
        else:
            try:
                addr = item.SenderEmailAddress
                if addr and "@" in addr:
                    return str(addr)
            except Exception:
                pass

        for tag in (PR_SENDER_SMTP_ADDRESS, PR_SENT_REPRESENTING_SMTP_ADDRESS):
            addr = _prop(item, tag)
            if addr and "@" in str(addr):
                return str(addr)
        try:
            return str(item.SenderEmailAddress or "")
        except Exception:
            return ""

    def internet_message_id(self, item) -> str:
        value = _prop(item, PR_INTERNET_MESSAGE_ID)
        return str(value).strip() if value else ""

    def iter_messages(
        self,
        folder,
        cutoff_utc: datetime,
        max_items: int = 500,
        only_mail: bool = True,
        senders: Optional[List[str]] = None,
        until_utc: Optional[datetime] = None,
    ) -> Iterator[MessageInfo]:
        """按接收时间倒序遍历，早于截止时间的直接停止（外层无需全量扫描）。

        cutoff_utc：下界，早于它的邮件终止遍历。
        until_utc ：上界，晚于它的邮件跳过（用于"只看今天"，避免时钟偏差
                    导致的未来邮件被误收）。因为是倒序，这里用 continue 而非 return。
        """
        try:
            items = folder.Items
            items.Sort("[ReceivedTime]", True)
            total = int(items.Count)
        except Exception as exc:
            raise OutlookError(f"读取文件夹内容失败：{exc}") from exc

        limit = min(total, max_items) if max_items and max_items > 0 else total
        for i in range(1, limit + 1):
            try:
                item = items.Item(i)
            except Exception:
                continue
            try:
                if only_mail and int(item.Class) != OL_MAIL_ITEM:
                    continue
            except Exception:
                continue

            try:
                received = to_utc_naive(item.ReceivedTime)
            except Exception:
                continue
            if received is None:
                continue
            if received < cutoff_utc:
                return
            if until_utc is not None and received >= until_utc:
                continue

            sender = self.sender_smtp(item)
            if senders and not sender_matches(sender, senders):
                continue

            try:
                subject = str(item.Subject or "")
            except Exception:
                subject = ""
            try:
                entry_id = str(item.EntryID)
            except Exception:
                continue

            # 正文只在"通过发件人筛选"后才取，减少大邮箱下的 COM 调用开销
            try:
                body = str(item.Body or "")
            except Exception:
                body = ""

            yield MessageInfo(entry_id=entry_id, subject=subject, sender=sender,
                              received_utc=received, item=item, body=body)

    def iter_attachments(
        self,
        message: MessageInfo,
        skip_inline: bool = True,
        save_embedded: bool = False,
    ) -> Iterator[AttachmentInfo]:
        mail = message.item
        try:
            html_body = str(mail.HTMLBody or "").lower()
        except Exception:
            html_body = ""
        try:
            total = int(mail.Attachments.Count)
        except Exception:
            return

        for i in range(1, total + 1):
            try:
                att = mail.Attachments.Item(i)
            except Exception:
                continue
            try:
                att_type = int(att.Type)
            except Exception:
                att_type = OL_ATTACH_BY_VALUE

            embedded = att_type == OL_ATTACH_EMBEDDED_ITEM
            if embedded and not save_embedded:
                continue
            if not embedded and att_type != OL_ATTACH_BY_VALUE:
                continue

            try:
                name = str(att.FileName or att.DisplayName or f"attachment_{i}")
            except Exception:
                name = f"attachment_{i}"
            if embedded and not name.lower().endswith(".msg"):
                name += ".msg"

            try:
                size = int(att.Size)
            except Exception:
                try:
                    size = int(_prop(att, PR_ATTACH_SIZE) or 0)
                except Exception:
                    size = 0

            content_id = str(_prop(att, PR_ATTACH_CONTENT_ID) or "")
            inline = bool(content_id) and f"cid:{content_id}".lower() in html_body
            if skip_inline and inline:
                continue

            yield AttachmentInfo(index=i, name=name, size=size, inline=inline,
                                 embedded=embedded, obj=att)

    @staticmethod
    def save_attachment(attachment: AttachmentInfo, final_path: Path) -> int:
        """先写 .part 再原子改名，避免中断留下半截文件。"""
        final_path.parent.mkdir(parents=True, exist_ok=True)
        tmp_path = final_path.with_name(final_path.name + ".part")
        if tmp_path.exists():
            tmp_path.unlink()
        attachment.obj.SaveAsFile(str(tmp_path))
        os.replace(str(tmp_path), str(final_path))
        try:
            return final_path.stat().st_size
        except Exception:
            return attachment.size
