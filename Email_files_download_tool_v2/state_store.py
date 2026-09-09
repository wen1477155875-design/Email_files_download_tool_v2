"""已处理记录：增量去重的核心。下载成功才记账，失败的下次会重试。"""

from __future__ import annotations

import json
import os
from datetime import timedelta
from pathlib import Path
from typing import Dict, Optional

from utils import now_utc_naive


class StateStore:
    def __init__(self, path: Path, retention_days: int = 180, data: Optional[dict] = None) -> None:
        self.path = Path(path)
        self.retention_days = retention_days
        self.data: Dict = data or {"version": 1, "attachments": {}, "messages": {}}
        self.data.setdefault("version", 1)
        self.data.setdefault("attachments", {})
        self.data.setdefault("messages", {})

    @classmethod
    def load(cls, path: Path, retention_days: int = 180) -> "StateStore":
        data: dict = {}
        path = Path(path)
        if path.exists():
            try:
                data = json.loads(path.read_text(encoding="utf-8"))
            except Exception:
                data = {}
                backup = path.with_name(path.name + ".corrupt")
                try:
                    os.replace(str(path), str(backup))
                except Exception:
                    pass
        return cls(path, retention_days, data if isinstance(data, dict) else {})

    def has_attachment(self, key: str) -> bool:
        return key in self.data["attachments"]

    def get_attachment(self, key: str) -> Optional[dict]:
        return self.data["attachments"].get(key)

    def mark_attachment(self, key: str, record: dict) -> None:
        record = dict(record)
        record.setdefault("saved_at", now_utc_naive().isoformat() + "Z")
        self.data["attachments"][key] = record

    def mark_message(self, key: str, record: dict) -> None:
        record = dict(record)
        record.setdefault("processed_at", now_utc_naive().isoformat() + "Z")
        self.data["messages"][key] = record

    def prune(self) -> int:
        """清掉过老的记录，避免状态文件无限膨胀。"""
        if self.retention_days <= 0:
            return 0
        cutoff = (now_utc_naive() - timedelta(days=self.retention_days)).isoformat() + "Z"
        removed = 0
        for bucket, time_field in (("attachments", "saved_at"), ("messages", "processed_at")):
            store: Dict = self.data[bucket]
            for key in [k for k, v in store.items() if str(v.get(time_field, "")) < cutoff]:
                store.pop(key, None)
                removed += 1
        return removed

    def save(self) -> None:
        self.data["updated_at"] = now_utc_naive().isoformat() + "Z"
        self.path.parent.mkdir(parents=True, exist_ok=True)
        tmp = self.path.with_name(self.path.name + ".tmp")
        tmp.write_text(json.dumps(self.data, ensure_ascii=False, indent=2), encoding="utf-8")
        os.replace(str(tmp), str(self.path))

    def counts(self) -> Dict[str, int]:
        return {
            "attachments": len(self.data["attachments"]),
            "messages": len(self.data["messages"]),
        }


class RunLock:
    """防止计划任务把上一次还没跑完的进程又拉起一个。"""

    def __init__(self, path: Path, stale_hours: int = 6) -> None:
        self.path = Path(path)
        self.stale_hours = stale_hours
        self.acquired = False

    @staticmethod
    def _alive(pid: int) -> bool:
        if pid <= 0:
            return False
        try:
            os.kill(pid, 0)
        except OSError:
            return False
        except Exception:
            return True
        return True

    def acquire(self) -> bool:
        if self.path.exists():
            pid, ts = -1, 0.0
            try:
                raw = self.path.read_text(encoding="utf-8").strip().split("|")
                pid, ts = int(raw[0]), float(raw[1])
            except Exception:
                pass
            import time
            if self._alive(pid) and (time.time() - ts) < self.stale_hours * 3600:
                return False
            try:
                self.path.unlink()
            except Exception:
                pass
        import time
        self.path.parent.mkdir(parents=True, exist_ok=True)
        self.path.write_text(f"{os.getpid()}|{time.time()}", encoding="utf-8")
        self.acquired = True
        return True

    def release(self) -> None:
        if not self.acquired:
            return
        try:
            self.path.unlink()
        except Exception:
            pass
        self.acquired = False
