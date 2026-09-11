"""ZIP 解压层。

附件是不可信输入，解压必须防三件事：
1. zip slip —— 成员名含 ../ 或绝对路径，会把文件写到目标目录之外
2. zip bomb —— 一个几 KB 的 zip 解开是几十 GB，或压缩比高到离谱
3. 文件名乱码 —— 非 UTF-8 的 zip（Windows 上用系统编码打包的），
   成员名按 cp437 存，直接取会变成乱码，需要按 gbk 还原
"""

from __future__ import annotations

import os
import zipfile
from pathlib import Path
from typing import Dict, Optional

from utils import (
    clamp_path_length,
    resolve_collision,
    sanitize_component,
    sanitize_filename,
    unique_path,
    utc_to_local_naive,
)

CHUNK = 1024 * 1024
ZIP_SUFFIXES = (".zip",)


class ExtractError(RuntimeError):
    pass


def is_archive(name: str, suffixes=ZIP_SUFFIXES) -> bool:
    return str(name or "").lower().endswith(tuple(suffixes))


def decode_member_name(info: zipfile.ZipInfo) -> str:
    """还原 zip 成员名：UTF-8 标志位优先，否则按 cp437 编码的字节用 gbk 解。"""
    name = info.filename
    if info.flag_bits & 0x800:  # UTF-8 标志位
        return name
    try:
        raw = name.encode("cp437")
    except UnicodeEncodeError:
        return name
    for enc in ("gbk", "utf-8", "shift_jis"):
        try:
            return raw.decode(enc)
        except UnicodeDecodeError:
            continue
    return name


def build_extract_dir(
    template: str,
    target_dir: Path,
    sender: str,
    received_utc,
    subject: str,
    archive_stem: str,
) -> Path:
    """按模板生成解压目录，可用变量比下载目录多一个 {archive}（zip 文件名，不含扩展名）。"""
    local = utc_to_local_naive(received_utc)
    local_part, _, domain = str(sender or "").partition("@")
    mapping = {
        "{sender}": sanitize_component(sender or "unknown"),
        "{sender_local}": sanitize_component(local_part or "unknown"),
        "{sender_domain}": sanitize_component(domain or "unknown"),
        "{subject}": sanitize_component(subject or "no-subject"),
        "{archive}": sanitize_component(archive_stem or "archive"),
        "{date}": local.strftime("%Y-%m-%d"),
        "{year}": local.strftime("%Y"),
        "{month}": local.strftime("%m"),
        "{day}": local.strftime("%d"),
    }
    parts = []
    for raw_part in str(template or "").replace("\\", "/").split("/"):
        if not raw_part.strip():
            continue  # 空模板 / 连续斜杠 -> 直接落在解压根目录
        part = raw_part
        for token, value in mapping.items():
            part = part.replace(token, value)
        part = sanitize_component(part, max_len=120)
        if part:
            parts.append(part)
    return Path(target_dir, *parts)


def _safe_dest(dest_root: Path, member_name: str) -> Optional[Path]:
    """把成员名解析到 dest_root 内，越界返回 None。"""
    normalized = member_name.replace("\\", "/")
    if normalized.startswith("/") or (len(normalized) > 1 and normalized[1] == ":"):
        return None
    candidate = os.path.normpath(str(dest_root / normalized))
    root = os.path.normpath(str(dest_root))
    if candidate != root and not candidate.startswith(root + os.sep):
        return None
    return Path(candidate)


def extract_archive(
    zip_path: Path,
    dest_root: Path,
    overwrite: str = "overwrite",
    max_files: int = 2000,
    max_total_bytes: int = 2 * 1024 * 1024 * 1024,
    max_ratio: float = 200.0,
    final_stem: Optional[str] = None,
    logger=None,
) -> Dict[str, object]:
    """解压一个 zip。返回结果字典，不抛异常（失败用 ok=False 表达）。

    final_stem：只含主干的最终文件名（如 "aaaa"，不含扩展名）。
      仅当 zip 里只有一个文件且传入了它时生效：解出的文件直接按这个名字写出，
      扩展名沿用 zip 内原文件的。撞名时先比对内容——相同则覆盖，不同则另存为 名字_N。
      这样避免了"先按原文件名覆盖掉上一份结果、事后才发现撞名"导致的丢文件。

    返回：{"ok": bool, "files": int, "bytes": int, "paths": [写出的文件路径], "message": str}
    失败时会清理本次已写出的文件与目录。
    """
    result: Dict[str, object] = {"ok": False, "files": 0, "bytes": 0, "paths": [], "message": ""}
    written: list[Path] = []
    created_root = False

    try:
        zip_path = Path(zip_path)
        dest_root = Path(dest_root)
        if not zip_path.exists():
            result["message"] = f"压缩包不存在：{zip_path}"
            return result
        if not dest_root.exists():
            dest_root.mkdir(parents=True, exist_ok=True)
            created_root = True

        with zipfile.ZipFile(zip_path) as zf:
            members = [i for i in zf.infolist() if not i.is_dir()]
            if max_files and len(members) > max_files:
                result["message"] = f"成员数量 {len(members)} 超过上限 {max_files}"
                return result

            header_total = sum(i.file_size for i in members)
            if max_total_bytes and header_total > max_total_bytes:
                result["message"] = (
                    f"解压后总大小约 {header_total / 1024 / 1024:.0f} MB，超过上限 "
                    f"{max_total_bytes / 1024 / 1024:.0f} MB"
                )
                return result

            # 剥离公共顶层目录：压缩软件打包时常套一层文件夹（如 TEST/xxx.pdf），
            # 剥掉后文件直接落在解压根目录，避免 extracted/TEST/xxx.pdf 双层嵌套。
            # 仅当所有成员都在同一个顶层目录下时才剥；zip 根下散落的文件不受影响。
            names = [decode_member_name(i) for i in members]
            splits = [n.replace("\\", "/").split("/", 1) for n in names]
            if len({s[0] for s in splits}) == 1 and all(len(s) == 2 for s in splits):
                names = [s[1] for s in splits]

            use_final = bool(final_stem) and len(members) == 1

            total = 0
            tmp_paths: list[Path] = []
            try:
                for idx, (info, name) in enumerate(zip(members, names)):
                    if max_ratio and info.compress_size > 0:
                        ratio = info.file_size / info.compress_size
                        if ratio > max_ratio:
                            result["message"] = (
                                f"压缩比异常（{ratio:.0f}:1），疑似 zip bomb，已中止：{name}"
                            )
                            return result

                    tmp = None
                    if use_final:
                        suffix = Path(sanitize_filename(Path(name).name)).suffix
                        target = _safe_dest(dest_root, sanitize_filename(final_stem) + suffix)
                        if target is None:
                            result["message"] = f"目标名非法，已中止：{final_stem!r}"
                            return result
                        target = clamp_path_length(target)
                        tmp = dest_root / f".~tmp_{os.getpid()}_{idx}"
                        tmp_paths.append(tmp)
                    else:
                        dest = _safe_dest(dest_root, name)
                        if dest is None:
                            result["message"] = f"成员名越界，已中止（疑似 zip slip）：{info.filename!r}"
                            return result
                        dest = clamp_path_length(dest.parent / sanitize_filename(dest.name))
                        target = unique_path(dest, overwrite)
                        if target is None:
                            continue

                    out_path = tmp if tmp is not None else target
                    out_path.parent.mkdir(parents=True, exist_ok=True)
                    with zf.open(info) as src, open(out_path, "wb") as out:
                        while True:
                            chunk = src.read(CHUNK)
                            if not chunk:
                                break
                            total += len(chunk)
                            if max_total_bytes and total > max_total_bytes:
                                result["message"] = (
                                    f"解压后大小超过上限 {max_total_bytes / 1024 / 1024:.0f} MB，已中止"
                                )
                                return result
                            out.write(chunk)

                    if tmp is not None:
                        # 先比内容再定名：同名且内容相同 -> 覆盖；内容不同 -> 另存，绝不丢文件
                        final = resolve_collision(target, tmp)
                        if final != target and logger is not None:
                            logger.warning("目标同名但内容不同，保留两份：%s（另存为 %s）",
                                           target.name, final.name)
                        tmp.replace(final)
                        tmp_paths.remove(tmp)
                    else:
                        final = target
                    written.append(final)
            finally:
                for leftover in tmp_paths:
                    try:
                        leftover.unlink()
                    except Exception:
                        pass

        result.update(ok=True, files=len(written), bytes=total,
                      paths=[str(p) for p in written],
                      message=f"解压出 {len(written)} 个文件")
        return result

    except zipfile.BadZipFile as exc:
        result["message"] = f"不是有效的 zip 或文件已损坏：{exc}"
        return result
    except RuntimeError as exc:  # zipfile 用 RuntimeError 表示加密
        result["message"] = f"压缩包可能加密了，需要提供密码：{exc}"
        return result
    except Exception as exc:  # noqa: BLE001
        result["message"] = f"{type(exc).__name__}: {exc}"
        return result
    finally:
        if not result["ok"] and written:
            for path in written:
                try:
                    path.unlink()
                except Exception:
                    pass
            if created_root:
                try:
                    dest_root.rmdir()
                except Exception:
                    pass
