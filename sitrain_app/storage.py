from __future__ import annotations

import hashlib
import os
import re
import shutil
import zipfile
from pathlib import Path

from fastapi import HTTPException, UploadFile

from .config import (ALLOWED_EXTENSIONS, MAX_ARCHIVE_ENTRIES,
                     MAX_ARCHIVE_UNCOMPRESSED, MAX_UPLOAD_BYTES, STORAGE_ROOT)

_SAFE_NAME = re.compile(r"[^0-9A-Za-z._()\-\u4e00-\u9fff]+")


def safe_filename(name: str) -> str:
    leaf = Path(name or "upload").name
    clean = _SAFE_NAME.sub("_", leaf).strip(" .")
    return clean[:180] or "upload"


def file_dir(file_id: str) -> Path:
    path = (STORAGE_ROOT / file_id[:2] / file_id).resolve()
    if STORAGE_ROOT.resolve() not in path.parents:
        raise ValueError("invalid storage path")
    path.mkdir(parents=True, exist_ok=True)
    return path


def checksum(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


async def upload_checksum(upload: UploadFile) -> str:
    """Compare a retry with the immutable original without saving another copy."""
    h = hashlib.sha256()
    size = 0
    try:
        while chunk := await upload.read(1024 * 1024):
            size += len(chunk)
            if size > MAX_UPLOAD_BYTES:
                raise HTTPException(413, "文件超过上传大小限制")
            h.update(chunk)
    finally:
        await upload.seek(0)
    return h.hexdigest()


async def save_upload(upload: UploadFile, file_id: str) -> Path:
    name = safe_filename(upload.filename or "upload")
    ext = Path(name).suffix.lower()
    if ext not in ALLOWED_EXTENSIONS:
        raise HTTPException(415, f"不支持的文件格式：{ext or '无扩展名'}")
    target = file_dir(file_id) / "original" / name
    target.parent.mkdir(parents=True, exist_ok=True)
    size = 0
    with target.open("xb") as out:
        while chunk := await upload.read(1024 * 1024):
            size += len(chunk)
            if size > MAX_UPLOAD_BYTES:
                out.close()
                target.unlink(missing_ok=True)
                raise HTTPException(413, "文件超过上传大小限制")
            out.write(chunk)
    if size == 0:
        target.unlink(missing_ok=True)
        raise HTTPException(400, "空文件不能上传")
    validate_container(target)
    return target


def validate_container(path: Path) -> None:
    if path.suffix.lower() not in {".pptx", ".docx", ".xlsx"}:
        return
    try:
        with zipfile.ZipFile(path) as archive:
            infos = archive.infolist()
            if len(infos) > MAX_ARCHIVE_ENTRIES:
                raise HTTPException(400, "Office 文件包含过多内部条目")
            total = 0
            for info in infos:
                inner = Path(info.filename.replace("\\", "/"))
                if inner.is_absolute() or ".." in inner.parts:
                    raise HTTPException(400, "Office 文件包含不安全路径")
                total += info.file_size
                if total > MAX_ARCHIVE_UNCOMPRESSED:
                    raise HTTPException(400, "Office 文件解压规模超限")
    except zipfile.BadZipFile as exc:
        raise HTTPException(400, "Office 文件结构损坏或扩展名不符") from exc


def immutable_copy(source: Path, target: Path) -> Path:
    target.parent.mkdir(parents=True, exist_ok=True)
    if target.exists():
        raise FileExistsError(target)
    shutil.copy2(source, target)
    return target
