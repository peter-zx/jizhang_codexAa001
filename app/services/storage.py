import logging
import uuid
from pathlib import Path

from fastapi import UploadFile

from app.config import settings

logger = logging.getLogger(__name__)

ALLOWED_EXTENSIONS = {".xlsx"}


def ensure_storage_dirs() -> None:
    for path in [settings.storage_dir, settings.upload_dir, settings.export_dir]:
        settings.resolve(path).mkdir(parents=True, exist_ok=True)


def unique_name(prefix: str, suffix: str) -> str:
    return f"{prefix}_{uuid.uuid4().hex[:5]}{suffix}"


async def save_upload(file: UploadFile) -> tuple[str, Path]:
    try:
        suffix = Path(file.filename or "").suffix.lower()
        if suffix not in ALLOWED_EXTENSIONS:
            raise ValueError("只支持 .xlsx 文件")
        stored_name = unique_name("upload", suffix)
        target = settings.resolve(settings.upload_dir) / stored_name
        with target.open("wb") as out:
            while chunk := await file.read(1024 * 1024):
                out.write(chunk)
        return stored_name, target
    except Exception:
        logger.exception("保存上传文件失败: %s", file.filename)
        raise


def export_path(prefix: str, suffix: str = ".xlsx") -> tuple[str, Path]:
    stored_name = unique_name(prefix, suffix)
    return stored_name, settings.resolve(settings.export_dir) / stored_name
