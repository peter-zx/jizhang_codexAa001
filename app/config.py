import logging
import os
from dataclasses import dataclass
from pathlib import Path


BASE_DIR = Path(__file__).resolve().parent.parent


def _load_dotenv() -> None:
    env_file = BASE_DIR / ".env"
    if not env_file.exists():
        return
    for raw in env_file.read_text(encoding="utf-8").splitlines():
        line = raw.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, value = line.split("=", 1)
        os.environ.setdefault(key.strip(), value.strip().strip('"').strip("'"))


_load_dotenv()


@dataclass(frozen=True)
class Settings:
    app_name: str = os.getenv("APP_NAME", "分销记账工具")
    app_env: str = os.getenv("APP_ENV", "development")
    secret_key: str = os.getenv("SECRET_KEY", "dev-only-change-me")
    database_url: str = os.getenv("DATABASE_URL", "sqlite:///data/app.db")
    storage_dir: Path = Path(os.getenv("STORAGE_DIR", "data"))
    upload_dir: Path = Path(os.getenv("UPLOAD_DIR", "data/uploads"))
    export_dir: Path = Path(os.getenv("EXPORT_DIR", "data/exports"))
    log_level: str = os.getenv("LOG_LEVEL", "INFO")
    grandmaster_username: str = os.getenv("GRANDMASTER_USERNAME", "admin")
    grandmaster_password: str = os.getenv("GRANDMASTER_PASSWORD", "ChangeMe123!")

    def resolve(self, path: Path) -> Path:
        return path if path.is_absolute() else BASE_DIR / path


settings = Settings()


def setup_logging() -> None:
    log_dir = BASE_DIR / "logs"
    log_dir.mkdir(parents=True, exist_ok=True)
    logging.basicConfig(
        level=getattr(logging, settings.log_level.upper(), logging.INFO),
        format="%(asctime)s %(levelname)s [%(name)s] %(message)s",
        handlers=[
            logging.StreamHandler(),
            logging.FileHandler(log_dir / "app.log", encoding="utf-8"),
        ],
    )
