from __future__ import annotations

import os
import shutil
from pathlib import Path


BASE_DIR = Path(__file__).resolve().parent
DATA_DIR_ENV = "MARKET_NEWS_DATA_DIR"


def data_dir_enabled() -> bool:
    return bool(os.environ.get(DATA_DIR_ENV, "").strip())


def get_data_dir() -> Path:
    raw_value = os.environ.get(DATA_DIR_ENV, "").strip()
    if not raw_value:
        return BASE_DIR
    path = Path(raw_value).expanduser().resolve()
    path.mkdir(parents=True, exist_ok=True)
    return path


def data_path(*parts: str) -> Path:
    return get_data_dir().joinpath(*parts)


def outputs_dir() -> Path:
    return data_path("outputs")


def app_settings_path() -> Path:
    return data_path("web_app_settings.json")


def app_secret_path() -> Path:
    return data_path("web_app_secret.key")


def job_timing_history_path() -> Path:
    return data_path("job_timing_history.json")


def article_star_store_path() -> Path:
    return data_path("article_star_store.json")


def sqlite_db_path() -> Path:
    return data_path("news_crawler.db")


def custom_country_config_path() -> Path:
    return data_path("country_configs_custom.json")


def ensure_from_template(runtime_path: Path, template_path: Path) -> Path:
    if not data_dir_enabled() or runtime_path.exists() or not template_path.exists():
        return runtime_path
    runtime_path.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(template_path, runtime_path)
    return runtime_path


def runtime_project_path(value: str | None, *, copy_from_template: bool = False) -> Path:
    raw_value = str(value or "").strip().replace("\\", "/")
    if not raw_value:
        return get_data_dir() if data_dir_enabled() else BASE_DIR
    candidate = Path(raw_value)
    if candidate.is_absolute() or (len(raw_value) >= 3 and raw_value[1:3] == ":/") or raw_value.startswith("//"):
        return candidate
    if not data_dir_enabled():
        return (BASE_DIR / raw_value).resolve()
    runtime_path = (get_data_dir() / raw_value).resolve()
    if copy_from_template:
        template_path = (BASE_DIR / raw_value).resolve()
        ensure_from_template(runtime_path, template_path)
    return runtime_path


def runtime_output_dir_for_user_value(value: str | None) -> Path:
    raw_value = str(value or "outputs").strip() or "outputs"
    candidate = Path(raw_value)
    if candidate.is_absolute() or (len(raw_value) >= 3 and raw_value[1:3] in {":\\", ":/"}):
        return candidate
    if raw_value.replace("\\", "/").startswith("outputs"):
        suffix = raw_value.replace("\\", "/").split("/", 1)
        return outputs_dir() if len(suffix) == 1 else outputs_dir().joinpath(suffix[1])
    return runtime_project_path(raw_value)
