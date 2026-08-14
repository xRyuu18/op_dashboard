"""
Config sederhana buat Base URL + API Key OpenProject.

Disimpan di file JSON lokal (config.json), mirip fungsi sheet "Config"
(B1 = Base URL, B2 = API Key) di macro VBA-nya. Kalau file belum ada,
fallback ke environment variable OPENPROJECT_BASE_URL / OPENPROJECT_API_KEY.
"""
import json
import os
from pathlib import Path
from pydantic import BaseModel

CONFIG_PATH = Path(__file__).resolve().parent.parent / "config.json"


class Config(BaseModel):
    base_url: str = ""
    api_key: str = ""


def load_config() -> Config:
    if CONFIG_PATH.exists():
        try:
            data = json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
            return Config(**data)
        except Exception:
            pass
    return Config(
        base_url=os.environ.get("OPENPROJECT_BASE_URL", ""),
        api_key=os.environ.get("OPENPROJECT_API_KEY", ""),
    )


def save_config(cfg: Config) -> None:
    # Base URL nggak boleh diakhiri "/" (sama kayak macro VBA)
    base_url = cfg.base_url.rstrip("/")
    cfg = Config(base_url=base_url, api_key=cfg.api_key)
    CONFIG_PATH.write_text(cfg.model_dump_json(indent=2), encoding="utf-8")
