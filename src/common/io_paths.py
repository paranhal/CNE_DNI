from __future__ import annotations

from pathlib import Path
import json
from typing import Any, Dict


PROJECT_ROOT = Path(__file__).resolve().parents[2]
CONFIG_DIR = PROJECT_ROOT / "config"
LOCAL_CONFIG_PATH = CONFIG_DIR / "paths.local.json"
EXAMPLE_CONFIG_PATH = CONFIG_DIR / "paths.example.json"


class ConfigError(Exception):
    pass


def _read_json_file(json_path: Path) -> Dict[str, Any]:
    if not json_path.exists():
        raise ConfigError(f"설정 파일이 없습니다: {json_path}")

    try:
        text = json_path.read_text(encoding="utf-8-sig")  # BOM 대응
        data = json.loads(text)
    except json.JSONDecodeError as e:
        raise ConfigError(f"JSON 형식 오류: {json_path} | {e}") from e
    except Exception as e:
        raise ConfigError(f"설정 파일 읽기 실패: {json_path} | {e}") from e

    if not isinstance(data, dict):
        raise ConfigError(f"설정 파일 형식 오류(객체 형식 필요): {json_path}")

    return data


def load_paths_config() -> Dict[str, str]:
    cfg = _read_json_file(LOCAL_CONFIG_PATH)

    required_keys = ["RAW_ROOT", "OUT_ROOT", "LOG_ROOT"]
    missing = [k for k in required_keys if not cfg.get(k)]

    if missing:
        missing_text = ", ".join(missing)
        raise ConfigError(
            f"필수 설정값 누락: {missing_text} | 파일 확인: {LOCAL_CONFIG_PATH}"
        )

    cleaned = {k: str(v).strip() for k, v in cfg.items()}
    return cleaned


def get_path(key: str, must_exist: bool = False, make_dir: bool = False) -> Path:
    cfg = load_paths_config()

    if key not in cfg:
        raise ConfigError(f"설정 키가 없습니다: {key}")

    p = Path(cfg[key]).expanduser()

    if make_dir:
        p.mkdir(parents=True, exist_ok=True)

    if must_exist and not p.exists():
        raise ConfigError(f"경로가 존재하지 않습니다: {key} -> {p}")

    return p


def get_raw_root(must_exist: bool = True) -> Path:
    return get_path("RAW_ROOT", must_exist=must_exist, make_dir=False)


def get_out_root(make_dir: bool = True) -> Path:
    return get_path("OUT_ROOT", must_exist=False, make_dir=make_dir)


def get_log_root(make_dir: bool = True) -> Path:
    return get_path("LOG_ROOT", must_exist=False, make_dir=make_dir)


def print_paths_summary() -> None:
    print(f"[PROJECT_ROOT] {PROJECT_ROOT}")
    print(f"[CONFIG]       {LOCAL_CONFIG_PATH}")

    try:
        raw_root = get_raw_root(must_exist=False)
        out_root = get_out_root(make_dir=True)
        log_root = get_log_root(make_dir=True)

        print(f"[RAW_ROOT]     {raw_root}  (exists={raw_root.exists()})")
        print(f"[OUT_ROOT]     {out_root}  (exists={out_root.exists()})")
        print(f"[LOG_ROOT]     {log_root}  (exists={log_root.exists()})")

    except ConfigError as e:
        print("[CONFIG ERROR]")
        print(e)


if __name__ == "__main__":
    print_paths_summary()
