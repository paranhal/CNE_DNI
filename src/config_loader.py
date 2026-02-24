from __future__ import annotations

import json
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parent.parent
CONFIG_DIR = PROJECT_ROOT / "config"


def _load_json(path: Path) -> dict:
    if not path.exists():
        raise FileNotFoundError(f"설정 파일이 없습니다: {path}")
    with path.open("r", encoding="utf-8") as f:
        return json.load(f)


def load_paths() -> dict[str, str]:
    """
    우선순위:
    1) config/paths.local.json (장비별 로컬 설정)
    2) config/paths.example.json (기본 템플릿)
    """
    local_path = CONFIG_DIR / "paths.local.json"
    example_path = CONFIG_DIR / "paths.example.json"

    if local_path.exists():
        data = _load_json(local_path)
    else:
        data = _load_json(example_path)

    data.setdefault("output_root", "./output")
    data.setdefault("log_root", "./logs")
    return data


def get_path(key: str) -> Path:
    paths = load_paths()
    if key not in paths:
        raise KeyError(f"경로 키가 없습니다: {key}")

    p = Path(paths[key])
    if not p.is_absolute():
        p = (PROJECT_ROOT / p).resolve()
    return p


def get_path_optional(key: str) -> Path | None:
    """경로 반환. 키가 없거나 값이 비어 있으면 None."""
    paths = load_paths()
    val = paths.get(key)
    if val is None or (isinstance(val, str) and val.strip() == ""):
        return None
    p = Path(val)
    if not p.is_absolute():
        p = (PROJECT_ROOT / p).resolve()
    return p


def ensure_runtime_dirs() -> dict[str, Path]:
    output_dir = get_path("output_root")
    log_dir = get_path("log_root")

    output_dir.mkdir(parents=True, exist_ok=True)
    log_dir.mkdir(parents=True, exist_ok=True)

    return {
        "output_root": output_dir,
        "log_root": log_dir,
    }


def check_paths() -> dict[str, bool]:
    """
    주요 경로 존재 여부 점검
    """
    results: dict[str, bool] = {}
    for key in ["raw_data_root", "output_root", "log_root"]:
        try:
            p = get_path(key)
            exists = p.exists()
            results[key] = exists
        except Exception:
            results[key] = False
    return results


if __name__ == "__main__":
    paths = load_paths()
    print("[설정값]")
    for k, v in paths.items():
        print(f"- {k}: {v}")

    print("\n[절대경로 변환]")
    for key in ["raw_data_root", "output_root", "log_root"]:
        try:
            print(f"- {key}: {get_path(key)}")
        except Exception as e:
            print(f"- {key}: ERROR - {e}")

    made = ensure_runtime_dirs()
    print("\n[생성/확인 완료]")
    for k, v in made.items():
        print(f"- {k}: {v}")

    print("\n[경로 존재 여부]")
    checks = check_paths()
    for k, ok in checks.items():
        mark = "OK" if ok else "WARN"
        print(f"- {k}: {mark}")

    if not checks.get("raw_data_root", False):
        print("\n[안내] raw_data_root 폴더가 아직 없거나 경로가 잘못되었습니다.")
        print("      Google Drive 내 원본 데이터 폴더를 확인해 주세요.")
