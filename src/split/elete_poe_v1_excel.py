# 파일명 예: delete_poe_v1_excels.py
# 저장 위치(예시): D:\Lee_20260202\보고서작성\기호준_측정값\delete_poe_v1_excels.py

from pathlib import Path

# 기준 폴더 (스크립트 위치 기준)
BASE_DIR = Path(__file__).resolve().parent
TARGET_ROOT = BASE_DIR / "OUTPUT" / "DNI"

# 삭제 대상 조건
KEYWORD = "POE_V1"
EXCEL_EXTS = {".xlsx", ".xls", ".xlsm", ".xlsb", ".csv"}  # 필요 없으면 .csv 제거

def main():
    if not TARGET_ROOT.exists():
        print(f"[오류] 대상 폴더가 없습니다: {TARGET_ROOT}")
        return

    deleted = []
    failed = []

    # 모든 하위 폴더 순회
    for file_path in TARGET_ROOT.rglob("*"):
        if not file_path.is_file():
            continue

        # 엑셀 확장자 + 파일명에 키워드 포함
        if file_path.suffix.lower() in EXCEL_EXTS and KEYWORD.lower() in file_path.name.lower():
            try:
                file_path.unlink()  # 파일 삭제
                deleted.append(file_path)
                print(f"[삭제] {file_path}")
            except Exception as e:
                failed.append((file_path, str(e)))
                print(f"[실패] {file_path} -> {e}")

    print("\n" + "=" * 80)
    print(f"대상 폴더: {TARGET_ROOT}")
    print(f"삭제 완료: {len(deleted)}건")
    print(f"삭제 실패: {len(failed)}건")

    if failed:
        print("\n[삭제 실패 목록]")
        for p, err in failed:
            print(f"- {p} | {err}")

if __name__ == "__main__":
    main()