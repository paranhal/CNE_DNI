from __future__ import annotations

from datetime import datetime
from pathlib import Path
from typing import Optional
import traceback

from src.common.io_paths import get_log_root


class SimpleLogger:
    def __init__(self, task_name: str, log_dir: Optional[Path] = None):
        self.task_name = self._safe_name(task_name)
        self.log_dir = log_dir or get_log_root(make_dir=True)

        now = datetime.now()
        self.started_at = now
        self.log_file = self.log_dir / f"{now:%Y%m%d}_{self.task_name}.log"

        self._write_line("=" * 80)
        self._write_line(f"[START] {now:%Y-%m-%d %H:%M:%S} | task={self.task_name}")
        self._write_line(f"[LOG]   {self.log_file}")
        self._write_line("=" * 80)

    @staticmethod
    def _safe_name(name: str) -> str:
        text = str(name).strip()
        if not text:
            return "task"
        # Windows 파일명 금지문자 치환
        bad_chars = ['<', '>', ':', '\"', '/', '\\\\', '|', '?', '*']
        for ch in bad_chars:
            text = text.replace(ch, "_")
        return text.replace(" ", "_")

    def _write_line(self, text: str) -> None:
        line = str(text)
        print(line)
        with self.log_file.open("a", encoding="utf-8-sig") as f:
            f.write(line + "\n")

    def _stamp(self, level: str, msg: str) -> str:
        ts = datetime.now().strftime("%H:%M:%S")
        return f"[{ts}] [{level}] {msg}"

    def info(self, msg: str) -> None:
        self._write_line(self._stamp("INFO", msg))

    def ok(self, msg: str) -> None:
        self._write_line(self._stamp("OK", msg))

    def warn(self, msg: str) -> None:
        self._write_line(self._stamp("WARN", msg))

    def error(self, msg: str) -> None:
        self._write_line(self._stamp("ERROR", msg))

    def exception(self, msg: str, exc: Optional[BaseException] = None) -> None:
        self.error(msg)
        if exc is not None:
            self._write_line(self._stamp("ERROR", f"{type(exc).__name__}: {exc}"))
        self._write_line(self._stamp("TRACE", traceback.format_exc().rstrip()))

    def section(self, title: str) -> None:
        self._write_line("-" * 80)
        self._write_line(self._stamp("SECTION", title))
        self._write_line("-" * 80)

    def close(self) -> None:
        ended = datetime.now()
        elapsed = ended - self.started_at
        self._write_line("=" * 80)
        self._write_line(f"[END]   {ended:%Y-%m-%d %H:%M:%S} | elapsed={elapsed}")
        self._write_line("=" * 80)

    # with 문 지원
    def __enter__(self) -> "SimpleLogger":
        return self

    def __exit__(self, exc_type, exc, tb) -> bool:
        if exc is not None:
            self.exception("작업 중 예외 발생", exc)
        self.close()
        # False: 예외를 다시 올림 (기본 동작)
        return False


if __name__ == "__main__":
    # 단독 실행 테스트
    with SimpleLogger("logger_test") as log:
        log.info("로거 테스트 시작")
        log.section("환경 점검")
        log.ok("로그 파일 생성 확인")
        log.warn("샘플 경고 메시지")
        log.info("로거 테스트 종료")
