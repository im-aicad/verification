"""Subprocess worker for wheelhouse regulation verification."""

from __future__ import annotations

import contextlib
import importlib.util
import io
import json
import sys
import traceback
from pathlib import Path
from typing import Any


BASE_DIR = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
ALGORITHM_MAIN = BASE_DIR / "pycatia_wheelhouse_regulation_verification" / "main.py"


def load_algorithm_module() -> Any:
    """
    功能: 加载轮罩法规校核算法模块。
    输入: 无。
    输出: Python 模块对象。
    """
    if not ALGORITHM_MAIN.exists():
        raise FileNotFoundError(f"未找到算法文件: {ALGORITHM_MAIN}")
    spec = importlib.util.spec_from_file_location("wheelhouse_regulation_worker_main", ALGORITHM_MAIN)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"无法加载算法模块: {ALGORITHM_MAIN}")
    module = importlib.util.module_from_spec(spec)
    sys.modules["wheelhouse_regulation_worker_main"] = module
    spec.loader.exec_module(module)
    return module


def write_result(result_path: Path, result: dict[str, Any]) -> None:
    """
    功能: 写入 worker 结果 JSON。
    输入: 结果路径和结果字典。
    输出: JSON 文件。
    """
    result_path.write_text(
        json.dumps(result, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


def main() -> int:
    """
    功能: 子进程入口，运行轮罩法规校核。
    输入: 四个轮罩、车轮装配、结果 JSON、输出目录。
    输出: 进程退出码。
    """
    if len(sys.argv) != 8:
        print(
            "用法: run_algorithm_worker.py <左前轮罩CATPart> <右前轮罩CATPart> <左后轮罩CATPart> <右后轮罩CATPart> <车轮装配CATProduct/CATPart> <结果JSON路径> <输出目录>",
            file=sys.stderr,
        )
        return 2

    def parse_optional_path(text: str) -> Path | None:
        value = str(text or "").strip()
        return Path(value) if value else None

    left_front_path = parse_optional_path(sys.argv[1])
    right_front_path = parse_optional_path(sys.argv[2])
    left_rear_path = parse_optional_path(sys.argv[3])
    right_rear_path = parse_optional_path(sys.argv[4])
    wheel_path = Path(sys.argv[5])
    result_path = Path(sys.argv[6])
    output_dir = Path(sys.argv[7])

    pythoncom = None
    try:
        import pythoncom as _pythoncom  # type: ignore

        pythoncom = _pythoncom
        pythoncom.CoInitialize()
    except Exception:
        pythoncom = None

    try:
        module = load_algorithm_module()
        module.configure_console_encoding()
        result = module.main(
            left_front_wheelhouse_part_path=left_front_path,
            right_front_wheelhouse_part_path=right_front_path,
            left_rear_wheelhouse_part_path=left_rear_path,
            right_rear_wheelhouse_part_path=right_rear_path,
            wheel_assembly_path=wheel_path,
            output_dir=output_dir,
            json_result_path=None,
            save_product=True,
        )
        if not result.get("summary_text"):
            summary_buffer = io.StringIO()
            with contextlib.redirect_stdout(summary_buffer), contextlib.redirect_stderr(summary_buffer):
                module.print_result_summary(result)
            result["summary_text"] = summary_buffer.getvalue().strip()
        write_result(result_path, module.make_json_safe(result))
        return 0 if result.get("success") else 1
    except Exception as exc:
        result = {
            "success": False,
            "error": str(exc),
            "traceback": traceback.format_exc(),
        }
        try:
            write_result(result_path, result)
        except Exception:
            pass
        print(f"检测异常: {exc}", file=sys.stderr)
        print(traceback.format_exc(), file=sys.stderr)
        return 1
    finally:
        if pythoncom is not None:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass


if __name__ == "__main__":
    raise SystemExit(main())
