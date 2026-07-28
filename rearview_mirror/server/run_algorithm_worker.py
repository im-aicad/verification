"""Subprocess worker for CATIA automation.

Keeping CATIA COM automation out of the FastAPI process prevents native COM
crashes from taking the web server down.
"""

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
ALGORITHM_MAIN = BASE_DIR / "pycatia_regulation_reflection_point_detection" / "main.py"


def load_algorithm_module() -> Any:
    if not ALGORITHM_MAIN.exists():
        raise FileNotFoundError(f"未找到算法文件: {ALGORITHM_MAIN}")
    spec = importlib.util.spec_from_file_location("rearview_detection_main_worker", ALGORITHM_MAIN)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"无法加载算法模块: {ALGORITHM_MAIN}")
    module = importlib.util.module_from_spec(spec)
    sys.modules["rearview_detection_main_worker"] = module
    spec.loader.exec_module(module)
    return module


def write_result(result_path: Path, result: dict[str, Any]) -> None:
    result_path.write_text(
        json.dumps(result, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


def load_worker_config(config_path: Path | None) -> dict[str, str]:
    if config_path is None or not config_path.exists():
        return {}
    data = json.loads(config_path.read_text(encoding="utf-8"))
    if not isinstance(data, dict):
        return {}
    return {
        str(key): str(value).strip()
        for key, value in data.items()
        if value is not None and str(value).strip()
    }


def main() -> int:
    if len(sys.argv) not in (3, 4):
        print("用法: run_algorithm_worker.py <CATPart路径> <结果JSON路径> [配置JSON路径]", file=sys.stderr)
        return 2

    read_file_path = Path(sys.argv[1])
    result_path = Path(sys.argv[2])
    config_path = Path(sys.argv[3]) if len(sys.argv) == 4 else None

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
        config = load_worker_config(config_path)
        result = module.run_rearview_analysis(
            read_file_path=read_file_path,
            input_parameter_geo_set_name=config.get(
                "input_parameter_geo_set_name",
                module.INPUT_PARAMETER_GEO_SET_NAME,
            ),
            regulation_line_geo_set_name=module.REGULATION_LINE_GEO_SET_NAME,
            parametric_rearview_mirror_geo_set_name=module.PARAMETRIC_REARVIEW_MIRROR_GEO_SET_NAME,
            regulation_reflection_point_geo_set_name=module.REGULATION_REFLECTION_POINT_GEO_SET_NAME,
            gap_check_geo_set_name=module.GAP_CHECK_GEO_SET_NAME,
            left_mirror_feature_name=config.get(
                "left_mirror_feature_name",
                module.LEFT_MIRROR_FEATURE_NAME,
            ),
            right_mirror_feature_name=config.get(
                "right_mirror_feature_name",
                module.RIGHT_MIRROR_FEATURE_NAME,
            ),
            left_eye_point_feature_name=config.get(
                "left_eye_point_feature_name",
                module.LEFT_EYE_POINT_FEATURE_NAME,
            ),
            right_eye_point_feature_name=config.get(
                "right_eye_point_feature_name",
                module.RIGHT_EYE_POINT_FEATURE_NAME,
            ),
            ground_feature_name=config.get("ground_feature_name", module.GROUND_FEATURE_NAME),
            left_vehicle_width_line_feature_name=config.get(
                "left_vehicle_width_line_feature_name",
                module.LEFT_VEHICLE_WIDTH_LINE_FEATURE_NAME,
            ),
            right_vehicle_width_line_feature_name=config.get(
                "right_vehicle_width_line_feature_name",
                module.RIGHT_VEHICLE_WIDTH_LINE_FEATURE_NAME,
            ),
        )

        summary_buffer = io.StringIO()
        with contextlib.redirect_stdout(summary_buffer), contextlib.redirect_stderr(summary_buffer):
            module.print_result_summary(result)
        result["summary_text"] = summary_buffer.getvalue().strip()
        write_result(result_path, result)
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
