"""
测试 `create_distance_visualization_from_points`。

运行前准备:
1. 打开 CATIA。
2. 确保当前活动文档是 CATProduct。
3. 运行本脚本。

本脚本只调用 utils.GeneralClass 中的通用函数，不修改工具内容。
"""

from __future__ import annotations

import sys
from pathlib import Path
from pprint import pprint

import win32com.client as win32


PROJECT_ROOT = Path(__file__).resolve().parents[3]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from utils.GeneralClass import create_distance_visualization_from_points  # noqa: E402


POINT1 = (-466.0, -470.6, 505.8)
POINT2 = (464.4, -470.6, 505.8)
TEXT_OFFSET = (0.0, 0.0, 80.0)
TEXT_SIZE = 8.0
COLOR = (0, 255, 0)
NAME = "__DistanceVisualization_Test__"


def main() -> dict[str, object]:
    """在当前 CATProduct 中创建一条距离可视化标注。"""
    catia = win32.GetActiveObject("CATIA.Application")
    result = create_distance_visualization_from_points(
        POINT1,
        POINT2,
        name=NAME,
        text=None,
        text_unit="mm",
        text_offset=TEXT_OFFSET,
        text_size=TEXT_SIZE,
        color=COLOR,
        carrier_part_number="__DistanceVisualization_Carrier__",
        same_point_tolerance=1e-6,
    )
    if result is None:
        raise RuntimeError("距离可视化创建失败。")

    pprint(
        {
            "status": "success",
            "active_document": getattr(catia.ActiveDocument, "Name", ""),
            "distance": round(float(result.get("distance", 0.0)), 6),
            "label": result.get("label"),
            "document_kind": result.get("document_kind"),
            "text_created": result.get("text_created"),
            "method": result.get("method"),
        }
    )
    return result


if __name__ == "__main__":
    main()
