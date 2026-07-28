"""
测试通用 CATIA 标注工具。

运行前准备:
1. 打开 CATIA。
2. 确保当前活动文档是 CATProduct。
3. 运行本脚本。

本脚本只调用 catia_annotation_tools，不修改工具内容。
"""

from __future__ import annotations

from pprint import pprint

from catia_annotation_tools import create_distance_annotations


TEST_ANNOTATION_ROWS = [
    {
        "point1": (-466.0, -470.6, 505.8),
        "point2": (464.4, -470.6, 505.8),
        "annotation_name": "box1",
        "hybrid_body_name": "box1",
        "text_offset_direction": (0.0, 1.0, 0.0),
        "text_offset_distance": 100.0,
        "feature_color": (0, 255, 0),
        "text_color": (0, 255, 0),
        "line_width": 2,
        "text_size": 8.0,
    },
    {
        "point1": (2445.6, -531.2, 557.4),
        "point2": (3422.5, -531.1, 557.4),
        "annotation_name": "box1",
        "hybrid_body_name": "box2",
        "text_offset_direction": (0.0, 1.0, 0.0),
        "text_offset_distance": 100.0,
        "feature_color": (255, 0, 0),
        "text_color": (255, 0, 0),
        "line_width": 2,
        "text_size": 8.0,
    },
]


def main() -> dict:
    """连接当前 CATIA 活动 CATProduct 并创建测试标注。"""
    result = create_distance_annotations(
        annotations=TEST_ANNOTATION_ROWS,
        process_part_number="WheelTrim_Annotation_Test_Process",
        process_part_name="轮罩标注测试过程Part",
        create_process_part_if_missing=True,
    )
    pprint(
        {
            key: value
            for key, value in result.items()
            if not key.startswith("_")
        }
    )
    return result


if __name__ == "__main__":
    main()

