from __future__ import annotations

import argparse
import hashlib
import importlib.util
import itertools
import json
import math
import shutil
import sys
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any, Iterable

import win32com.client


Vector = tuple[float, float, float]


DEFAULT_PRODUCT_PART_NUMBER = "Wheelhouse_Regulation_Verification"
DEFAULT_PRODUCT_NAME = "Wheelhouse_Regulation_Verification"
DEFAULT_OUTPUT_DIR = Path(__file__).resolve().parent / "output"
AXIS_DIRECTION_TOLERANCE_DEGREES = 2.0
WHEEL_POSITION_TOLERANCE = 50.0
TIRE_HUB_COG_TOLERANCE = 30.0
TIRE_SIZE_ADVANTAGE_RATIO = 1.05
REGULATION_AXIS_HALF_LENGTH = 500.0
REGULATION_AXIS_EXTRUDE_LENGTH = 800.0
REGULATION_AXIS_ROTATION_ANGLE = -30.0
MIN_TOPOLOGY_CIRCLE_COUNT = 4
MIN_TOPOLOGY_RADIUS_LEVELS = 2
TOPOLOGY_AXIS_DIRECTION_TOLERANCE_DEGREES = 3.0
TOPOLOGY_CENTER_CLUSTER_TOLERANCE = 15.0
EXPECTED_TIRE_COUNT = 4
TIRE_BBOX_SIZE_EQUAL_TOLERANCE = 5.0
WHEELHOUSE_NEAR_DISTANCE = 400.0
SECTION_EXTREME_FAR_PLANE_OFFSET = 100000.0
SECTION_EXTREME_FAR_PLANE_SIZE = 200000.0

# ==================== 变量命名区 ====================
# 直接运行本文件时，修改下面三个固定路径即可执行轮罩法规检测主程序。
FRONT_WHEELHOUSE_PART_PATH: str | Path = r"C:\\Users\\Administrator\\Desktop\\catia_project_test\\wheel_house\\11954872_05.CATPart"
REAR_WHEELHOUSE_PART_PATH: str | Path = r"C:\\Users\\Administrator\\Desktop\\catia_project_test\\wheel_house\\12128341_06.CATPart"
WHEEL_ASSEMBLY_PATH: str | Path = r"C:\\Users\\Administrator\\Desktop\\catia_project_test\\wheel\\11940666_04.CATProduct"

OUTPUT_DIR: str | Path | None = DEFAULT_OUTPUT_DIR
JSON_RESULT_PATH: str | Path | None = None
SAVE_PRODUCT_FILE = True
SECTION_CURVE_EXPORT_TOOL_PATH: str | Path = r"C:\Users\Administrator\Desktop\section_curve_export_tool\section_curve_export_tool.py"

PRODUCT_PART_NUMBER = DEFAULT_PRODUCT_PART_NUMBER
PRODUCT_NAME = DEFAULT_PRODUCT_NAME
AXIS_TOLERANCE_DEGREES = AXIS_DIRECTION_TOLERANCE_DEGREES
WHEEL_POSITION_CLUSTER_TOLERANCE = WHEEL_POSITION_TOLERANCE
TIRE_HUB_CENTER_TOLERANCE = TIRE_HUB_COG_TOLERANCE
# ================== 变量命名区结束 ==================

FRONT_WHEELHOUSE_LABEL = "Front_Wheelhouse"
REAR_WHEELHOUSE_LABEL = "Rear_Wheelhouse"
LEFT_FRONT_WHEELHOUSE_LABEL = "Left_Front_Wheelhouse"
RIGHT_FRONT_WHEELHOUSE_LABEL = "Right_Front_Wheelhouse"
LEFT_REAR_WHEELHOUSE_LABEL = "Left_Rear_Wheelhouse"
RIGHT_REAR_WHEELHOUSE_LABEL = "Right_Rear_Wheelhouse"
WHEEL_ASSEMBLY_LABEL = "Wheel_Assembly"
REGULATION_AXIS_PART_NUMBER = "Wheelhouse_Regulation_Axis_Lines"
WHEELHOUSE_ANNOTATION_PART_NUMBER = "Wheelhouse_Regulation_Annotations"
WHEELHOUSE_ANNOTATION_PART_NAME = "Wheelhouse Regulation Annotations"
WHEELHOUSE_ANNOTATION_BODY_NAME = "Wheelhouse Regulation Distance Annotations"
WHEELHOUSE_ANNOTATION_COLOR = (0, 255, 0)
WHEELHOUSE_ANNOTATION_OFFSET_DIRECTION = (0.0, 1.0, 0.0)
WHEELHOUSE_ANNOTATION_BBOX_OFFSET_DIRECTION = (1.0, 0.0, 0.0)
WHEELHOUSE_ANNOTATION_OFFSET_DISTANCE = 150.0
WHEELHOUSE_ANNOTATION_LINE_WIDTH = 2
WHEELHOUSE_ANNOTATION_TEXT_SIZE = 8.0
WHEELHOUSE_SCREENSHOT_BBOX_VIEW_DISTANCE = 5000.0
WHEELHOUSE_SCREENSHOT_AXIS_CLEARANCE_VIEW_DISTANCE = 5000.0
WHEELHOUSE_SCREENSHOT_SECTION_VIEW_DISTANCE = 1000.0
WHEELHOUSE_SCREENSHOT_IMAGE_FORMAT = "png"
WHEELHOUSE_SCREENSHOT_DIRECTIONS = {
    "bbox": {
        "sight_direction": (0.0, 0.0, -1.0),
        "up_direction": (-1.0, 0.0, 0.0),
    },
    "axis_clearance": {
        "sight_direction": (0.0, -1.0, 0.0),
        "up_direction": (0.0, 0.0, 1.0),
    },
    "section": {
        "sight_direction": (-1.0, 0.0, 0.0),
        "up_direction": (0.0, 0.0, 1.0),
    },
}
WHEELHOUSE_ANNOTATION_KEYS = (
    "Left-Front-p",
    "Left-Front-p30",
    "Right-Front-p",
    "Right-Front-p30",
    "Left-Rear-p",
    "Left-Rear-p30",
    "Right-Rear-p",
    "Right-Rear-p30",
    "Left-Front-c",
    "Right-Front-c",
    "Left-Rear-c",
    "Right-Rear-c",
    "Left-Front-q",
    "Right-Front-q",
    "Left-Rear-q",
    "Right-Rear-q",
)
WHEELHOUSE_SCREENSHOT_SEQUENCE = (
    "Left-Front-q",
    "Right-Front-q",
    "Left-Rear-q",
    "Right-Rear-q",
    "Left-Front-c",
    "Right-Front-c",
    "Left-Rear-c",
    "Right-Rear-c",
    "Left-Front-p",
    "Left-Front-p30",
    "Right-Front-p",
    "Right-Front-p30",
    "Left-Rear-p",
    "Left-Rear-p30",
    "Right-Rear-p",
    "Right-Rear-p30",
)
FRONT_REGULATION_GEOMETRY_SET_NAME = "法规校核前轮罩"
REAR_REGULATION_GEOMETRY_SET_NAME = "法规校核后轮罩"
WHEELHOUSE_SLOT_DEFINITIONS: dict[str, dict[str, str]] = {
    "left_front": {
        "label": LEFT_FRONT_WHEELHOUSE_LABEL,
        "measurement_prefix": "Left-Front",
        "component_part_number": "Left_Front_Wheelhouse",
        "geometry_set_name": "法规校核左前轮罩",
        "line_name": "Left_Front_Wheelhouse_Regulation_Axis_Line",
        "section_prefix": "左前轮罩",
        "position": "front",
        "side": "left",
    },
    "right_front": {
        "label": RIGHT_FRONT_WHEELHOUSE_LABEL,
        "measurement_prefix": "Right-Front",
        "component_part_number": "Right_Front_Wheelhouse",
        "geometry_set_name": "法规校核右前轮罩",
        "line_name": "Right_Front_Wheelhouse_Regulation_Axis_Line",
        "section_prefix": "右前轮罩",
        "position": "front",
        "side": "right",
    },
    "left_rear": {
        "label": LEFT_REAR_WHEELHOUSE_LABEL,
        "measurement_prefix": "Left-Rear",
        "component_part_number": "Left_Rear_Wheelhouse",
        "geometry_set_name": "法规校核左后轮罩",
        "line_name": "Left_Rear_Wheelhouse_Regulation_Axis_Line",
        "section_prefix": "左后轮罩",
        "position": "rear",
        "side": "left",
    },
    "right_rear": {
        "label": RIGHT_REAR_WHEELHOUSE_LABEL,
        "measurement_prefix": "Right-Rear",
        "component_part_number": "Right_Rear_Wheelhouse",
        "geometry_set_name": "法规校核右后轮罩",
        "line_name": "Right_Rear_Wheelhouse_Regulation_Axis_Line",
        "section_prefix": "右后轮罩",
        "position": "rear",
        "side": "right",
    },
}
LEGACY_WHEELHOUSE_DEFINITIONS: dict[str, dict[str, str]] = {
    FRONT_WHEELHOUSE_LABEL: {
        "label": FRONT_WHEELHOUSE_LABEL,
        "measurement_prefix": "Left-Front",
        "component_part_number": "Front_Wheelhouse",
        "geometry_set_name": FRONT_REGULATION_GEOMETRY_SET_NAME,
        "line_name": "Front_Wheelhouse_Regulation_Axis_Line",
        "section_prefix": "前轮罩",
        "position": "front",
        "side": "unknown",
    },
    REAR_WHEELHOUSE_LABEL: {
        "label": REAR_WHEELHOUSE_LABEL,
        "measurement_prefix": "Left-Rear",
        "component_part_number": "Rear_Wheelhouse",
        "geometry_set_name": REAR_REGULATION_GEOMETRY_SET_NAME,
        "line_name": "Rear_Wheelhouse_Regulation_Axis_Line",
        "section_prefix": "后轮罩",
        "position": "rear",
        "side": "unknown",
    },
}

AXIS_NAME_KEYWORDS = (
    "axis",
    "axle",
    "centerline",
    "centreline",
    "rotation",
    "rot",
    "轴",
    "轮轴",
    "中心线",
    "旋转",
)
AXIS_SYSTEM_SPECIFIC_KEYWORDS = (
    "axle",
    "wheel",
    "rotation",
    "rot",
    "轮轴",
    "车轮",
    "旋转",
)
TIRE_NAME_KEYWORDS = (
    "tire",
    "tyre",
    "pneu",
    "轮胎",
)
HUB_NAME_KEYWORDS = (
    "hub",
    "rim",
    "wheel_hub",
    "wheelhub",
    "alloy",
    "轮毂",
    "轮辋",
)


@dataclass
class AxisRecord:
    """
    功能: 记录旧版轴方向去重结果。
    输入: 轴方向、特征名和组件信息。
    输出: 数据对象。
    """

    direction: Vector
    feature_name: str
    component_path: str
    component_name: str
    component_part_number: str


@dataclass
class Transform:
    """
    功能: 表示 CATIA 装配位姿矩阵。
    输入: 三个方向轴和原点。
    输出: 可用于点/方向转换的数据对象。
    """

    x_axis: Vector
    y_axis: Vector
    z_axis: Vector
    origin: Vector


@dataclass
class BoundingBox:
    """
    功能: 表示装配世界坐标下的包围盒。
    输入: 最小点、最大点、尺寸和来源。
    输出: 包围盒数据对象。
    """

    min_point: Vector
    max_point: Vector
    size: Vector
    diagonal: float
    volume: float
    source: str


@dataclass
class WheelPartContext:
    """
    功能: 保存车轮装配中一个叶子零件的上下文。
    输入: Product、Part、Document、装配链和标识信息。
    输出: 叶子零件上下文。
    """

    product: Any
    part: Any
    document: Any
    product_chain: list[Any]
    component_path: str
    component_name: str
    component_part_number: str


@dataclass
class WheelCandidate:
    """
    功能: 保存一个有轴车轮候选件。
    输入: 轴线、重心、包围盒、质量体积和告警。
    输出: 车轮候选数据对象。
    """

    context: WheelPartContext
    feature_name: str | None
    axis_source: str
    axis_direction_world: Vector | None
    axis_point_world: Vector | None
    component_cog_world: Vector
    bbox_world: BoundingBox | None
    mass: float | None
    volume: float | None
    topology_circle_count: int
    topology_radius_levels: int
    topology_score: float
    warnings: list[str]
    axis_face_centers_world: tuple[Vector, Vector] | None = None
    wheel_group_key: str | None = None


@dataclass
class TopologyCircleRecord:
    """
    功能: 保存从圆线或圆弧拓扑中读取到的圆信息。
    输入: 特征名称、局部圆心、局部轴方向和半径。
    输出: 圆拓扑记录数据对象。
    """

    feature_name: str
    center_local: Vector
    axis_direction_local: Vector
    radius: float


@dataclass
class WheelhouseInput:
    """
    功能: 保存一个待校核轮罩的输入和命名规则。
    输入: 槽位、标签、路径、测量前缀和 CATIA 命名信息。
    输出: 主流程统一使用的轮罩对象。
    """

    slot: str
    label: str
    path: Path
    work_path: Path | None
    measurement_prefix: str
    component_part_number: str
    geometry_set_name: str
    line_name: str
    section_prefix: str
    position: str
    side: str
    component: Any | None = None


@dataclass
class TopologyAxisAnalysis:
    """
    功能: 保存由圆线拓扑推导出的候选轴和轮胎评分。
    输入: 轴方向、轴线上点、圆数量、半径层级、中心聚类和评分。
    输出: 拓扑轴分析数据对象。
    """

    axis_direction_world: Vector
    axis_point_world: Vector
    circle_count: int
    radius_levels: int
    center_cluster_radius: float
    min_radius: float
    max_radius: float
    score: float
    feature_names: list[str]


@dataclass
class WheelPositionGroup:
    """
    功能: 保存同一空间车轮位置的一组候选件。
    输入: 分组编号、候选列表和 Tire/排除项。
    输出: 车轮位置组数据对象。
    """

    group_id: str
    candidates: list[WheelCandidate]
    tire_candidate: WheelCandidate | None
    excluded_candidates: list[WheelCandidate]
    warnings: list[str]


@dataclass
class RegulationAxisSegment:
    """
    功能: 保存法规校核轴线段的计算结果。
    输入: 轮罩标签、线段名称、端点、方向、中心点、长度和对应 Tire 信息。
    输出: 轴线段数据对象。
    """

    wheelhouse_label: str
    geometry_set_name: str
    line_name: str
    start_point: Vector
    end_point: Vector
    axis_direction_world: Vector
    axis_point_world: Vector
    center_point_world: Vector
    axis_length: float
    half_length: float
    extrude_length: float
    rotation_angle: float
    tire_component_path: str


@dataclass
class RunOutputPaths:
    """
    功能: 保存单次程序运行的统一输出目录。
    输入: 运行时间戳、运行根目录、结果目录、过程目录和 JSON 路径。
    输出: 运行输出路径对象。
    """

    timestamp: str
    run_root_dir: Path
    result_dir: Path
    course_dir: Path
    json_result_path: Path


def configure_console_encoding() -> None:
    """
    功能: 尽量让 Windows 终端正确输出中文。
    输入: 无。
    输出: 无。
    """
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass


def build_timestamp() -> str:
    """
    功能: 生成文件名时间戳。
    输入: 无。
    输出: yyyyMMdd_HHmmss 字符串。
    """
    return datetime.now().strftime("%Y%m%d_%H%M%S")


def create_run_output_paths(output_dir: str | Path | None = None) -> RunOutputPaths:
    """
    功能: 在 output 目录下创建本次运行的 result/course 输出目录。
    输入: 输出根目录。
    输出: RunOutputPaths。
    """
    timestamp = build_timestamp()
    base_dir = Path(output_dir).expanduser().resolve() if output_dir else DEFAULT_OUTPUT_DIR
    run_root_dir = base_dir / f"output-{timestamp}"
    result_dir = run_root_dir / f"result-{timestamp}"
    course_dir = run_root_dir / f"course-{timestamp}"
    result_dir.mkdir(parents=True, exist_ok=True)
    course_dir.mkdir(parents=True, exist_ok=True)
    return RunOutputPaths(
        timestamp=timestamp,
        run_root_dir=run_root_dir,
        result_dir=result_dir,
        course_dir=course_dir,
        json_result_path=result_dir / f"wheelhouse_regulation_verification_result_{timestamp}.json",
    )


def build_product_save_path(output_dir: str | Path | None = None, timestamp: str | None = None) -> Path:
    """
    功能: 构造校核 CATProduct 保存路径。
    输入: 可选输出目录。
    输出: CATProduct 文件路径。
    """
    target_dir = Path(output_dir).expanduser().resolve() if output_dir else DEFAULT_OUTPUT_DIR
    target_dir.mkdir(parents=True, exist_ok=True)
    return target_dir / f"wheelhouse_regulation_verification_{timestamp or build_timestamp()}.CATProduct"


def build_process_part_save_path(part_number: str, output_dir: str | Path | None = None) -> Path:
    """
    功能: 构造过程 CATPart 的唯一保存路径。
    输入: 过程 PartNumber 和可选输出目录。
    输出: CATPart 文件路径。
    """
    target_dir = Path(output_dir).expanduser().resolve() if output_dir else DEFAULT_OUTPUT_DIR
    target_dir.mkdir(parents=True, exist_ok=True)
    return target_dir / f"{safe_filename_text(part_number)}.CATPart"


def copy_wheelhouse_to_temp(
    source_path: Path,
    output_dir: str | Path | None,
    label: str,
    timestamp: str | None = None,
) -> Path:
    """
    功能: 将原始轮罩 CATPart 复制到临时计算目录。
    输入: 原始轮罩路径、输出目录和轮罩标签。
    输出: 临时轮罩 CATPart 路径。
    """
    temp_dir = Path(output_dir).expanduser().resolve() if output_dir else DEFAULT_OUTPUT_DIR
    temp_dir.mkdir(parents=True, exist_ok=True)
    temp_path = temp_dir / f"{label}_work_{timestamp or build_timestamp()}{source_path.suffix}"
    shutil.copy2(source_path, temp_path)
    return temp_path


def copy_part_to_section_work(
    source_path: Path,
    output_dir: str | Path | None,
    label: str,
    timestamp: str | None = None,
) -> Path:
    """
    功能: 将轮罩工作 CATPart 再复制一份作为相交截面专用文件。
    输入: 轮罩工作文件路径、输出目录和轮罩标签。
    输出: 截面专用 CATPart 路径。
    """
    temp_dir = Path(output_dir).expanduser().resolve() if output_dir else DEFAULT_OUTPUT_DIR
    temp_dir.mkdir(parents=True, exist_ok=True)
    temp_path = temp_dir / f"{label}_section_source_{timestamp or build_timestamp()}{source_path.suffix}"
    shutil.copy2(source_path, temp_path)
    return temp_path


def wheelhouse_input_from_path(slot: str, path: str | Path, work_path: str | Path | None = None) -> WheelhouseInput:
    definition = WHEELHOUSE_SLOT_DEFINITIONS[slot]
    return WheelhouseInput(
        slot=slot,
        label=definition["label"],
        path=Path(path),
        work_path=Path(work_path) if work_path is not None else None,
        measurement_prefix=definition["measurement_prefix"],
        component_part_number=definition["component_part_number"],
        geometry_set_name=definition["geometry_set_name"],
        line_name=definition["line_name"],
        section_prefix=definition["section_prefix"],
        position=definition["position"],
        side=definition["side"],
    )


def legacy_wheelhouse_input(
    slot: str,
    label: str,
    path: str | Path,
    work_path: str | Path | None = None,
) -> WheelhouseInput:
    definition = LEGACY_WHEELHOUSE_DEFINITIONS[label]
    return WheelhouseInput(
        slot=slot,
        label=definition["label"],
        path=Path(path),
        work_path=Path(work_path) if work_path is not None else None,
        measurement_prefix=definition["measurement_prefix"],
        component_part_number=definition["component_part_number"],
        geometry_set_name=definition["geometry_set_name"],
        line_name=definition["line_name"],
        section_prefix=definition["section_prefix"],
        position=definition["position"],
        side=definition["side"],
    )


def wheelhouse_short_token(wheelhouse_label: Any) -> str:
    prefix = wheelhouse_measurement_prefix(wheelhouse_label)
    words = [word for word in prefix.replace("_", "-").split("-") if word]
    if words:
        return "".join(word[:1].upper() for word in words)
    return short_stable_token(wheelhouse_label, length=4).upper()


def safe_filename_text(text: Any) -> str:
    """
    功能: 将文本转换为适合 Windows 文件名的字符串。
    输入: 任意文本。
    输出: 安全文件名片段。
    """
    raw = str(text or "unnamed")
    for char in '<>:"/\\|?*':
        raw = raw.replace(char, "_")
    return raw.replace("°", "deg").replace(" ", "_")


def short_stable_token(text: Any, length: int = 8) -> str:
    """
    功能: 为长名称生成稳定短标识。
    输入: 任意文本和长度。
    输出: 十六进制短标识。
    """
    raw = str(text or "unnamed").encode("utf-8", errors="ignore")
    return hashlib.sha1(raw).hexdigest()[:length]


def compact_section_run_id(
    wheelhouse_label: Any,
    angle_suffix: str,
    target_index: int,
    target_name: Any,
    timestamp: str | None = None,
) -> str:
    """
    功能: 生成较短的截面导出运行编号，避免 CATIA 文件名或 PartNumber 过长。
    输入: 轮罩标签、角度后缀、目标序号和目标名称。
    输出: 短运行编号。
    """
    wheelhouse_token = wheelhouse_short_token(wheelhouse_label)
    angle_token = "30" if "30" in str(angle_suffix) else "0"
    target_token = short_stable_token(target_name, length=6)
    return f"{wheelhouse_token}_{angle_token}_T{target_index:03d}_{target_token}"


def make_embedded_section_target_hooks(
    section_tool: Any,
    part: Any,
    target_row: dict[str, Any],
    unique_target_name: str,
) -> tuple[Any, Any]:
    """
    功能: 为内嵌截面工具创建精确目标钩子，避免同名 Body 按名称解析到同一对象。
    输入: 截面工具模块、Part、目标行和唯一目标名。
    输出: target_candidates 和 reference_from_any_name 两个替换函数。
    """
    original_reference_from_any_name = section_tool.reference_from_any_name
    target_object = target_row.get("object")
    target_kind = str(target_row.get("kind") or "body_or_named_reference")
    measurement = target_row.get("measurement") or {}
    candidate = {
        "kind": target_kind,
        "name": unique_target_name,
        "source_name": target_row.get("name"),
        "label": target_row.get("label"),
        "target_index": target_row.get("index"),
        "provided_by_main": True,
        "area": measurement.get("area"),
        "length": measurement.get("length"),
        "volume": measurement.get("volume"),
        "cog": measurement.get("cog"),
    }

    def target_candidates_override(_document: Any, _part: Any, _target_name: str | None, _surface_limit: int) -> tuple[list[dict[str, Any]], list[dict[str, Any]], list[dict[str, Any]]]:
        surface_rows = [candidate] if target_kind == "surface" else []
        body_rows = [candidate] if target_kind != "surface" else []
        return [candidate], surface_rows, body_rows

    def reference_from_any_name_override(current_part: Any, name: str) -> tuple[Any, dict[str, Any]]:
        if name == unique_target_name:
            if target_object is None:
                raise RuntimeError(f"目标对象为空，无法创建 Reference: {target_row.get('label')}")
            try:
                return current_part.CreateReferenceFromObject(target_object), {
                    "source_type": target_kind,
                    "source_name": target_row.get("name"),
                    "label": target_row.get("label"),
                    "target_index": target_row.get("index"),
                    "provided_by_main": True,
                    "unique_target_name": unique_target_name,
                }
            except Exception as exc:
                raise RuntimeError(f"无法为精确目标创建 Reference: {target_row.get('label')}: {exc}") from exc
        return original_reference_from_any_name(current_part, name)

    return target_candidates_override, reference_from_any_name_override


def resolve_section_curve_export_tool_path() -> Path:
    """
    功能: 获取独立截面曲线导出工具的脚本路径。
    输入: 无。
    输出: section_curve_export_tool.py 的绝对路径。
    """
    configured_path = Path(SECTION_CURVE_EXPORT_TOOL_PATH).expanduser()
    if configured_path.is_file():
        return configured_path.resolve()
    local_path = Path(__file__).resolve().with_name("section_curve_export_tool.py")
    if local_path.is_file():
        return local_path
    raise FileNotFoundError(f"未找到截面曲线导出工具: {configured_path}")


def format_float_for_cli(value: float) -> str:
    """
    功能: 将浮点数格式化为命令行参数文本。
    输入: 浮点数。
    输出: 稳定的小数字符串。
    """
    text = f"{float(value):.12g}"
    return "0" if text == "-0" else text


def format_plane_equation_for_cli(equation: Iterable[float]) -> str:
    """
    功能: 将平面方程转换为独立工具需要的 A,B,C,D 文本。
    输入: 平面方程数值序列。
    输出: 逗号分隔的命令行字符串。
    """
    return ",".join(format_float_for_cli(float(value)) for value in equation)


def validate_existing_catia_file(path: str | Path, label: str) -> Path:
    """
    功能: 校验输入 CATIA 文件路径。
    输入: 文件路径和显示标签。
    输出: 规范化后的 Path。
    """
    if not str(path).strip():
        raise ValueError(f"{label}路径未配置，请先修改 main.py 顶部固定变量。")
    file_path = Path(path).expanduser().resolve()
    if not file_path.is_file():
        raise FileNotFoundError(f"{label}文件不存在: {file_path}")
    if file_path.suffix.casefold() not in {".catpart", ".catproduct"}:
        raise ValueError(f"{label}必须是 CATPart 或 CATProduct 文件: {file_path}")
    return file_path


def start_or_connect_catia() -> Any:
    """
    功能: 连接已打开的 CATIA，未启动时自动启动。
    输入: 无。
    输出: CATIA Application 对象。
    """
    try:
        catia = win32com.client.GetActiveObject("CATIA.Application")
    except Exception:
        catia = win32com.client.Dispatch("CATIA.Application")
    catia.Visible = True
    return catia


def iter_collection(collection: Any) -> Iterable[Any]:
    """
    功能: 遍历 CATIA 1 基集合。
    输入: CATIA 集合对象。
    输出: 集合元素迭代器。
    """
    try:
        count = int(collection.Count)
    except Exception:
        return
    for index in range(1, count + 1):
        yield collection.Item(index)


def collection_count(collection: Any) -> int:
    """
    功能: 安全读取 CATIA 集合数量。
    输入: CATIA 集合对象。
    输出: 数量，失败时为 0。
    """
    try:
        return int(collection.Count)
    except Exception:
        return 0


def safe_attr_text(target: Any, attribute_name: str, default: str = "") -> str:
    """
    功能: 安全读取对象文本属性。
    输入: 对象、属性名和默认值。
    输出: 字符串属性值。
    """
    try:
        value = getattr(target, attribute_name)
        if value is not None:
            return str(value).strip()
    except Exception:
        pass
    return default


def product_display_name(product: Any) -> str:
    """
    功能: 获取 Product 的显示名称。
    输入: Product 对象。
    输出: Name、PartNumber 或占位名称。
    """
    return (
        safe_attr_text(product, "Name")
        or safe_attr_text(product, "PartNumber")
        or "<unnamed product>"
    )


def product_part_number(product: Any) -> str:
    """
    功能: 获取 Product 的零件编号。
    输入: Product 对象。
    输出: PartNumber 或显示名称。
    """
    return safe_attr_text(product, "PartNumber", product_display_name(product))


def set_if_possible(target: Any, attribute_name: str, value: str) -> None:
    """
    功能: 尝试设置 COM 对象属性。
    输入: 对象、属性名和值。
    输出: 无，失败时忽略。
    """
    try:
        setattr(target, attribute_name, value)
    except Exception:
        pass


def add_component_from_file_to_product(document: Any, root_product: Any, file_path: Path) -> Any:
    """
    功能: 将 CATPart 或 CATProduct 文件装配到根产品，并用 VBA 处理 SAFEARRAY 参数。
    输入: ProductDocument、根 Product 和文件路径。
    输出: 新增的子 Product。
    """
    vba_code = """
Public Function add_component(rootProduct, componentPath)
    Dim fileList(0)
    fileList(0) = componentPath
    rootProduct.Products.AddComponentsFromFiles fileList, "All"
    add_component = True
End Function
"""
    before_count = collection_count(root_product.Products)
    try:
        document.Application.SystemService.Evaluate(
            vba_code,
            0,
            "add_component",
            [root_product, str(file_path)],
        )
    except Exception as exc:
        raise RuntimeError(f"无法装配文件: {file_path}") from exc

    after_count = collection_count(root_product.Products)
    if after_count > before_count:
        return root_product.Products.Item(after_count)
    raise RuntimeError(f"文件已装配但无法定位新增组件: {file_path}")


def save_document_if_modified(document: Any, fallback_path: Path | None = None) -> dict[str, Any]:
    """
    功能: 保存已修改的 CATIA 文档，未保存文档可使用 fallback_path 执行 SaveAs。
    输入: CATIA Document 和可选保存路径。
    输出: 保存结果字典。
    """
    document_name = safe_attr_text(document, "Name", "<unknown>")
    try:
        is_saved = bool(getattr(document, "Saved"))
    except Exception:
        is_saved = False
    if is_saved:
        return {
            "status": "skipped",
            "message": "文档未修改。",
            "document": document_name,
        }
    try:
        full_name = safe_attr_text(document, "FullName")
        if full_name:
            document.Save()
            saved_path = full_name
        elif fallback_path is not None:
            fallback_path.parent.mkdir(parents=True, exist_ok=True)
            document.SaveAs(str(fallback_path))
            saved_path = str(fallback_path)
        else:
            return {
                "status": "skipped",
                "message": "文档未保存且未提供 SaveAs 路径。",
                "document": document_name,
            }
        return {
            "status": "success",
            "message": "文档已保存。",
            "document": document_name,
            "path": saved_path,
        }
    except Exception as exc:
        return {
            "status": "failed",
            "message": str(exc),
            "document": document_name,
        }


def create_wheelhouse_product_from_inputs(
    catia: Any,
    wheelhouse_inputs: list[WheelhouseInput],
    wheel_assembly_path: Path,
    product_part_number: str = DEFAULT_PRODUCT_PART_NUMBER,
    product_name: str = DEFAULT_PRODUCT_NAME,
) -> tuple[Any, Any, dict[str, Any], Any]:
    """
    功能: 创建校核 CATProduct 并装配 1-4 个轮罩和车轮装配。
    输入: CATIA、轮罩输入列表和车轮装配路径。
    输出: ProductDocument、根 Product、轮罩组件映射和车轮装配组件。
    """
    product_document = catia.Documents.Add("Product")
    root_product = product_document.Product
    set_if_possible(root_product, "PartNumber", product_part_number)
    set_if_possible(root_product, "Name", product_name)

    wheelhouse_components: dict[str, Any] = {}
    for item in wheelhouse_inputs:
        component = add_component_from_file_to_product(
            product_document,
            root_product,
            item.work_path or item.path,
        )
        set_if_possible(component, "Name", item.label)
        set_if_possible(component, "PartNumber", item.component_part_number)
        wheelhouse_components[item.label] = component
    wheel_assembly_component = add_component_from_file_to_product(
        product_document,
        root_product,
        wheel_assembly_path,
    )

    set_if_possible(wheel_assembly_component, "Name", WHEEL_ASSEMBLY_LABEL)
    set_if_possible(wheel_assembly_component, "PartNumber", "Wheel_Assembly")

    return (
        product_document,
        root_product,
        wheelhouse_components,
        wheel_assembly_component,
    )


def create_wheelhouse_product(
    catia: Any,
    front_wheelhouse_path: Path,
    rear_wheelhouse_path: Path,
    wheel_assembly_path: Path,
    product_part_number: str = DEFAULT_PRODUCT_PART_NUMBER,
    product_name: str = DEFAULT_PRODUCT_NAME,
) -> tuple[Any, Any, Any, Any, Any]:
    """
    功能: 兼容旧流程的双轮罩创建入口。
    输入: CATIA 和三个文件路径。
    输出: ProductDocument、根 Product 和三个组件。
    """
    front_input = legacy_wheelhouse_input("front", FRONT_WHEELHOUSE_LABEL, front_wheelhouse_path)
    rear_input = legacy_wheelhouse_input("rear", REAR_WHEELHOUSE_LABEL, rear_wheelhouse_path)
    product_document, root_product, wheelhouse_components, wheel_assembly_component = create_wheelhouse_product_from_inputs(
        catia,
        [front_input, rear_input],
        wheel_assembly_path,
        product_part_number=product_part_number,
        product_name=product_name,
    )
    return (
        product_document,
        root_product,
        wheelhouse_components[FRONT_WHEELHOUSE_LABEL],
        wheelhouse_components[REAR_WHEELHOUSE_LABEL],
        wheel_assembly_component,
    )


def get_child_products(product: Any) -> list[Any]:
    """
    功能: 获取 Product 的直接子产品。
    输入: Product 对象。
    输出: 子 Product 列表。
    """
    try:
        products = product.Products
    except Exception:
        return []
    return list(iter_collection(products))


def get_part_and_document_from_product(product: Any) -> tuple[Any, Any]:
    """
    功能: 从 Product 实例获取 Part 和 PartDocument。
    输入: Product 对象。
    输出: (Part, Document)。
    """
    attempts = (
        lambda: product.ReferenceProduct.Parent,
        lambda: product.Parent,
    )
    for attempt in attempts:
        try:
            document = attempt()
            part = document.Part
            return part, document
        except Exception:
            pass
    try:
        part = product.ReferenceProduct.Parent.GetItem("Part")
        document = part.Parent
        return part, document
    except Exception as exc:
        raise RuntimeError(f"无法从组件获取 Part: {product_display_name(product)}") from exc


def save_modified_component_documents(
    root_product: Any,
    exclude_documents: Iterable[Any] | None = None,
) -> list[dict[str, Any]]:
    """
    功能: 保存总成树中已修改的子 CATPart/CATProduct 文档。
    输入: 根 Product 和需要排除的文档列表。
    输出: 每个文档的保存结果列表。
    """
    excluded_ids = {id(document) for document in (exclude_documents or [])}
    visited_documents: set[int] = set()
    save_results: list[dict[str, Any]] = []

    def visit(product: Any) -> None:
        try:
            _part, document = get_part_and_document_from_product(product)
        except Exception:
            document = None
        if document is not None:
            document_id = id(document)
            if document_id not in excluded_ids and document_id not in visited_documents:
                visited_documents.add(document_id)
                save_result = save_document_if_modified(document)
                save_results.append(save_result)
                if save_result.get("status") == "failed":
                    print(f"[警告] 子文档保存失败: {save_result.get('document')}: {save_result.get('message')}")
        for child in get_child_products(product):
            visit(child)

    for child_product in get_child_products(root_product):
        visit(child_product)
    return save_results


def iter_leaf_part_contexts(
    root_product: Any,
    root_path: str | None = None,
    product_chain: list[Any] | None = None,
    *,
    visibility_document: Any | None = None,
    visible_only: bool = False,
    visibility_skipped: list[dict[str, Any]] | None = None,
) -> list[WheelPartContext]:
    """
    功能: 递归遍历装配树并收集叶子零件上下文。
    输入: 根 Product、路径、装配链和可选显隐过滤参数。
    输出: WheelPartContext 列表。
    """
    contexts: list[WheelPartContext] = []
    current_name = product_display_name(root_product)
    current_path = root_path or current_name
    current_chain = [*(product_chain or []), root_product]
    if visible_only and visibility_document is not None:
        visibility = get_object_visibility(visibility_document, root_product)
        if visibility.get("status") == "hidden" or visibility.get("visible") is False:
            if visibility_skipped is not None:
                visibility_skipped.append(
                    {
                        "component_path": current_path,
                        "component_name": current_name,
                        "component_part_number": product_part_number(root_product),
                        "visibility": visibility,
                        "reason": "hidden_product_skipped_before_cog",
                    }
                )
            return contexts
    children = get_child_products(root_product)

    if children:
        for index, child in enumerate(children, start=1):
            child_name = product_display_name(child)
            contexts.extend(
                iter_leaf_part_contexts(
                    child,
                    f"{current_path}/{index:03d}_{child_name}",
                    current_chain,
                    visibility_document=visibility_document,
                    visible_only=visible_only,
                    visibility_skipped=visibility_skipped,
                )
            )
        return contexts

    try:
        part, document = get_part_and_document_from_product(root_product)
    except Exception:
        return contexts

    contexts.append(
        WheelPartContext(
            product=root_product,
            part=part,
            document=document,
            product_chain=current_chain,
            component_path=current_path,
            component_name=current_name,
            component_part_number=product_part_number(root_product),
        )
    )
    return contexts


def create_reference(part: Any, feature: Any) -> Any:
    """
    功能: 创建 CATIA Reference。
    输入: Part 和特征对象。
    输出: Reference 对象。
    """
    return part.CreateReferenceFromObject(feature)


def evaluate_measurable_array(
    document: Any,
    measurable: Any,
    method_name: str,
    item_count: int,
    label: str = "",
) -> tuple[float, ...]:
    """
    功能: 调用 Measurable 的数组输出方法。
    输入: Document、Measurable、方法名、数组长度和标签。
    输出: 浮点元组。
    """
    function_name = f"read_{method_name.casefold()}"
    vba_code = f"""
Public Function {function_name}(measurable)
    Dim values({item_count - 1})
    measurable.{method_name} values
    {function_name} = values
End Function
"""
    try:
        values = document.Application.SystemService.Evaluate(
            vba_code,
            0,
            function_name,
            [measurable],
        )
        return tuple(float(value) for value in values)
    except Exception as exc:
        target = f"，对象: {label}" if label else ""
        raise RuntimeError(f"CATIA 测量方法 {method_name} 执行失败{target}") from exc


def evaluate_product_array(
    document: Any,
    product: Any,
    vba_code: str,
    function_name: str,
    label: str,
) -> tuple[float, ...]:
    """
    功能: 通过 VBA 读取 Product 相关数组数据。
    输入: Document、Product、VBA 代码、函数名和标签。
    输出: 浮点元组。
    """
    try:
        values = document.Application.SystemService.Evaluate(
            vba_code,
            0,
            function_name,
            [product],
        )
        return tuple(float(value) for value in values)
    except Exception as exc:
        raise RuntimeError(f"无法读取{label}: {product_display_name(product)}") from exc


def identity_transform() -> Transform:
    """
    功能: 生成单位装配变换。
    输入: 无。
    输出: Transform。
    """
    return Transform((1.0, 0.0, 0.0), (0.0, 1.0, 0.0), (0.0, 0.0, 1.0), (0.0, 0.0, 0.0))


def product_position_transform(document: Any, product: Any) -> Transform:
    """
    功能: 读取 Product.Position 并转换为 Transform。
    输入: Document 和 Product。
    输出: 装配位姿，失败时为单位变换。
    """
    vba_code = """
Public Function read_product_position(product)
    Dim values(11)
    product.Position.GetComponents values
    read_product_position = values
End Function
"""
    try:
        values = evaluate_product_array(
            document,
            product,
            vba_code,
            "read_product_position",
            "装配位置矩阵",
        )
    except Exception:
        return identity_transform()
    if len(values) < 12:
        return identity_transform()
    return Transform(
        as_vector(values[0:3], "Position X Axis"),
        as_vector(values[3:6], "Position Y Axis"),
        as_vector(values[6:9], "Position Z Axis"),
        as_vector(values[9:12], "Position Origin"),
    )


def apply_transform_to_direction(transform: Transform, direction: Vector) -> Vector:
    """
    功能: 将局部方向向量转换到父坐标系。
    输入: Transform 和方向向量。
    输出: 转换后的方向向量。
    """
    return (
        transform.x_axis[0] * direction[0]
        + transform.y_axis[0] * direction[1]
        + transform.z_axis[0] * direction[2],
        transform.x_axis[1] * direction[0]
        + transform.y_axis[1] * direction[1]
        + transform.z_axis[1] * direction[2],
        transform.x_axis[2] * direction[0]
        + transform.y_axis[2] * direction[1]
        + transform.z_axis[2] * direction[2],
    )


def apply_transform_to_point(transform: Transform, point: Vector) -> Vector:
    """
    功能: 将局部点转换到父坐标系。
    输入: Transform 和点坐标。
    输出: 转换后的点坐标。
    """
    direction_part = apply_transform_to_direction(transform, point)
    return (
        direction_part[0] + transform.origin[0],
        direction_part[1] + transform.origin[1],
        direction_part[2] + transform.origin[2],
    )


def inverse_transform_point(transform: Transform, point: Vector) -> Vector:
    """
    功能: 将父坐标系点转换回局部坐标系。
    输入: Transform 和父坐标系点。
    输出: 局部坐标点。
    """
    offset = subtract_vectors(point, transform.origin)
    return (
        dot_product(offset, transform.x_axis),
        dot_product(offset, transform.y_axis),
        dot_product(offset, transform.z_axis),
    )


def compose_transform(parent: Transform, child: Transform) -> Transform:
    """
    功能: 组合父子装配变换。
    输入: 父 Transform 和子 Transform。
    输出: 合成后的 Transform。
    """
    return Transform(
        apply_transform_to_direction(parent, child.x_axis),
        apply_transform_to_direction(parent, child.y_axis),
        apply_transform_to_direction(parent, child.z_axis),
        apply_transform_to_point(parent, child.origin),
    )


def product_chain_world_transform(document: Any, product_chain: list[Any]) -> Transform:
    """
    功能: 累乘叶子零件到根装配的位姿。
    输入: Document 和 Product 链。
    输出: 世界坐标 Transform。
    """
    transform = identity_transform()
    for product in product_chain:
        transform = compose_transform(transform, product_position_transform(document, product))
    return transform


def evaluate_product_cog(document: Any, product: Any, label: str) -> Vector:
    """
    功能: 读取 Product 在当前装配中的重心坐标。
    输入: ProductDocument、Product 和显示标签。
    输出: 世界坐标系下的重心坐标。
    """
    analyze_code = """
Public Function read_product_analyze_cog(product)
    Dim values(2)
    product.Analyze.GetGravityCenter values
    read_product_analyze_cog = values
End Function
"""
    inertia_code = """
Public Function read_product_inertia_cog(product)
    Dim inertia
    Set inertia = product.GetTechnologicalObject("Inertia")
    Dim values(2)
    inertia.GetCOGPosition values
    read_product_inertia_cog = values
End Function
"""
    for function_name, vba_code in (
        ("read_product_analyze_cog", analyze_code),
        ("read_product_inertia_cog", inertia_code),
    ):
        try:
            values = document.Application.SystemService.Evaluate(
                vba_code,
                0,
                function_name,
                [product],
            )
            return as_vector(values, f"{label}重心")
        except Exception:
            pass
    raise RuntimeError(f"无法读取组件重心: {label}")


def evaluate_product_scalar(product: Any, attribute_name: str) -> float | None:
    """
    功能: 读取 Product.Analyze 的标量属性。
    输入: Product 和属性名。
    输出: 浮点值或 None。
    """
    try:
        analyze = product.Analyze
        value = getattr(analyze, attribute_name)
        return float(value)
    except Exception:
        return None


def build_bounding_box(values: tuple[float, ...], source: str) -> BoundingBox | None:
    """
    功能: 将 CATIA 返回的 6 个包围盒数值转换为 BoundingBox。
    输入: 原始数值和来源。
    输出: BoundingBox 或 None。
    """
    if len(values) < 6:
        return None
    first = as_vector(values[0:3], "BoundingBox Min")
    second = as_vector(values[3:6], "BoundingBox Max")
    min_point = tuple(min(a, b) for a, b in zip(first, second))
    max_point = tuple(max(a, b) for a, b in zip(first, second))
    size = tuple(max_point[index] - min_point[index] for index in range(3))
    diagonal = vector_length(size)
    volume = max(size[0], 0.0) * max(size[1], 0.0) * max(size[2], 0.0)
    return BoundingBox(min_point, max_point, size, diagonal, volume, source)


def evaluate_product_bounding_box(document: Any, product: Any) -> tuple[BoundingBox | None, list[str]]:
    """
    功能: 尝试读取 Product 包围盒。
    输入: Document 和 Product。
    输出: 包围盒和告警列表。
    """
    attempts = (
        (
            "Product.Analyze.GetBoundingBox",
            """
Public Function read_product_bbox(product)
    Dim values(5)
    product.Analyze.GetBoundingBox values
    read_product_bbox = values
End Function
""",
            "read_product_bbox",
        ),
        (
            "Inertia.GetBoundingBox",
            """
Public Function read_product_bbox(product)
    Dim inertia
    Set inertia = product.GetTechnologicalObject("Inertia")
    Dim values(5)
    inertia.GetBoundingBox values
    read_product_bbox = values
End Function
""",
            "read_product_bbox",
        ),
        (
            "MeasureInertia.GetBoundingBox",
            """
Public Function read_product_bbox(product)
    Dim inertia
    Set inertia = product.GetTechnologicalObject("MeasureInertia")
    Dim values(5)
    inertia.GetBoundingBox values
    read_product_bbox = values
End Function
""",
            "read_product_bbox",
        ),
    )
    warnings: list[str] = []
    for source, vba_code, function_name in attempts:
        try:
            values = evaluate_product_array(document, product, vba_code, function_name, source)
            bbox = build_bounding_box(values, source)
            if bbox is not None and bbox.diagonal > 1e-9:
                return bbox, warnings
        except Exception as exc:
            warnings.append(str(exc))
    return None, warnings


def product_document_from_root_product(root_product: Any) -> Any | None:
    """
    功能: 从根 Product 获取 ProductDocument。
    输入: 根 Product。
    输出: ProductDocument 或 None。
    """
    try:
        return root_product.Parent
    except Exception:
        return None


def same_com_reference(left: Any, right: Any) -> bool:
    """
    功能: 粗略判断两个 COM 包装对象是否指向同一对象。
    输入: 两个对象。
    输出: True 或 False。
    """
    if left is None or right is None:
        return False
    if left is right:
        return True
    try:
        return left == right
    except Exception:
        return False


def resolve_product_for_part_or_product(target_part: Any, root_product: Any | None) -> Any | None:
    """
    功能: 将输入的 Part 或 Product 解析成装配树中的 Product 实例。
    输入: 目标 Part/Product 和可选根 Product。
    输出: Product 实例或 None。
    """
    if target_part is None:
        return None
    try:
        getattr(target_part, "Products")
        return target_part
    except Exception:
        pass
    try:
        getattr(target_part, "PartNumber")
        getattr(target_part, "ReferenceProduct")
        return target_part
    except Exception:
        pass
    if root_product is None:
        return None

    def visit(product: Any) -> Any | None:
        try:
            part, _document = get_part_and_document_from_product(product)
            if same_com_reference(part, target_part):
                return product
        except Exception:
            pass
        for child in get_child_products(product):
            found = visit(child)
            if found is not None:
                return found
        return None

    return visit(root_product)


def normalize_axis_direction(direction: str | Iterable[float]) -> Vector | None:
    """
    功能: 将坐标轴字符串或向量归一化为方向向量。
    输入: "X"/"-X"/"Y"/"-Y"/"Z"/"-Z" 或三维向量。
    输出: 单位方向向量或 None。
    """
    if isinstance(direction, str):
        key = direction.strip().upper().replace(" ", "")
        mapping = {
            "X": (1.0, 0.0, 0.0),
            "+X": (1.0, 0.0, 0.0),
            "-X": (-1.0, 0.0, 0.0),
            "Y": (0.0, 1.0, 0.0),
            "+Y": (0.0, 1.0, 0.0),
            "-Y": (0.0, -1.0, 0.0),
            "Z": (0.0, 0.0, 1.0),
            "+Z": (0.0, 0.0, 1.0),
            "-Z": (0.0, 0.0, -1.0),
        }
        return mapping.get(key)
    try:
        return normalize_vector(as_vector(tuple(direction), "方向"))
    except Exception:
        return None


def bounding_box_from_axis_extreme_tuple(
    extremes: tuple[Vector, Vector, Vector, Vector, Vector, Vector] | None,
    source: str,
) -> BoundingBox | None:
    """
    功能: 将六方向极值点元组转换为 BoundingBox。
    输入: 极值点元组和来源说明。
    输出: BoundingBox 或 None。
    """
    if extremes is None:
        return None
    min_x = extremes[0][0]
    max_x = extremes[1][0]
    min_y = extremes[2][1]
    max_y = extremes[3][1]
    min_z = extremes[4][2]
    max_z = extremes[5][2]
    return build_bounding_box(
        (min_x, min_y, min_z, max_x, max_y, max_z),
        source,
    )


def evaluate_wheelhouse_extreme_bounding_boxes(
    root_product: Any,
    front_component: Any,
    rear_component: Any | None = None,
    offset: float = 5000.0,
    plane_size: float = 2000.0,
) -> dict[str, Any]:
    """
    功能: 测量前后轮罩基于六方向极值点的包围盒。
    输入: 根 Product、前轮罩 Product、后轮罩 Product、参考平面距离和边长。
    输出: 前后轮罩包围盒结果字典。
    """
    results: dict[str, Any] = {}
    if isinstance(front_component, dict):
        component_items = list(front_component.items())
    else:
        component_items = [
            (FRONT_WHEELHOUSE_LABEL, front_component),
            (REAR_WHEELHOUSE_LABEL, rear_component),
        ]
    for label, component in component_items:
        if component is None:
            continue
        bbox_result = get_product_bounding_box_by_far_plane(
            component,
            root_product=root_product,
            offset=offset,
            plane_size=plane_size,
        )
        bbox = bbox_result.get("bbox")
        extremes = bbox_result.get("extreme_points")
        results[label] = {
            "status": "success" if bbox is not None else "failed",
            "component": product_display_name(component),
            "part_number": product_part_number(component),
            "extreme_points": [round_vector(point) for point in extremes] if extremes else None,
            "bbox_world": bounding_box_to_dict(bbox),
            "direction_results": bbox_result.get("direction_results"),
            "measured_collections": bbox_result.get("measured_collections"),
            "visible_only": bbox_result.get("visible_only"),
            "visibility_skipped": bbox_result.get("visibility_skipped"),
            "message": bbox_result.get("message"),
        }
        if bbox is None:
            print(f"{label} 极值包围盒测量失败")
        else:
            print(
                f"{label} 极值包围盒: min={round_vector(bbox.min_point)} "
                f"max={round_vector(bbox.max_point)} size={round_vector(bbox.size)}"
            )
    return results


def get_product_bounding_box_by_far_plane(
    target_part: Any,
    root_product: Any | None = None,
    offset: float = 3000.0,
    plane_size: float = 2000.0,
    collection_names: tuple[str, ...] = ("Bodies",),
    visible_only: bool = True,
) -> dict[str, Any]:
    """
    功能: 远平面极值法包围盒主函数，内部完成临时平面、装配测距、极值和包围盒计算。
    输入: 目标 Part/Product、可选装配根 Product、参考平面参数、参与测量集合和是否只测可见对象。
    输出: 成功返回包含极值元组和 BoundingBox 的字典，失败返回 status=failed。
    """
    visibility_skipped: list[dict[str, Any]] = []

    def local_plane_basis(normal: Vector) -> tuple[Vector, Vector] | None:
        """根据法向量计算远平面内两个正交方向。"""
        helper = (0.0, 0.0, 1.0)
        if abs(dot_product(normal, helper)) > 0.9:
            helper = (0.0, 1.0, 0.0)
        try:
            u_direction = normalize_vector(cross_product(helper, normal))
            v_direction = normalize_vector(cross_product(normal, u_direction))
        except Exception:
            return None
        return u_direction, v_direction

    def local_delete_product(product_document: Any, product: Any) -> None:
        """删除远平面临时 Product。"""
        if product is None:
            return
        try:
            selection = product_document.Selection
            selection.Clear()
            selection.Add(product)
            selection.Delete()
            selection.Clear()
            return
        except Exception:
            try:
                product_document.Selection.Clear()
            except Exception:
                pass
        try:
            product_document.Product.Products.Remove(product)
        except Exception:
            try:
                product_document.Product.Products.Remove(product_display_name(product))
            except Exception:
                pass

    def local_create_far_plane(
        product_document: Any,
        root: Any,
        normal: Vector,
        target_center: Vector,
    ) -> dict[str, Any] | None:
        """创建一个沿 normal 偏移的临时远平面 CATPart。"""
        plane_product = None
        try:
            plane_product = add_new_part_to_product(
                root,
                f"Extreme_Far_Plane_{build_timestamp()}",
                "__EXTREME_FAR_PLANE_PRODUCT__",
            )
            plane_part, plane_document = get_part_and_document_from_product(plane_product)
            factory = plane_part.HybridShapeFactory
            hybrid_body = plane_part.HybridBodies.Add()
            set_if_possible(hybrid_body, "Name", "__EXTREME_FAR_PLANE_GEOMETRY__")
            basis = local_plane_basis(normal)
            if basis is None:
                raise RuntimeError("无法计算远平面基向量")
            u_direction, v_direction = basis
            center = add_vectors(
                target_center,
                tuple(float(offset) * value for value in normal),
            )
            half_size = float(plane_size) / 2.0
            corners = [
                add_vectors(add_vectors(center, tuple(-half_size * value for value in u_direction)), tuple(-half_size * value for value in v_direction)),
                add_vectors(add_vectors(center, tuple(-half_size * value for value in u_direction)), tuple(half_size * value for value in v_direction)),
                add_vectors(add_vectors(center, tuple(half_size * value for value in u_direction)), tuple(half_size * value for value in v_direction)),
                add_vectors(add_vectors(center, tuple(half_size * value for value in u_direction)), tuple(-half_size * value for value in v_direction)),
            ]
            points = []
            for index, corner in enumerate(corners, start=1):
                point = factory.AddNewPointCoord(*corner)
                set_if_possible(point, "Name", f"ExtremePlane_P{index}")
                hybrid_body.AppendHybridShape(point)
                points.append(point)
            try:
                plane_part.Update()
            except Exception:
                pass

            lines = []
            for index, (start_point, end_point) in enumerate(zip(points, points[1:] + points[:1]), start=1):
                line = factory.AddNewLinePtPt(
                    create_reference(plane_part, start_point),
                    create_reference(plane_part, end_point),
                )
                set_if_possible(line, "Name", f"ExtremePlane_Edge_{index}")
                hybrid_body.AppendHybridShape(line)
                lines.append(line)
            try:
                plane_part.Update()
            except Exception:
                pass

            fill = factory.AddNewFill()
            set_if_possible(fill, "Name", "__EXTREME_FAR_PLANE_SURFACE__")
            for line in lines:
                fill.AddBound(create_reference(plane_part, line))
            hybrid_body.AppendHybridShape(fill)
            try:
                plane_part.UpdateObject(fill)
            except Exception:
                plane_part.Update()
            return {
                "plane_product": plane_product,
                "plane_part": plane_part,
                "plane_document": plane_document,
                "plane_feature": fill,
                "plane_center": center,
            }
        except Exception as exc:
            print(f"极值点计算失败: 创建远处参考平面失败: {exc}")
            if plane_product is not None:
                local_delete_product(product_document, plane_product)
            return None

    def local_context_points(product_document: Any, measurable: Any, reference2: Any, instance2: Any) -> tuple[float, ...] | None:
        """读取装配上下文最小距离点。"""
        try:
            vba_code = """
Public Function read_context_minimum_distance_points(measurable, ref2, instance2)
    Dim values(8)
    measurable.GetMinimumDistancePointsInContext ref2, instance2, values
    read_context_minimum_distance_points = values
End Function
"""
            values = product_document.Application.SystemService.Evaluate(
                vba_code,
                0,
                "read_context_minimum_distance_points",
                [measurable, reference2, instance2],
            )
            return tuple(float(value) for value in values)
        except Exception:
            pass
        try:
            coords = [0.0] * 9
            measurable.GetMinimumDistancePointsInContext(reference2, instance2, coords)
            return tuple(float(value) for value in coords)
        except Exception:
            return None

    def local_measure_feature_in_context(
        product_document: Any,
        feature1: Any,
        part1: Any,
        instance1: Any,
        feature2: Any,
        part2: Any,
        instance2: Any,
    ) -> dict[str, Any] | None:
        """装配上下文测两个特征距离。"""
        try:
            ref1 = create_reference(part1, feature1)
            ref2 = create_reference(part2, feature2)
            spa = product_document.GetWorkbench("SPAWorkbench")
            measurable = spa.GetMeasurableInContext(ref1, instance1)
            distance_value = float(measurable.GetMinimumDistanceInContext(ref2, instance2))
            points = local_context_points(product_document, measurable, ref2, instance2)
            result = {
                "distance": distance_value,
                "feature1_name": safe_attr_text(feature1, "Name"),
                "feature2_name": safe_attr_text(feature2, "Name"),
                "leaf_component_name": product_display_name(instance1),
                "leaf_component_part_number": product_part_number(instance1),
                "leaf_component_path": product_display_name(instance1),
                "method": "GetMeasurableInContext",
            }
            if points is not None and len(points) >= 6:
                result["point_on_element1"] = as_vector(points[0:3], "element1 最近点")
                result["point_on_element2"] = as_vector(points[3:6], "element2 最近点")
            return result
        except Exception:
            return None

    def local_leaf_products(product: Any) -> Iterable[Any]:
        """遍历 Product 下所有叶子 Product。"""
        children = get_child_products(product)
        if not children:
            yield product
            return
        for child in children:
            yield from local_leaf_products(child)

    def local_measure_product_to_feature(
        product_document: Any,
        product: Any,
        feature: Any,
        feature_product: Any,
        feature_part: Any,
    ) -> dict[str, Any] | None:
        """展开目标 Product 的可测叶子，与远平面特征测距。"""
        def local_visibility_row(target: Any, role: str) -> dict[str, Any]:
            row = get_object_visibility(product_document, target)
            row["role"] = role
            row["name"] = safe_attr_text(target, "Name", "<unnamed>")
            return row

        def local_is_visible(target: Any, role: str) -> bool:
            if not visible_only:
                return True
            visibility_row = local_visibility_row(target, role)
            if visibility_row.get("status") == "hidden" or visibility_row.get("visible") is False:
                visibility_skipped.append(visibility_row)
                return False
            return True

        best: dict[str, Any] | None = None
        for leaf_product in local_leaf_products(product):
            try:
                leaf_part, _leaf_document = get_part_and_document_from_product(leaf_product)
            except Exception:
                continue
            leaf_features: list[Any] = []
            for collection_name in collection_names:
                collection = child_collection(leaf_part, collection_name)
                if collection is None:
                    continue
                for child in iter_collection(collection):
                    if not local_is_visible(child, collection_name):
                        continue
                    leaf_features.extend(collect_measurable_leaves(leaf_part, child))
            for leaf_feature in leaf_features:
                if not local_is_visible(leaf_feature, "measurable_leaf"):
                    continue
                result = local_measure_feature_in_context(
                    product_document,
                    leaf_feature,
                    leaf_part,
                    leaf_product,
                    feature,
                    feature_part,
                    feature_product,
                )
                if result is None:
                    continue
                result["leaf_component_name"] = product_display_name(leaf_product)
                result["leaf_component_part_number"] = product_part_number(leaf_product)
                result["leaf_component_path"] = product_display_name(leaf_product)
                if best is None or float(result["distance"]) < float(best["distance"]):
                    best = result
        return best

    def local_extreme_from_distance(normal: Vector, distance_value: float) -> Vector:
        """当最近点无效时，根据远平面距离反推轴向极值投影点。"""
        plane_projection = dot_product(target_center, normal) + float(offset)
        projection = plane_projection - float(distance_value)
        offset_along_normal = tuple((projection - dot_product(target_center, normal)) * value for value in normal)
        return add_vectors(target_center, offset_along_normal)

    def local_point_is_near_origin(point: Vector | None, tolerance: float = 1.0e-9) -> bool:
        """判断点是否为空或接近原点。"""
        return point is None or vector_length(point) <= tolerance

    def local_target_center(product_document: Any, target_product: Any) -> tuple[Vector, str]:
        """优先读取目标 Product 重心作为远平面参考中心，失败时返回装配原点。"""
        try:
            center = evaluate_product_cog(
                product_document,
                target_product,
                product_display_name(target_product),
            )
            return center, "product_cog"
        except Exception as exc:
            print(f"[警告] 目标中心点读取失败，使用装配原点: {product_display_name(target_product)}: {exc}")
            return (0.0, 0.0, 0.0), "assembly_origin_fallback"

    def local_find_extreme(
        product_document: Any,
        root: Any,
        target_product: Any,
        target_center: Vector,
        direction: str,
    ) -> dict[str, Any]:
        """计算单个方向的极值点。"""
        normal = normalize_axis_direction(direction)
        if normal is None:
            return {"status": "failed", "direction": direction, "message": "方向无效"}
        plane_info = local_create_far_plane(product_document, root, normal, target_center)
        if plane_info is None:
            return {"status": "failed", "direction": direction, "message": "远平面创建失败"}
        try:
            measure_result = local_measure_product_to_feature(
                product_document,
                target_product,
                plane_info["plane_feature"],
                plane_info["plane_product"],
                plane_info["plane_part"],
            )
            if measure_result is None:
                return {"status": "failed", "direction": direction, "message": "远平面测距失败"}
            distance_value = float(measure_result.get("distance", 0.0))
            extreme_point = measure_result.get("point_on_element1")
            point_source = "GetMinimumDistancePointsInContext"
            if local_point_is_near_origin(extreme_point) and abs(distance_value) > 1.0e-6:
                extreme_point = local_extreme_from_distance(normal, distance_value)
                point_source = "distance_projection_fallback"
            if extreme_point is None:
                return {"status": "failed", "direction": direction, "message": "未取得极值点"}
            print(
                f"极值点: direction={direction} point={round_vector(extreme_point)} "
                f"distance={distance_value:.6f} "
                f"feature={measure_result.get('feature1_name')} "
                f"part_number={measure_result.get('leaf_component_part_number')} "
                f"component={measure_result.get('leaf_component_name')} "
                f"source={point_source}"
            )
            return {
                "status": "success",
                "direction": direction,
                "point": extreme_point,
                "distance": distance_value,
                "feature_name": measure_result.get("feature1_name"),
                "leaf_component_name": measure_result.get("leaf_component_name"),
                "leaf_component_part_number": measure_result.get("leaf_component_part_number"),
                "leaf_component_path": measure_result.get("leaf_component_path"),
                "source": point_source,
                "plane_center": round_vector(plane_info.get("plane_center")),
            }
        finally:
            local_delete_product(product_document, plane_info.get("plane_product"))

    target_product = resolve_product_for_part_or_product(target_part, root_product)
    if target_product is None:
        return {
            "status": "failed",
            "message": f"无法解析目标 Product: {safe_attr_text(target_part, 'Name', '<unknown>')}",
        }
    product_document = product_document_from_root_product(root_product) if root_product is not None else None
    if product_document is None:
        try:
            product_document = target_product.Parent
        except Exception:
            product_document = None
    if product_document is None or root_product is None:
        return {"status": "failed", "message": "无法获取 ProductDocument 或根 Product"}

    target_center, target_center_source = local_target_center(product_document, target_product)
    direction_results: list[dict[str, Any]] = []
    extreme_points: list[Vector] = []
    for direction in ("-X", "X", "-Y", "Y", "-Z", "Z"):
        direction_result = local_find_extreme(
            product_document,
            root_product,
            target_product,
            target_center,
            direction,
        )
        direction_results.append(direction_result)
        if direction_result.get("status") != "success":
            return {
                "status": "failed",
                "message": f"{direction} 方向极值计算失败: {direction_result.get('message')}",
                "direction_results": direction_results,
                "measured_collections": list(collection_names),
                "visible_only": visible_only,
                "visibility_skipped": visibility_skipped,
            }
        extreme_points.append(direction_result["point"])

    extremes = tuple(extreme_points)  # type: ignore[assignment]
    bbox = bounding_box_from_axis_extreme_tuple(extremes, "far_plane_axis_extreme")
    if bbox is None:
        return {
            "status": "failed",
            "message": "极值点转换包围盒失败",
            "extreme_points": extremes,
            "direction_results": direction_results,
            "measured_collections": list(collection_names),
            "visible_only": visible_only,
            "visibility_skipped": visibility_skipped,
        }
    return {
        "status": "success",
        "message": "远平面极值法包围盒计算成功",
        "extreme_points": extremes,
        "bbox": bbox,
        "bbox_dict": bounding_box_to_dict(bbox),
        "direction_results": direction_results,
        "target_center": round_vector(target_center),
        "target_center_source": target_center_source,
        "measured_collections": list(collection_names),
        "visible_only": visible_only,
        "visibility_skipped": visibility_skipped,
    }


def get_measurable(document: Any, part: Any, feature: Any) -> Any:
    """
    功能: 获取特征的 SPA Measurable。
    输入: Document、Part 和特征。
    输出: Measurable 对象。
    """
    spa_workbench = document.GetWorkbench("SPAWorkbench")
    return spa_workbench.GetMeasurable(create_reference(part, feature))


def get_feature_direction(document: Any, part: Any, feature: Any) -> Vector:
    """
    功能: 读取线性特征方向。
    输入: Document、Part 和特征。
    输出: 归一化轴方向。
    """
    measurable = get_measurable(document, part, feature)
    values = evaluate_measurable_array(
        document,
        measurable,
        "GetDirection",
        3,
        safe_attr_text(feature, "Name"),
    )
    return normalized_axis_direction(as_vector(values, "轴方向"))


def get_feature_axis_point(document: Any, part: Any, feature: Any) -> Vector:
    """
    功能: 读取轴线特征上的一个点。
    输入: Document、Part 和特征。
    输出: 局部坐标点。
    """
    measurable = get_measurable(document, part, feature)
    label = safe_attr_text(feature, "Name")
    for method_name, item_count, slice_start in (
        ("GetPointsOnCurve", 9, 3),
        ("GetCOG", 3, 0),
        ("GetPoint", 3, 0),
    ):
        try:
            values = evaluate_measurable_array(document, measurable, method_name, item_count, label)
            return as_vector(values[slice_start : slice_start + 3], "轴线上一点")
        except Exception:
            pass
    raise RuntimeError(f"无法读取轴线上一点: {label}")


def evaluate_measurable_scalar(measurable: Any, method_or_property_name: str) -> float:
    """
    功能: 读取 Measurable 的标量属性或无参方法。
    输入: Measurable 对象和属性/方法名。
    输出: 浮点值。
    """
    value = getattr(measurable, method_or_property_name)
    if callable(value):
        value = value()
    return float(value)


def get_circle_radius(measurable: Any) -> float:
    """
    功能: 读取圆线或圆弧的半径。
    输入: Measurable 对象。
    输出: 半径。
    """
    for name in ("Radius", "GetRadius"):
        try:
            radius = evaluate_measurable_scalar(measurable, name)
            if radius > 1e-6:
                return radius
        except Exception:
            pass
    raise RuntimeError("无法读取圆半径。")


def get_circle_center(document: Any, measurable: Any, label: str) -> Vector:
    """
    功能: 读取圆线或圆弧中心点。
    输入: Document、Measurable 和标签。
    输出: 局部坐标圆心。
    """
    for method_name in ("GetCenter", "GetCOG", "GetPoint"):
        try:
            values = evaluate_measurable_array(document, measurable, method_name, 3, label)
            return as_vector(values, "圆心")
        except Exception:
            pass
    raise RuntimeError(f"无法读取圆心: {label}")


def get_circle_axis_direction(document: Any, measurable: Any, label: str) -> Vector:
    """
    功能: 读取圆线或圆弧所在平面的法向作为轴方向。
    输入: Document、Measurable 和标签。
    输出: 局部坐标轴方向。
    """
    for method_name in ("GetAxis", "GetDirection"):
        try:
            values = evaluate_measurable_array(document, measurable, method_name, 3, label)
            return normalized_axis_direction(as_vector(values, "圆轴方向"))
        except Exception:
            pass
    try:
        values = evaluate_measurable_array(document, measurable, "GetPlane", 9, label)
        first_direction = as_vector(values[3:6], "圆平面第一方向")
        second_direction = as_vector(values[6:9], "圆平面第二方向")
        return normalized_axis_direction(cross_product(first_direction, second_direction))
    except Exception as exc:
        raise RuntimeError(f"无法读取圆轴方向: {label}") from exc


def iter_topology_circle_records(context: WheelPartContext) -> Iterable[TopologyCircleRecord]:
    """
    功能: 从零件几何图形集中提取可测圆线或圆弧拓扑记录。
    输入: 叶子零件上下文。
    输出: TopologyCircleRecord 迭代器。
    """
    seen_feature_names: set[str] = set()
    for hybrid_body in iter_hybrid_bodies(context.part):
        for feature in iter_hybrid_shapes(hybrid_body):
            feature_name = safe_attr_text(feature, "Name", "Circle")
            if feature_name in seen_feature_names:
                continue
            try:
                measurable = get_measurable(context.document, context.part, feature)
                radius = get_circle_radius(measurable)
                center = get_circle_center(context.document, measurable, feature_name)
                direction = get_circle_axis_direction(context.document, measurable, feature_name)
            except Exception:
                continue
            seen_feature_names.add(feature_name)
            yield TopologyCircleRecord(feature_name, center, direction, radius)


def read_axis_system_directions(document: Any, axis_system: Any) -> list[Vector]:
    """
    功能: 读取 AxisSystem 的三个方向。
    输入: Document 和 AxisSystem。
    输出: X/Y/Z 方向列表。
    """
    vba_code = """
Public Function read_axis_system_vectors(axisSystem)
    Dim xVector(2)
    Dim yVector(2)
    axisSystem.GetVectors xVector, yVector
    Dim values(5)
    values(0) = xVector(0)
    values(1) = xVector(1)
    values(2) = xVector(2)
    values(3) = yVector(0)
    values(4) = yVector(1)
    values(5) = yVector(2)
    read_axis_system_vectors = values
End Function
"""
    try:
        values = document.Application.SystemService.Evaluate(
            vba_code,
            0,
            "read_axis_system_vectors",
            [axis_system],
        )
        x_direction = normalized_axis_direction(as_vector(values[0:3], "轴系 X 方向"))
        y_direction = normalized_axis_direction(as_vector(values[3:6], "轴系 Y 方向"))
        z_direction = normalized_axis_direction(cross_product(x_direction, y_direction))
        return [x_direction, y_direction, z_direction]
    except Exception:
        return []


def read_axis_system_origin(document: Any, axis_system: Any) -> Vector | None:
    """
    功能: 读取 AxisSystem 原点。
    输入: Document 和 AxisSystem。
    输出: 原点坐标或 None。
    """
    vba_code = """
Public Function read_axis_system_origin(axisSystem)
    Dim values(2)
    axisSystem.GetOrigin values
    read_axis_system_origin = values
End Function
"""
    try:
        values = document.Application.SystemService.Evaluate(
            vba_code,
            0,
            "read_axis_system_origin",
            [axis_system],
        )
        return as_vector(values, "轴系原点")
    except Exception:
        return None


def as_vector(values: Any, label: str) -> Vector:
    """
    功能: 将序列转换为三维向量。
    输入: 数值序列和标签。
    输出: Vector。
    """
    try:
        items = tuple(float(value) for value in values)
    except Exception as exc:
        raise TypeError(f"无法读取{label}") from exc
    if len(items) < 3:
        raise ValueError(f"{label}数量不足: {items}")
    return items[0], items[1], items[2]


def vector_length(vector: Vector) -> float:
    """
    功能: 计算向量长度。
    输入: 三维向量。
    输出: 模长。
    """
    return math.sqrt(sum(component * component for component in vector))


def normalize_vector(vector: Vector) -> Vector:
    """
    功能: 归一化向量。
    输入: 三维向量。
    输出: 单位向量。
    """
    length = vector_length(vector)
    if length <= 1e-9:
        raise ValueError(f"零长度方向向量: {vector}")
    return vector[0] / length, vector[1] / length, vector[2] / length


def normalized_axis_direction(vector: Vector) -> Vector:
    """
    功能: 按正反同轴规则归一化轴方向。
    输入: 三维方向向量。
    输出: 稳定的单位方向向量。
    """
    direction = normalize_vector(vector)
    for component in direction:
        if abs(component) > 1e-9:
            if component < 0:
                return -direction[0], -direction[1], -direction[2]
            return direction
    return direction


def cross_product(first: Vector, second: Vector) -> Vector:
    """
    功能: 计算叉乘。
    输入: 两个三维向量。
    输出: 叉乘向量。
    """
    return (
        first[1] * second[2] - first[2] * second[1],
        first[2] * second[0] - first[0] * second[2],
        first[0] * second[1] - first[1] * second[0],
    )


def dot_product(first: Vector, second: Vector) -> float:
    """
    功能: 计算点乘。
    输入: 两个三维向量。
    输出: 点乘值。
    """
    return sum(a * b for a, b in zip(first, second))


def subtract_vectors(first: Vector, second: Vector) -> Vector:
    """
    功能: 向量相减。
    输入: 两个三维向量。
    输出: first - second。
    """
    return first[0] - second[0], first[1] - second[1], first[2] - second[2]


def add_vectors(first: Vector, second: Vector) -> Vector:
    """
    功能: 向量相加。
    输入: 两个三维向量。
    输出: first + second。
    """
    return first[0] + second[0], first[1] + second[1], first[2] + second[2]


def scale_vector(vector: Vector, scale: float) -> Vector:
    """
    功能: 向量数乘。
    输入: 三维向量和缩放系数。
    输出: 缩放后的三维向量。
    """
    return vector[0] * scale, vector[1] * scale, vector[2] * scale


def rotate_vector_around_axis(vector: Vector, axis_direction: Vector, angle_degrees: float) -> Vector:
    """
    功能: 使用罗德里格旋转公式将向量绕指定轴旋转。
    输入: 三维向量、旋转轴方向和角度。
    输出: 旋转后的三维向量。
    """
    axis = normalize_vector(axis_direction)
    angle = math.radians(angle_degrees)
    cos_value = math.cos(angle)
    sin_value = math.sin(angle)
    cross = cross_product(axis, vector)
    dot = dot_product(axis, vector)
    return (
        vector[0] * cos_value + cross[0] * sin_value + axis[0] * dot * (1.0 - cos_value),
        vector[1] * cos_value + cross[1] * sin_value + axis[1] * dot * (1.0 - cos_value),
        vector[2] * cos_value + cross[2] * sin_value + axis[2] * dot * (1.0 - cos_value),
    )


def plane_equation_from_points(first: Vector, second: Vector, third: Vector) -> tuple[float, float, float, float]:
    """
    功能: 根据三个点计算平面方程。
    输入: 三个不共线点。
    输出: A*x + B*y + C*z = D 的 A、B、C、D。
    """
    first_vector = subtract_vectors(second, first)
    second_vector = subtract_vectors(third, first)
    normal = normalize_vector(cross_product(first_vector, second_vector))
    d_value = dot_product(normal, first)
    return normal[0], normal[1], normal[2], d_value


def average_vectors(vectors: list[Vector]) -> Vector:
    """
    功能: 计算三维向量平均值。
    输入: 三维向量列表。
    输出: 平均向量。
    """
    if not vectors:
        raise ValueError("向量列表为空。")
    count = float(len(vectors))
    return (
        sum(vector[0] for vector in vectors) / count,
        sum(vector[1] for vector in vectors) / count,
        sum(vector[2] for vector in vectors) / count,
    )


def directions_match(
    first: Vector,
    second: Vector,
    tolerance_degrees: float = AXIS_DIRECTION_TOLERANCE_DEGREES,
) -> bool:
    """
    功能: 判断两个轴方向是否在角度容差内平行。
    输入: 两个方向和角度容差。
    输出: 是否匹配。
    """
    first_normalized = normalized_axis_direction(first)
    second_normalized = normalized_axis_direction(second)
    dot_value = max(-1.0, min(1.0, abs(dot_product(first_normalized, second_normalized))))
    return dot_value >= math.cos(math.radians(tolerance_degrees))


def distance(first: Vector, second: Vector) -> float:
    """
    功能: 计算两点距离。
    输入: 两个三维坐标。
    输出: 欧氏距离。
    """
    return math.sqrt(sum((a - b) ** 2 for a, b in zip(first, second)))


def point_to_axis_distance(point: Vector, axis_point: Vector, axis_direction: Vector) -> float:
    """
    功能: 计算点到轴线距离。
    输入: 点、轴线上一点和轴方向。
    输出: 最短距离。
    """
    direction = normalize_vector(axis_direction)
    offset = subtract_vectors(point, axis_point)
    return vector_length(cross_product(offset, direction))


def project_point_to_axis(point: Vector, axis_point: Vector, axis_direction: Vector) -> Vector:
    """
    功能: 将空间点投影到轴线上。
    输入: 点、轴线上一点和轴方向。
    输出: 轴线投影点。
    """
    direction = normalize_vector(axis_direction)
    parameter = dot_product(subtract_vectors(point, axis_point), direction)
    return add_vectors(axis_point, scale_vector(direction, parameter))


def axis_line_distance(
    first_point: Vector,
    first_direction: Vector,
    second_point: Vector,
    second_direction: Vector,
) -> float:
    """
    功能: 计算两条空间轴线的最短距离。
    输入: 两条轴线的点和方向。
    输出: 最短距离。
    """
    first = normalize_vector(first_direction)
    second = normalize_vector(second_direction)
    cross = cross_product(first, second)
    cross_length = vector_length(cross)
    offset = subtract_vectors(second_point, first_point)
    if cross_length <= 1e-9:
        return vector_length(cross_product(offset, first))
    return abs(dot_product(offset, cross)) / cross_length


def iter_hybrid_bodies(part: Any) -> Iterable[Any]:
    """
    功能: 遍历 Part 下所有顶层和嵌套几何图形集。
    输入: Part。
    输出: HybridBody 迭代器。
    """
    try:
        hybrid_bodies = part.HybridBodies
    except Exception:
        return
    for hybrid_body in iter_collection(hybrid_bodies):
        yield hybrid_body
        try:
            nested_bodies = hybrid_body.HybridBodies
        except Exception:
            nested_bodies = None
        if nested_bodies is not None:
            for nested_body in iter_collection(nested_bodies):
                yield from iter_hybrid_bodies_from_body(nested_body)


def iter_hybrid_bodies_from_body(hybrid_body: Any) -> Iterable[Any]:
    """
    功能: 递归遍历指定几何图形集及其子几何图形集。
    输入: HybridBody。
    输出: HybridBody 迭代器。
    """
    yield hybrid_body
    try:
        nested_bodies = hybrid_body.HybridBodies
    except Exception:
        return
    for nested_body in iter_collection(nested_bodies):
        yield from iter_hybrid_bodies_from_body(nested_body)


def iter_hybrid_shapes(hybrid_body: Any) -> Iterable[Any]:
    """
    功能: 遍历几何图形集内的几何特征。
    输入: HybridBody。
    输出: HybridShape 迭代器。
    """
    try:
        hybrid_shapes = hybrid_body.HybridShapes
    except Exception:
        return
    for hybrid_shape in iter_collection(hybrid_shapes):
        yield hybrid_shape


def find_hybrid_body_by_name(part: Any, body_name: str) -> Any | None:
    """
    功能: 按名称查找 Part 下的几何图形集。
    输入: Part 和几何图形集名称。
    输出: 找到的 HybridBody，未找到时返回 None。
    """
    for hybrid_body in iter_hybrid_bodies(part):
        if safe_attr_text(hybrid_body, "Name") == body_name:
            return hybrid_body
    return None


def find_hybrid_shape_in_body_by_name(part: Any, body_name: str, feature_name: str) -> tuple[Any, Any, dict[str, Any]]:
    """
    功能: 在指定几何图形集下按名称查找 HybridShape。
    输入: Part、几何图形集名称和特征名称。
    输出: 特征对象、几何图形集和匹配元数据。
    """
    hybrid_body = find_hybrid_body_by_name(part, body_name)
    if hybrid_body is None:
        raise RuntimeError(f"未找到几何图形集: {body_name}")
    matches: list[tuple[Any, int]] = []
    for index, hybrid_shape in enumerate(iter_hybrid_shapes(hybrid_body), start=1):
        if safe_attr_text(hybrid_shape, "Name") == feature_name:
            matches.append((hybrid_shape, index))
    if not matches:
        raise RuntimeError(f"未在几何图形集 {body_name} 中找到特征: {feature_name}")
    feature, index = matches[-1]
    return feature, hybrid_body, {
        "source_type": "hybrid_shape",
        "source_name": feature_name,
        "body_name": safe_attr_text(hybrid_body, "Name"),
        "selected_index": index,
        "matched_count": len(matches),
    }


def get_or_create_hybrid_body(part: Any, body_name: str) -> Any:
    """
    功能: 获取或创建 Part 下的几何图形集。
    输入: Part 和几何图形集名称。
    输出: HybridBody 对象。
    """
    existing = find_hybrid_body_by_name(part, body_name)
    if existing is not None:
        return existing
    hybrid_bodies = part.HybridBodies
    hybrid_body = hybrid_bodies.Add()
    set_if_possible(hybrid_body, "Name", body_name)
    return hybrid_body


def append_update_hybrid_shape(part: Any, hybrid_body: Any, feature: Any, feature_name: str) -> None:
    """
    功能: 将 HybridShape 添加到几何图形集并更新。
    输入: Part、目标几何图形集、特征对象和特征名称。
    输出: 无；更新失败时抛出异常。
    """
    set_if_possible(feature, "Name", feature_name)
    hybrid_body.AppendHybridShape(feature)
    try:
        part.InWorkObject = feature
    except Exception:
        pass
    update_object = getattr(part, "UpdateObject", None)
    if callable(update_object):
        update_object(feature)
    else:
        part.Update()


def reference_from_hybrid_shape_name(part: Any, feature_name: str) -> tuple[Any, dict[str, Any]]:
    """
    功能: 按名称查找 HybridShape 并创建 Reference。
    输入: Part 和特征名称。
    输出: Reference 和匹配元数据。
    """
    matches: list[tuple[Any, Any, int]] = []
    for hybrid_body in iter_hybrid_bodies(part):
        for index, hybrid_shape in enumerate(iter_hybrid_shapes(hybrid_body), start=1):
            if safe_attr_text(hybrid_shape, "Name") == feature_name:
                matches.append((hybrid_shape, hybrid_body, index))
    if not matches:
        raise RuntimeError(f"未找到几何特征: {feature_name}")
    feature, hybrid_body, index = matches[-1]
    return create_reference(part, feature), {
        "source_type": "hybrid_shape",
        "source_name": feature_name,
        "body_name": safe_attr_text(hybrid_body, "Name"),
        "selected_index": index,
        "matched_count": len(matches),
    }


def find_hybrid_shape_object_by_name(part: Any, feature_name: str) -> tuple[Any, Any, dict[str, Any]]:
    """
    功能: 按名称查找 HybridShape 对象，优先返回最后创建的同名对象。
    输入: Part 和特征名称。
    输出: 特征对象、所在几何图形集和匹配元数据。
    """
    matches: list[tuple[Any, Any, int]] = []
    for hybrid_body in iter_hybrid_bodies(part):
        for index, hybrid_shape in enumerate(iter_hybrid_shapes(hybrid_body), start=1):
            if safe_attr_text(hybrid_shape, "Name") == feature_name:
                matches.append((hybrid_shape, hybrid_body, index))
    if not matches:
        raise RuntimeError(f"未找到几何特征对象: {feature_name}")
    feature, hybrid_body, index = matches[-1]
    return feature, hybrid_body, {
        "source_type": "hybrid_shape",
        "source_name": feature_name,
        "body_name": safe_attr_text(hybrid_body, "Name"),
        "selected_index": index,
        "matched_count": len(matches),
    }


def copy_hybrid_shape_result_between_parts(
    source_document: Any,
    source_part: Any,
    source_feature_name: str,
    target_document: Any,
    target_part: Any,
    target_hybrid_body: Any,
    result_name: str,
) -> dict[str, Any]:
    """
    功能: 将源 Part 中成功的 HybridShape 按无链接结果复制到目标 Part 几何图形集。
    输入: 源 Document/Part/特征名、目标 Document/Part/几何图形集和结果名称。
    输出: 复制粘贴结果字典。
    """
    source_feature, source_body, source_metadata = find_hybrid_shape_object_by_name(
        source_part,
        source_feature_name,
    )
    try:
        target_before_count = collection_count(target_hybrid_body.HybridShapes)
    except Exception:
        target_before_count = 0
    paste_errors: list[dict[str, str]] = []
    source_selection = source_document.Selection
    target_selection = target_document.Selection
    try:
        source_document.Activate()
        source_selection.Clear()
        source_selection.Add(source_feature)
        source_selection.Copy()
        source_selection.Clear()

        target_document.Activate()
        try:
            target_part.InWorkObject = target_hybrid_body
        except Exception:
            pass
        target_selection.Clear()
        target_selection.Add(target_hybrid_body)
        used_method = ""
        for paste_format in ("CATPrtResultWithOutLink", "CATPrtResult", "CATPrtCont"):
            try:
                target_selection.PasteSpecial(paste_format)
                used_method = f"PasteSpecial({paste_format})"
                break
            except Exception as exc:
                paste_errors.append({"method": f"PasteSpecial({paste_format})", "message": str(exc)})
        if not used_method:
            target_selection.Paste()
            used_method = "Paste()"
        target_selection.Clear()
        try:
            target_part.Update()
        except Exception:
            pass
    except Exception as exc:
        try:
            source_selection.Clear()
        except Exception:
            pass
        try:
            target_selection.Clear()
        except Exception:
            pass
        return {
            "status": "failed",
            "message": f"断参曲线复制到过程 Part 失败: {exc}",
            "source_feature_name": source_feature_name,
            "source_body_name": safe_attr_text(source_body, "Name"),
            "target_body_name": safe_attr_text(target_hybrid_body, "Name"),
            "paste_errors": paste_errors,
            "source_metadata": source_metadata,
        }

    renamed_features: list[str] = []
    try:
        target_shapes = list(iter_hybrid_shapes(target_hybrid_body))
        for index, pasted_feature in enumerate(target_shapes[target_before_count:], start=1):
            pasted_name = result_name if index == 1 else f"{result_name}_{index:03d}"
            set_if_possible(pasted_feature, "Name", pasted_name)
            renamed_features.append(pasted_name)
        try:
            target_part.Update()
        except Exception:
            pass
    except Exception:
        pass
    return {
        "status": "success",
        "message": "断参曲线已复制到过程 Part。",
        "source_feature_name": source_feature_name,
        "source_body_name": safe_attr_text(source_body, "Name"),
        "target_body_name": safe_attr_text(target_hybrid_body, "Name"),
        "result_names": renamed_features,
        "paste_method": used_method,
        "paste_errors": paste_errors,
        "source_metadata": source_metadata,
    }


def get_curve_sample_points(document: Any, part: Any, feature: Any) -> list[Vector]:
    """
    功能: 读取曲线的代表点，优先取起点、中点和终点。
    输入: Document、Part 和曲线特征。
    输出: 曲线局部坐标点列表。
    """
    label = safe_attr_text(feature, "Name")
    try:
        measurable = get_measurable(document, part, feature)
        values = evaluate_measurable_array(document, measurable, "GetPointsOnCurve", 9, label)
        return [
            as_vector(values[0:3], "曲线起点"),
            as_vector(values[3:6], "曲线中点"),
            as_vector(values[6:9], "曲线终点"),
        ]
    except Exception:
        pass
    for method_name in ("GetCOG", "GetPoint"):
        try:
            measurable = get_measurable(document, part, feature)
            values = evaluate_measurable_array(document, measurable, method_name, 3, label)
            return [as_vector(values, f"曲线{method_name}")]
        except Exception:
            pass
    return []


def get_feature_point_local(document: Any, part: Any, feature: Any) -> Vector | None:
    """
    功能: 读取点状特征的局部坐标。
    输入: Document、Part 和点特征。
    输出: 局部坐标点，失败时为 None。
    """
    label = safe_attr_text(feature, "Name")
    try:
        measurable = get_measurable(document, part, feature)
    except Exception:
        return None
    for method_name in ("GetPoint", "GetCOG"):
        try:
            values = evaluate_measurable_array(document, measurable, method_name, 3, label)
            return as_vector(values, f"点{method_name}")
        except Exception:
            pass
    return None


def local_direction_for_world_direction(transform: Transform, world_direction: Vector) -> Vector:
    """
    功能: 将装配世界方向转换为组件局部方向。
    输入: 组件装配 Transform 和世界方向。
    输出: 局部方向向量。
    """
    direction = normalize_vector(world_direction)
    return normalize_vector(
        (
            dot_product(direction, transform.x_axis),
            dot_product(direction, transform.y_axis),
            dot_product(direction, transform.z_axis),
        )
    )


def create_section_far_plane_surface(
    document: Any,
    part: Any,
    center_local: Vector,
    normal_local: Vector,
    name: str,
    offset: float = SECTION_EXTREME_FAR_PLANE_OFFSET,
    plane_size: float = SECTION_EXTREME_FAR_PLANE_SIZE,
) -> dict[str, Any]:
    """
    功能: 在截面 Part 内创建用于极值测距的远平面面片。
    输入: Document、Part、局部中心、局部法向、名称、偏移和尺寸。
    输出: 包含面片特征和构造特征的字典。
    """
    factory = part.HybridShapeFactory
    hybrid_body = get_or_create_hybrid_body(part, "__SECTION_EXTREME_FAR_PLANES__")
    normal = normalize_vector(normal_local)
    helper = (0.0, 0.0, 1.0)
    if abs(dot_product(normal, helper)) > 0.9:
        helper = (0.0, 1.0, 0.0)
    u_direction = normalize_vector(cross_product(helper, normal))
    v_direction = normalize_vector(cross_product(normal, u_direction))
    plane_center = add_vectors(center_local, scale_vector(normal, float(offset)))
    half_size = float(plane_size) / 2.0
    corners = [
        add_vectors(add_vectors(plane_center, scale_vector(u_direction, -half_size)), scale_vector(v_direction, -half_size)),
        add_vectors(add_vectors(plane_center, scale_vector(u_direction, -half_size)), scale_vector(v_direction, half_size)),
        add_vectors(add_vectors(plane_center, scale_vector(u_direction, half_size)), scale_vector(v_direction, half_size)),
        add_vectors(add_vectors(plane_center, scale_vector(u_direction, half_size)), scale_vector(v_direction, -half_size)),
    ]
    construction_features: list[Any] = []
    points: list[Any] = []
    for index, corner in enumerate(corners, start=1):
        point = factory.AddNewPointCoord(*corner)
        set_if_possible(point, "Name", f"{name}_P{index}")
        hybrid_body.AppendHybridShape(point)
        points.append(point)
        construction_features.append(point)
    try:
        part.Update()
    except Exception:
        pass
    lines: list[Any] = []
    for index, (start_point, end_point) in enumerate(zip(points, points[1:] + points[:1]), start=1):
        line = factory.AddNewLinePtPt(
            create_reference(part, start_point),
            create_reference(part, end_point),
        )
        set_if_possible(line, "Name", f"{name}_E{index}")
        hybrid_body.AppendHybridShape(line)
        lines.append(line)
        construction_features.append(line)
    try:
        part.Update()
    except Exception:
        pass
    fill = factory.AddNewFill()
    set_if_possible(fill, "Name", name)
    for line in lines:
        fill.AddBound(create_reference(part, line))
    hybrid_body.AppendHybridShape(fill)
    try:
        part.UpdateObject(fill)
    except Exception:
        part.Update()
    return {
        "hybrid_body": hybrid_body,
        "plane_feature": fill,
        "construction_features": construction_features,
        "center_local": plane_center,
        "normal_local": normal,
    }


def hide_section_far_plane_geometry_set(document: Any, part: Any, active_feature: Any | None = None) -> None:
    """
    功能: 隐藏截面极值远平面几何图形集。
    输入: Document、Part 和可选当前工作特征。
    输出: 无。
    """
    far_plane_body = find_hybrid_body_by_name(part, "__SECTION_EXTREME_FAR_PLANES__")
    if far_plane_body is None:
        return
    try:
        document.Activate()
    except Exception:
        pass
    try:
        if active_feature is not None:
            part.InWorkObject = active_feature
    except Exception:
        pass
    try:
        set_object_visibility(document, far_plane_body, False)
    except Exception:
        selection = document.Selection
        try:
            selection.Clear()
            selection.Add(far_plane_body)
            selection.VisProperties.SetShow(1)
        finally:
            try:
                selection.Clear()
            except Exception:
                pass
    try:
        part.Update()
    except Exception:
        pass
    try:
        document.Application.ActiveWindow.ActiveViewer.Update()
    except Exception:
        pass


def get_curve_direction_extreme_points(
    document: Any,
    part: Any,
    feature: Any,
    transform: Transform,
    world_direction: Vector,
) -> tuple[Vector | None, Vector | None, dict[str, Any]]:
    """
    功能: 获取曲线沿指定装配世界方向的最大点和最小点。
    输入: Document、Part、曲线特征、组件装配 Transform 和世界方向。
    输出: 世界坐标方向最大点、方向最小点和方法信息。
    """
    direction_world = normalize_vector(world_direction)
    direction_local = local_direction_for_world_direction(transform, direction_world)
    errors: list[str] = []
    local_points = get_curve_sample_points(document, part, feature)
    if not local_points:
        return None, None, {
            "method": "failed",
            "world_direction": round_vector(direction_world, 8),
            "candidate_count": 0,
            "errors": ["无法读取曲线采样点，无法创建远平面参考中心。"],
        }
    center_local = average_vectors(local_points)

    def measure_extreme_by_far_plane(label: str, normal_local: Vector) -> tuple[Vector | None, dict[str, Any]]:
        plane_info: dict[str, Any] | None = None
        try:
            plane_info = create_section_far_plane_surface(
                document,
                part,
                center_local,
                normal_local,
                f"_Regulation_{safe_attr_text(feature, 'Name', 'Curve')}_{label}_FarPlane",
            )
            measure_result = measure_distance_between(
                getattr(document, "Application", None),
                feature,
                plane_info["plane_feature"],
                part=part,
                part_doc=document,
                return_points=True,
            )
            if measure_result is None:
                raise RuntimeError("曲线到远平面测距失败")
            point_local = measure_result.get("point_on_element1")
            if point_local is None:
                raise RuntimeError(f"曲线到远平面未返回曲线侧最近点: {measure_result.get('points_error')}")
            return apply_transform_to_point(transform, point_local), {
                "label": label,
                "distance": round(float(measure_result.get("distance", 0.0)), 6),
                "point_source": measure_result.get("method"),
                "points_error": measure_result.get("points_error"),
                "plane_name": safe_attr_text(plane_info["plane_feature"], "Name"),
                "plane_center_local": round_vector(plane_info["center_local"]),
                "plane_normal_local": round_vector(plane_info["normal_local"], 8),
            }
        except Exception as exc:
            return None, {"label": label, "error": str(exc)}

    max_point, max_info = measure_extreme_by_far_plane("Max", direction_local)
    min_point, min_info = measure_extreme_by_far_plane("Min", scale_vector(direction_local, -1.0))
    try:
        hide_section_far_plane_geometry_set(document, part, feature)
    except Exception as exc:
        errors.append(f"hide_far_plane_body: {exc}")
    for info in (max_info, min_info):
        if info.get("error"):
            errors.append(f"{info.get('label')}: {info.get('error')}")
    if max_point is not None and min_point is not None:
        high_point = max((max_point, min_point), key=lambda point: dot_product(point, direction_world))
        low_point = min((max_point, min_point), key=lambda point: dot_product(point, direction_world))
        return high_point, low_point, {
            "method": "far_plane_minimum_distance",
            "world_direction": round_vector(direction_world, 8),
            "local_direction": round_vector(direction_local, 8),
            "candidate_count": 2,
            "max_plane": max_info,
            "min_plane": min_info,
            "errors": errors,
        }

    sample_points = [
        apply_transform_to_point(transform, point)
        for point in local_points
    ]
    return (
        max(sample_points, key=lambda point: dot_product(point, direction_world)),
        min(sample_points, key=lambda point: dot_product(point, direction_world)),
        {
            "method": "sample_fallback",
            "world_direction": round_vector(direction_world, 8),
            "local_direction": round_vector(direction_local, 8),
            "candidate_count": len(sample_points),
            "errors": errors,
        },
    )


def append_section_curve_extreme_geometry(
    document: Any | None,
    part: Any,
    hybrid_body: Any,
    feature_name: str,
    high_point_local: Vector,
    low_point_local: Vector,
) -> dict[str, Any]:
    """
    功能: 在导出的截面 CATPart 内为单条 SectionResult 曲线创建方向最大点、方向最小点和连线。
    输入: 导出 Part、SectionResult 几何集、曲线名称和该 Part 局部坐标点。
    输出: 创建结果字典。
    """
    factory = part.HybridShapeFactory
    safe_feature = safe_filename_text(feature_name)
    high_name = f"{safe_feature}_Dir_Max"
    low_name = f"{safe_feature}_Dir_Min"
    line_name = f"{safe_feature}_Dir_Distance"
    high_feature = factory.AddNewPointCoord(*high_point_local)
    low_feature = factory.AddNewPointCoord(*low_point_local)
    set_if_possible(high_feature, "Name", high_name)
    set_if_possible(low_feature, "Name", low_name)
    hybrid_body.AppendHybridShape(high_feature)
    hybrid_body.AppendHybridShape(low_feature)
    line_feature = factory.AddNewLinePtPt(
        create_reference(part, high_feature),
        create_reference(part, low_feature),
    )
    set_if_possible(line_feature, "Name", line_name)
    hybrid_body.AppendHybridShape(line_feature)
    try:
        part.UpdateObject(line_feature)
    except Exception:
        try:
            part.Update()
        except Exception:
            pass
    if document is not None:
        try:
            set_object_visibility(document, line_feature, False)
        except Exception:
            pass
    return {
        "status": "success",
        "target_part": safe_attr_text(part, "Name"),
        "target_body": safe_attr_text(hybrid_body, "Name"),
        "source_feature_name": feature_name,
        "high_point_name": high_name,
        "low_point_name": low_name,
        "distance_line_name": line_name,
        "distance_line_length": round(vector_length(subtract_vectors(high_point_local, low_point_local)), 6),
        "coordinate_source": "world_to_exported_section_local",
    }


def collect_section_result_curve_rows_from_part_document(
    document: Any,
    part: Any,
    component_name: str,
    component_part_number: str,
    extreme_world_direction: Vector,
    *,
    transform: Transform | None = None,
    section_body_name: str = "SectionResult",
    section_curve_name: str | None = None,
) -> list[dict[str, Any]]:
    """
    功能: 在已打开的截面结果 CATPart 中收集 SectionResult 曲线方向极值点并创建点线。
    输入: PartDocument、Part、组件名称、组件编号、世界判断方向和可选坐标变换。
    输出: 曲线极值行列表。
    """
    rows: list[dict[str, Any]] = []
    transform = transform or identity_transform()
    if not section_curve_name:
        raise RuntimeError("缺少导出的截面曲线名称，无法精确定位 SectionResult 中的对象。")
    try:
        feature, hybrid_body, feature_metadata = find_hybrid_shape_in_body_by_name(
            part,
            section_body_name,
            section_curve_name,
        )
    except Exception as exc:
        print(
            f"[警告] 截面曲线定位失败: {component_name} "
            f"{section_body_name}/{section_curve_name}: {exc}"
        )
        return rows

    local_points = get_curve_sample_points(document, part, feature)
    if not local_points:
        print(
            f"[警告] 截面曲线无法读取采样点: "
            f"{component_name} / {safe_attr_text(feature, 'Name')}"
        )
        return rows
    world_points = [apply_transform_to_point(transform, point) for point in local_points]
    high_point, low_point, extreme_info = get_curve_direction_extreme_points(
        document,
        part,
        feature,
        transform,
        extreme_world_direction,
    )
    if high_point is None or low_point is None:
        print(
            f"[警告] 截面曲线无法读取方向极值: "
            f"{component_name} / {safe_attr_text(feature, 'Name')}"
        )
        return rows
    created_geometry: dict[str, Any] = {}
    try:
        created_geometry = append_section_curve_extreme_geometry(
            document,
            part,
            hybrid_body,
            safe_attr_text(feature, "Name"),
            inverse_transform_point(transform, high_point),
            inverse_transform_point(transform, low_point),
        )
    except Exception as exc:
        created_geometry = {
            "status": "failed",
            "message": str(exc),
            "coordinate_source": "world_to_exported_section_local",
        }
    print(
        f"截面结果极值点: component={component_name} "
        f"body={safe_attr_text(hybrid_body, 'Name')} feature={safe_attr_text(feature, 'Name')} "
        f"matched_index={(feature_metadata or {}).get('selected_index')} "
        f"method={(extreme_info or {}).get('method')} "
        f"max={round_vector(high_point)} min={round_vector(low_point)}"
    )
    rows.append(
        {
            "component_name": component_name,
            "component_part_number": component_part_number,
            "feature_name": safe_attr_text(feature, "Name"),
            "body_name": safe_attr_text(hybrid_body, "Name"),
            "sample_points_world": world_points,
            "start_point_world": world_points[0],
            "end_point_world": world_points[-1],
            "curve_direction_max_point_world": high_point,
            "curve_direction_min_point_world": low_point,
            "curve_high_point_world": high_point,
            "curve_low_point_world": low_point,
            "extreme_world_direction": normalize_vector(extreme_world_direction),
            "extreme_info": extreme_info,
            "created_geometry": created_geometry,
        }
    )
    try:
        part.Update()
    except Exception as exc:
        print(f"[警告] 截面结果极值点线更新失败: {component_name}: {exc}")
    return rows


def collect_section_result_curve_rows(
    product_document: Any,
    component: Any,
    extreme_world_direction: Vector,
    *,
    section_body_name: str = "SectionResult",
    section_curve_name: str | None = None,
) -> list[dict[str, Any]]:
    """
    功能: 收集已装配截面结果组件中 SectionResult 几何集曲线的方向极值点。
    输入: ProductDocument、截面结果组件和世界判断方向。
    输出: 曲线极值行列表。
    """
    rows: list[dict[str, Any]] = []
    try:
        part, document = get_part_and_document_from_product(component)
        transform = product_position_transform(product_document, component)
    except Exception as exc:
        print(f"[警告] 截面结果组件无法读取: {product_display_name(component)}: {exc}")
        return rows
    rows = collect_section_result_curve_rows_from_part_document(
        document,
        part,
        product_display_name(component),
        product_part_number(component),
        extreme_world_direction,
        transform=transform,
        section_body_name=section_body_name,
        section_curve_name=section_curve_name,
    )
    save_result = save_document_if_modified(document)
    if save_result.get("status") == "failed":
        print(f"[警告] 截面结果极值点线保存失败: {save_result.get('document')}: {save_result.get('message')}")
    return rows


def group_curve_rows_by_connectivity(
    curve_rows: list[dict[str, Any]],
    tolerance: float = 2.0,
) -> list[list[dict[str, Any]]]:
    """
    功能: 按曲线端点接近关系分组截面曲线。
    输入: 曲线采样行和端点距离容差。
    输出: 连通曲线组列表。
    """
    groups: list[list[dict[str, Any]]] = []
    for row in curve_rows:
        row_endpoints = [row["start_point_world"], row["end_point_world"]]
        matched_indices: list[int] = []
        for group_index, group in enumerate(groups):
            group_endpoints: list[Vector] = []
            for group_row in group:
                group_endpoints.extend([group_row["start_point_world"], group_row["end_point_world"]])
            if any(
                vector_length(subtract_vectors(first, second)) <= tolerance
                for first in row_endpoints
                for second in group_endpoints
            ):
                matched_indices.append(group_index)
        if not matched_indices:
            groups.append([row])
            continue
        primary = matched_indices[0]
        groups[primary].append(row)
        for extra_index in reversed(matched_indices[1:]):
            groups[primary].extend(groups.pop(extra_index))
    return groups


def append_section_topology_group_geometry(
    part: Any,
    hybrid_body: Any,
    section_plane_name: str,
    group_index: int,
    high_point_local: Vector,
    low_point_local: Vector,
) -> dict[str, Any]:
    """
    功能: 在过程 Part 中为截面拓扑组创建方向最大点、方向最小点和连线。
    输入: Part、几何图形集、截面名称、组序号和过程 Part 局部坐标最大/最小点。
    输出: 创建结果字典。
    """
    factory = part.HybridShapeFactory
    safe_plane = safe_filename_text(section_plane_name)
    high_name = f"{safe_plane}_Group{group_index:03d}_Dir_Max"
    low_name = f"{safe_plane}_Group{group_index:03d}_Dir_Min"
    line_name = f"{safe_plane}_Group{group_index:03d}_Dir_Distance"
    high_feature = factory.AddNewPointCoord(*high_point_local)
    low_feature = factory.AddNewPointCoord(*low_point_local)
    set_if_possible(high_feature, "Name", high_name)
    set_if_possible(low_feature, "Name", low_name)
    hybrid_body.AppendHybridShape(high_feature)
    hybrid_body.AppendHybridShape(low_feature)
    line_feature = factory.AddNewLinePtPt(
        create_reference(part, high_feature),
        create_reference(part, low_feature),
    )
    set_if_possible(line_feature, "Name", line_name)
    hybrid_body.AppendHybridShape(line_feature)
    return {
        "high_point_name": high_name,
        "low_point_name": low_name,
        "distance_line_name": line_name,
        "distance_line_length": round(vector_length(subtract_vectors(high_point_local, low_point_local)), 6),
    }


def wheelhouse_definition_from_label(wheelhouse: Any) -> dict[str, str]:
    label = str(wheelhouse or "").strip()
    for definition in WHEELHOUSE_SLOT_DEFINITIONS.values():
        if label in {
            definition["label"],
            definition["measurement_prefix"],
            definition["component_part_number"],
        }:
            return definition
    return LEGACY_WHEELHOUSE_DEFINITIONS.get(label) or {
        "label": label or "Wheelhouse",
        "measurement_prefix": label.replace("_Wheelhouse", "").replace("_", "-") or "Wheelhouse",
        "component_part_number": label.replace("-", "_") or "Wheelhouse",
        "geometry_set_name": f"法规校核{label or '轮罩'}",
        "line_name": f"{(label or 'Wheelhouse').replace('-', '_')}_Regulation_Axis_Line",
        "section_prefix": label or "轮罩",
        "position": "unknown",
        "side": "unknown",
    }


def wheelhouse_measurement_prefix(wheelhouse: Any) -> str:
    return wheelhouse_definition_from_label(wheelhouse)["measurement_prefix"]


def wheelhouse_measurement_key(wheelhouse: Any, suffix: str) -> str:
    clean_suffix = str(suffix).strip().lstrip("-")
    return f"{wheelhouse_measurement_prefix(wheelhouse)}-{clean_suffix}"


def is_front_wheelhouse_label(wheelhouse: Any) -> bool:
    definition = wheelhouse_definition_from_label(wheelhouse)
    return definition.get("position") == "front" or str(wheelhouse or "") == FRONT_WHEELHOUSE_LABEL


def is_rear_wheelhouse_label(wheelhouse: Any) -> bool:
    definition = wheelhouse_definition_from_label(wheelhouse)
    return definition.get("position") == "rear" or str(wheelhouse or "") == REAR_WHEELHOUSE_LABEL


def wheelhouse_section_prefix(wheelhouse: Any) -> str:
    return wheelhouse_definition_from_label(wheelhouse)["section_prefix"]


def section_distance_measurement_key(wheelhouse: Any, section_plane_name: Any) -> str | None:
    """
    功能: 将法规截面距离线映射到固定结果键。
    输入: 轮罩标签和截面平面名称。
    输出: Left-Front-p、Left-Front-p30 等测量键，或 None。
    """
    section_text = str(section_plane_name or "")
    is_30 = "30" in section_text
    if not str(wheelhouse or "").strip():
        return None
    return wheelhouse_measurement_key(wheelhouse, "p30" if is_30 else "p")


def analyze_section_result_topology_groups(
    curve_rows: list[dict[str, Any]],
) -> list[dict[str, Any]]:
    """
    功能: 按轮罩和截面角度汇总截面结果曲线，并计算组内沿截面偏转方向的最终跨度。
    输入: 截面结果列表和导出截面 CATPart 曲线采样行。
    输出: 拓扑分组结果列表。
    """
    rows_by_key: dict[tuple[str, str], list[dict[str, Any]]] = {}
    metadata_by_key: dict[tuple[str, str], dict[str, Any]] = {}
    for row in curve_rows:
        key = (str(row.get("wheelhouse")), str(row.get("section_plane_name")))
        rows_by_key.setdefault(key, []).append(row)
        metadata_by_key.setdefault(
            key,
            {
                "wheelhouse": row.get("wheelhouse"),
                "section_plane_name": row.get("section_plane_name"),
                "geometry_set_name": row.get("geometry_set_name"),
            },
        )

    results: list[dict[str, Any]] = []
    for key, rows in rows_by_key.items():
        metadata = metadata_by_key[key]
        # 同一轮罩、同一角度的多个 SectionResult CATPart 属于同一法规校核组。
        # 这里不再按端点连通性拆分，避免每个相交 Part 都各自生成一对极值点。
        groups = [rows]
        for group_index, group in enumerate(groups, start=1):
            group_direction = normalize_vector(
                group[0].get("extreme_world_direction") or (0.0, 0.0, 1.0)
            )
            high_candidates = [
                group_row["curve_direction_max_point_world"]
                for group_row in group
                if group_row.get("curve_direction_max_point_world") is not None
            ]
            low_candidates = [
                group_row["curve_direction_min_point_world"]
                for group_row in group
                if group_row.get("curve_direction_min_point_world") is not None
            ]
            if not high_candidates or not low_candidates:
                continue
            high_point = max(high_candidates, key=lambda point: dot_product(point, group_direction))
            low_point = min(low_candidates, key=lambda point: dot_product(point, group_direction))
            distance = vector_length(subtract_vectors(high_point, low_point))
            projection_span = abs(dot_product(subtract_vectors(high_point, low_point), group_direction))
            z_distance = abs(float(high_point[2]) - float(low_point[2]))
            measurement_key = section_distance_measurement_key(
                metadata.get("wheelhouse"),
                metadata.get("section_plane_name"),
            )
            result = {
                "wheelhouse": metadata.get("wheelhouse"),
                "section_plane_name": metadata.get("section_plane_name"),
                "geometry_set_name": metadata.get("geometry_set_name"),
                "group_index": group_index,
                "curve_count": len(group),
                "component_part_numbers": sorted(
                    {
                        str(group_row.get("component_part_number"))
                        for group_row in group
                        if group_row.get("component_part_number")
                    }
                ),
                "extreme_methods": sorted(
                    {
                        str((group_row.get("extreme_info") or {}).get("method"))
                        for group_row in group
                        if (group_row.get("extreme_info") or {}).get("method")
                    }
                ),
                "high_point_world": round_vector(high_point),
                "low_point_world": round_vector(low_point),
                "extreme_world_direction": round_vector(group_direction, 8),
                "direction_distance": round(distance, 6),
                "direction_projection_span": round(projection_span, 6),
                "z_distance": round(z_distance, 6),
                "z_height": round(z_distance, 6),
                "measurement_key": measurement_key,
                "measurement_value": round(z_distance, 6) if measurement_key else None,
                "measurement_source": "world_z_delta_between_section_extreme_points",
                "created_geometry": {
                    "status": "skipped",
                    "message": "极值点和连线已创建在各自导出的截面 CATPart 中。",
                    "section_part_geometry": [
                        group_row.get("created_geometry")
                        for group_row in group
                        if group_row.get("created_geometry")
                    ],
                },
            }
            results.append(result)
            print(
                f"截面拓扑组: {metadata.get('section_plane_name')} "
                f"group={group_index} curves={len(group)} "
                f"distance={distance:.3f} direction_span={projection_span:.3f} "
                f"z_distance={z_distance:.3f}"
            )
    return results


def measure_feature_basic(document: Any, part: Any, feature: Any) -> dict[str, Any]:
    """
    功能: 测量 CATIA 特征的基础几何属性。
    输入: Document、Part 和特征对象。
    输出: 包含名称、面积、长度、体积和重心的字典。
    """
    result: dict[str, Any] = {"name": safe_attr_text(feature, "Name")}
    try:
        measurable = document.GetWorkbench("SPAWorkbench").GetMeasurable(create_reference(part, feature))
    except Exception as exc:
        result["measurement_error"] = str(exc)
        return result
    for attr_name in ("Area", "Length", "Volume"):
        try:
            result[attr_name.lower()] = float(getattr(measurable, attr_name))
        except Exception:
            pass
    try:
        cog = measurable.get_cog()
        result["cog"] = tuple(float(value) for value in cog)
    except Exception:
        pass
    return result


def delete_feature_safely(document: Any, part: Any, feature: Any, feature_name: str) -> dict[str, Any]:
    """
    功能: 尝试删除创建失败或无效的 CATIA 特征。
    输入: Document、Part、特征对象和特征名称。
    输出: 删除结果字典。
    """
    try:
        document.Activate()
    except Exception:
        pass
    selection = document.Selection
    try:
        selection.Clear()
        selection.Add(feature)
        selection.Delete()
        selection.Clear()
        try:
            part.Update()
        except Exception:
            pass
        return {"status": "success", "message": f"已删除无效特征: {feature_name}"}
    except Exception as exc:
        try:
            selection.Clear()
        except Exception:
            pass
        return {"status": "failed", "message": f"删除无效特征失败: {feature_name}: {exc}"}


def delete_hybrid_body_safely(document: Any, part: Any, body_name: str) -> dict[str, Any]:
    """
    功能: 删除指定名称的几何图形集。
    输入: Document、Part 和几何图形集名称。
    输出: 删除结果字典。
    """
    hybrid_body = find_hybrid_body_by_name(part, body_name)
    if hybrid_body is None:
        return {"status": "skipped", "message": f"未找到几何图形集: {body_name}"}
    try:
        document.Activate()
    except Exception:
        pass
    selection = document.Selection
    try:
        selection.Clear()
        selection.Add(hybrid_body)
        selection.Delete()
        selection.Clear()
        try:
            part.Update()
        except Exception:
            pass
        return {"status": "success", "message": f"已删除几何图形集: {body_name}"}
    except Exception as exc:
        try:
            selection.Clear()
        except Exception:
            pass
        return {"status": "failed", "message": f"删除几何图形集失败: {body_name}: {exc}"}


def validate_section_intersection_input(
    document: Any,
    part: Any,
    section_plane_name: str,
    target_row: dict[str, Any],
) -> tuple[Any, Any, dict[str, Any]]:
    """
    功能: 校验截面相交输入元素是否可引用、可测量。
    输入: 轮罩 Document、Part、截面平面名称和目标行。
    输出: 截面 Reference、目标 Reference 和校验信息。
    """
    section_ref, section_metadata = reference_from_hybrid_shape_name(part, section_plane_name)
    target_object = target_row.get("object")
    target_label = str(target_row.get("label") or target_row.get("name") or "<unnamed target>")
    if target_object is None:
        raise RuntimeError(f"目标元素为空，停止创建相交: {target_label}")
    try:
        target_ref = create_reference(part, target_object)
    except Exception as exc:
        raise RuntimeError(f"目标元素无法创建 Reference，停止创建相交: {target_label}: {exc}") from exc

    target_measurement = measure_feature_basic(document, part, target_object)
    if target_measurement.get("measurement_error"):
        raise RuntimeError(
            f"目标元素不可测量，停止创建相交: {target_label}: "
            f"{target_measurement.get('measurement_error')}"
        )
    return section_ref, target_ref, {
        "section_metadata": section_metadata,
        "target_measurement": target_measurement,
        "target_label": target_label,
    }


def scan_section_target_objects(document: Any, part: Any, surface_limit: int = 100) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    """
    功能: 扫描可用于截面的实体和曲面目标，并保留原始 COM 对象。
    输入: 轮罩 Document、Part 和曲面数量上限。
    输出: Body 候选列表和 Surface 候选列表。
    """
    body_targets: list[dict[str, Any]] = []
    bodies = getattr(part, "Bodies", None)
    if bodies is not None:
        for index, body in enumerate(iter_collection(bodies), start=1):
            name = safe_attr_text(body, "Name") or f"Body.{index}"
            measurement = measure_feature_basic(document, part, body)
            visibility = get_object_visibility(document, body)
            body_targets.append(
                {
                    "kind": "body",
                    "name": name,
                    "label": f"{name} [body#{index}]",
                    "object": body,
                    "index": index,
                    "measurement": measurement,
                    "visibility": visibility,
                }
            )

    surface_targets: list[dict[str, Any]] = []
    for hybrid_body in iter_hybrid_bodies(part):
        body_name = safe_attr_text(hybrid_body, "Name")
        for index, hybrid_shape in enumerate(iter_hybrid_shapes(hybrid_body), start=1):
            measurement = measure_feature_basic(document, part, hybrid_shape)
            area = measurement.get("area")
            if not isinstance(area, (int, float)) or float(area) <= 1e-6:
                continue
            name = safe_attr_text(hybrid_shape, "Name") or f"Surface.{index}"
            visibility = get_object_visibility(document, hybrid_shape)
            surface_targets.append(
                {
                    "kind": "surface",
                    "name": name,
                    "label": f"{name} [surface#{len(surface_targets) + 1}]",
                    "object": hybrid_shape,
                    "index": index,
                    "body_name": body_name,
                    "measurement": measurement,
                    "visibility": visibility,
                }
            )
    surface_targets.sort(key=lambda row: float((row.get("measurement") or {}).get("area") or 0.0), reverse=True)
    return body_targets, surface_targets[:surface_limit]


def create_local_axis_line_for_distance(
    part: Any,
    hybrid_body_name: str,
    axis_point_local: Vector,
    axis_direction_local: Vector,
    half_length: float = 10000.0,
) -> Any:
    """
    功能: 在 Part 中创建用于测距的临时轴线。
    输入: Part、几何图形集名称、局部轴点、局部轴方向和半长。
    输出: 轴线 HybridShape。
    """
    target_body = get_or_create_hybrid_body(part, hybrid_body_name)
    direction = normalize_vector(axis_direction_local)
    start_point = add_vectors(axis_point_local, scale_vector(direction, -half_length))
    end_point = add_vectors(axis_point_local, scale_vector(direction, half_length))
    factory = part.HybridShapeFactory
    start_feature = factory.AddNewPointCoord(*start_point)
    end_feature = factory.AddNewPointCoord(*end_point)
    set_if_possible(start_feature, "Name", "_Regulation_Axis_Distance_Start")
    set_if_possible(end_feature, "Name", "_Regulation_Axis_Distance_End")
    target_body.AppendHybridShape(start_feature)
    target_body.AppendHybridShape(end_feature)
    line_feature = factory.AddNewLinePtPt(
        create_reference(part, start_feature),
        create_reference(part, end_feature),
    )
    set_if_possible(line_feature, "Name", "_Regulation_Axis_Distance_Line")
    target_body.AppendHybridShape(line_feature)
    try:
        part.UpdateObject(line_feature)
    except Exception:
        try:
            part.Update()
        except Exception:
            pass
    return line_feature


def minimum_distance_points_between_features(
    document: Any,
    part: Any,
    first_feature: Any,
    second_feature: Any,
) -> tuple[float, Vector | None, Vector | None, dict[str, Any]]:
    """
    功能: 测量同一 Part 内两个几何对象的最小距离和最近点。
    输入: Document、Part、第一个对象和第二个对象。
    输出: 距离、第一个对象最近点、第二个对象最近点和测量信息。
    """
    ref1 = create_reference(part, first_feature)
    ref2 = create_reference(part, second_feature)
    measurable = document.GetWorkbench("SPAWorkbench").GetMeasurable(ref1)
    distance_value = float(measurable.GetMinimumDistance(ref2))
    point1: Vector | None = None
    point2: Vector | None = None
    info: dict[str, Any] = {"method": "GetMinimumDistance"}
    try:
        vba_code = """
Public Function read_minimum_distance_points(measurable, reference2)
    Dim values(8)
    measurable.GetMinimumDistancePoints reference2, values
    read_minimum_distance_points = values
End Function
"""
        coords = document.Application.SystemService.Evaluate(
            vba_code,
            0,
            "read_minimum_distance_points",
            [measurable, ref2],
        )
        coords = tuple(float(value) for value in coords)
        if len(coords) < 6:
            raise RuntimeError(f"GetMinimumDistancePoints 返回坐标数量不足: {len(coords)}")
        point1 = as_vector(coords[0:3], "最小距离点1")
        point2 = as_vector(coords[3:6], "最小距离点2")
        if (
            abs(distance_value) > 1.0e-6
            and vector_length(point1) <= 1.0e-9
            and vector_length(point2) <= 1.0e-9
        ):
            raise RuntimeError("GetMinimumDistancePoints 返回全零点，判定为无效输出")
        info["method"] = "GetMinimumDistancePoints_Evaluate"
        info["raw_points"] = tuple(round(float(value), 6) for value in coords[:9])
        if len(coords) >= 9:
            info["direction"] = tuple(float(value) for value in coords[6:9])
    except Exception as exc:
        info["points_error"] = str(exc)
    return distance_value, point1, point2, info


def can_create_reference(part: Any, feature: Any) -> bool:
    """
    功能: 判断对象是否能在指定 Part 下创建 CATIA Reference。
    输入: Part 和候选对象。
    输出: True 或 False。
    """
    try:
        create_reference(part, feature)
        return True
    except Exception:
        return False


def child_collection(feature: Any, attribute_name: str) -> Any | None:
    """
    功能: 安全读取 CATIA 对象的子集合。
    输入: CATIA 对象和集合属性名。
    输出: 子集合对象或 None。
    """
    for candidate_name in (attribute_name, attribute_name[:1].upper() + attribute_name[1:]):
        try:
            collection = getattr(feature, candidate_name)
            if collection is not None:
                return collection
        except Exception:
            pass
    return None


def collect_measurable_leaves(part: Any, feature: Any, depth: int = 0) -> list[Any]:
    """
    功能: 递归收集对象下可创建 Reference 的叶子几何。
    输入: Part、候选对象和递归深度。
    输出: 可测量叶子对象列表。
    """
    if feature is None or depth > 20:
        return []
    if can_create_reference(part, feature):
        return [feature]

    leaves: list[Any] = []
    for attribute_name in (
        "Shapes",
        "HybridShapes",
        "HybridBodies",
        "OrderedGeometricalSets",
        "Bodies",
    ):
        collection = child_collection(feature, attribute_name)
        if collection is None:
            continue
        for child in iter_collection(collection):
            leaves.extend(collect_measurable_leaves(part, child, depth + 1))
    return leaves


def add_closest_points_to_measure_result(result: dict[str, Any], points: Any) -> None:
    """
    功能: 将 CATIA 最小距离点坐标写入测距结果。
    输入: 测距结果字典和 CATIA 返回的坐标序列。
    输出: 更新后的结果字典。
    """
    if points is None:
        return
    try:
        coords = tuple(float(value) for value in points)
    except Exception as exc:
        result["points_error"] = str(exc)
        return
    if len(coords) < 6:
        result["points_error"] = f"最近点坐标数量不足: {len(coords)}"
        return
    result["point_on_element1"] = as_vector(coords[0:3], "测距点1")
    result["point_on_element2"] = as_vector(coords[3:6], "测距点2")
    result["raw_points"] = tuple(round(value, 6) for value in coords[:9])


def read_minimum_distance_points(
    document: Any,
    measurable: Any,
    reference2: Any,
    distance_value: float | None = None,
) -> tuple[tuple[float, ...] | None, str | None]:
    """
    功能: 读取 CATIA 最小距离两端点坐标，优先用 VBA Evaluate 处理 COM 输出数组。
    输入: Document、Measurable、第二个 Reference 和可选距离值。
    输出: (坐标元组, 错误信息)。成功时错误信息为 None。
    """
    errors: list[str] = []

    def normalize_minimum_distance_points(values: Any, source_label: str) -> tuple[float, ...]:
        """
        功能: 归一化 CATIA 最近点输出；前 6 个坐标必需，后 3 个方向值可选。
        输入: CATIA 返回值和来源标签。
        输出: 6 个坐标值，或 9 个坐标/方向值。
        """
        raw_values = tuple(values)
        if len(raw_values) < 6:
            raise RuntimeError(f"{source_label} 返回坐标数量不足: {len(raw_values)}")
        coords = tuple(float(value) for value in raw_values[:6])
        direction_values: list[float] = []
        if len(raw_values) >= 9:
            try:
                direction_values = [float(value) for value in raw_values[6:9]]
            except Exception:
                direction_values = []
        if direction_values:
            return coords + tuple(direction_values)
        return coords

    vba_code = """
Public Function read_minimum_distance_points(measurable, reference2)
    Dim values(8)
    measurable.GetMinimumDistancePoints reference2, values
    read_minimum_distance_points = values
End Function
"""
    try:
        evaluate_values = document.Application.SystemService.Evaluate(
            vba_code,
            0,
            "read_minimum_distance_points",
            [measurable, reference2],
        )
    except Exception as exc:
        errors.append(f"Evaluate: {exc}")
    else:
        try:
            coords = normalize_minimum_distance_points(evaluate_values, "Evaluate")
        except Exception as exc:
            errors.append(f"Evaluate结果解析: {exc}")
        else:
            return coords, None

    try:
        coords = [0.0] * 9
        measurable.GetMinimumDistancePoints(reference2, coords)
        coords = normalize_minimum_distance_points(coords, "GetMinimumDistancePoints")
        if (
            distance_value is not None
            and abs(float(distance_value)) > 1.0e-6
            and all(abs(value) <= 1.0e-9 for value in coords[:6])
        ):
            raise RuntimeError("GetMinimumDistancePoints 返回全零点，判定为无效输出")
        return coords, None
    except Exception as exc:
        errors.append(f"COM: {exc}")

    try:
        coords = measurable.get_minimum_distance_points(reference2)
        coords = normalize_minimum_distance_points(coords, "get_minimum_distance_points")
        return coords, None
    except Exception as exc:
        errors.append(f"pycatia: {exc}")

    return None, "；".join(errors)


def measure_distance_between(
    caa: Any,
    element1: Any,
    element2: Any,
    part: Any = None,
    part_doc: Any = None,
    return_points: bool = True,
) -> dict[str, Any] | None:
    """
    功能: 测量同一个 Part 内两个元素之间的最小距离。
    输入: CATIA Application、两个测量元素、可选 Part/PartDocument、是否返回最近点。
    输出: 成功返回测距结果字典，失败返回 None。
    """

    def active_part_document() -> Any | None:
        """获取当前 PartDocument。"""
        if part_doc is not None:
            return part_doc
        try:
            return caa.active_document
        except Exception:
            try:
                return caa.ActiveDocument
            except Exception:
                try:
                    return win32com.client.Dispatch(caa.com_object.ActiveDocument)
                except Exception:
                    return None

    def get_part_for_reference(doc: Any) -> Any | None:
        """获取用于创建 Reference 的 Part。"""
        if part is not None:
            return part
        try:
            return doc.Part
        except Exception:
            try:
                return doc.part
            except Exception:
                return None

    def direct_measure(spa: Any, ref_part: Any, obj1: Any, obj2: Any) -> dict[str, Any] | None:
        """直接测量两个引用对象的距离。"""
        try:
            ref1 = create_reference(ref_part, obj1)
            ref2 = create_reference(ref_part, obj2)
        except Exception:
            return None

        try:
            measurable = spa.GetMeasurable(ref1)
        except Exception:
            try:
                measurable = spa.get_measurable(ref1)
            except Exception:
                return None

        try:
            distance = float(measurable.GetMinimumDistance(ref2))
        except Exception:
            try:
                distance = float(measurable.get_minimum_distance(ref2))
            except Exception:
                return None

        points = None
        points_error = None
        if return_points:
            points, points_error = read_minimum_distance_points(doc, measurable, ref2, distance)

        result = {
            "distance": distance,
            "element1": obj1,
            "element2": obj2,
            "element1_name": safe_attr_text(obj1, "Name"),
            "element2_name": safe_attr_text(obj2, "Name"),
            "method": "same_part_measurable",
        }
        add_closest_points_to_measure_result(result, points)
        if points is None and points_error:
            result["points_error"] = points_error
        return result

    doc = active_part_document()
    if doc is None:
        print("测距失败: 无法获取零件文档")
        return None

    try:
        spa = doc.GetWorkbench("SPAWorkbench")
    except Exception:
        try:
            spa = doc.get_workbench("SPAWorkbench")
        except Exception:
            print("测距失败: 无法获取 SPAWorkbench")
            return None

    ref_part = get_part_for_reference(doc)
    if ref_part is None:
        print("测距失败: 无法获取用于创建 Reference 的 Part")
        return None

    result = direct_measure(spa, ref_part, element1, element2)
    if result is not None:
        return result

    leaves1 = collect_measurable_leaves(ref_part, element1)
    leaves2 = collect_measurable_leaves(ref_part, element2)
    best: dict[str, Any] | None = None
    for leaf1 in leaves1:
        for leaf2 in leaves2:
            result = direct_measure(spa, ref_part, leaf1, leaf2)
            if result is None:
                continue
            if best is None or float(result["distance"]) < float(best["distance"]):
                best = result

    if best is not None:
        best["source_element1"] = element1
        best["source_element2"] = element2
        best["source_element1_name"] = safe_attr_text(element1, "Name")
        best["source_element2_name"] = safe_attr_text(element2, "Name")
        best["method"] = "same_part_children_measurable"
        return best

    print(
        f"测距失败: {safe_attr_text(element1, 'Name')} 与 "
        f"{safe_attr_text(element2, 'Name')} 不在同一可测量 Part 引用下，或没有可测量几何"
    )
    return None


def minimum_distance_with_leaf_fallback(
    document: Any,
    part: Any,
    first_feature: Any,
    second_feature: Any,
) -> tuple[float, dict[str, Any]] | None:
    """
    功能: 测量同一 Part 内两个对象的最小距离，直接测失败时拆分到可测叶子。
    输入: Document、Part、第一个对象和第二个对象。
    输出: 距离和测量信息；完全失败时为 None。
    """
    try:
        distance_value, point1, point2, info = minimum_distance_points_between_features(
            document,
            part,
            first_feature,
            second_feature,
        )
        info["source_method"] = "direct"
        if point1 is None or point2 is None:
            info["points_status"] = "unavailable"
        return distance_value, info
    except Exception as exc:
        direct_error = str(exc)

    leaves1 = collect_measurable_leaves(part, first_feature)
    leaves2 = collect_measurable_leaves(part, second_feature)
    best: tuple[float, dict[str, Any]] | None = None
    for leaf1 in leaves1:
        for leaf2 in leaves2:
            try:
                distance_value, point1, point2, info = minimum_distance_points_between_features(
                    document,
                    part,
                    leaf1,
                    leaf2,
                )
            except Exception:
                continue
            info["source_method"] = "leaf_fallback"
            info["source_element_name"] = safe_attr_text(first_feature, "Name")
            info["measured_leaf_name"] = safe_attr_text(leaf1, "Name")
            info["direct_error"] = direct_error
            if point1 is None or point2 is None:
                info["points_status"] = "unavailable"
            candidate = (distance_value, info)
            if best is None or float(candidate[0]) < float(best[0]):
                best = candidate
    return best


def create_axis_clearance_measurement_geometry(
    process_part: Any,
    process_hybrid_body: Any,
    process_transform: Transform | None,
    wheelhouse_transform: Transform,
    measurement_key: str,
    measurement_info: dict[str, Any] | None,
) -> dict[str, Any] | None:
    """
    功能: 将轮罩到轴最小距离的两个端点创建到过程 Part，并创建连线。
    输入: 过程 Part/几何集、过程 Part 变换、轮罩变换、测量项和测距信息。
    输出: 创建结果；缺少点时返回 None。
    """
    if not measurement_info:
        return None
    point_on_body_local = measurement_info.get("point_on_element1")
    point_on_axis_local = measurement_info.get("point_on_element2")
    if point_on_body_local is None or point_on_axis_local is None:
        return None

    body_point_local = as_vector(point_on_body_local, "轮罩侧最近点")
    axis_point_local = as_vector(point_on_axis_local, "轴线侧最近点")
    body_point_world = apply_transform_to_point(wheelhouse_transform, body_point_local)
    axis_point_world = apply_transform_to_point(wheelhouse_transform, axis_point_local)
    body_point_process = (
        inverse_transform_point(process_transform, body_point_world)
        if process_transform is not None
        else body_point_world
    )
    axis_point_process = (
        inverse_transform_point(process_transform, axis_point_world)
        if process_transform is not None
        else axis_point_world
    )

    safe_key = str(measurement_key).replace("/", "_").replace("\\", "_")
    factory = process_part.HybridShapeFactory
    body_point_feature = factory.AddNewPointCoord(*body_point_process)
    axis_point_feature = factory.AddNewPointCoord(*axis_point_process)
    body_point_name = f"{safe_key}_Wheelhouse_MinDistance_Point"
    axis_point_name = f"{safe_key}_Axis_MinDistance_Point"
    line_name = f"{safe_key}_MinDistance_Line"
    set_if_possible(body_point_feature, "Name", body_point_name)
    set_if_possible(axis_point_feature, "Name", axis_point_name)
    process_hybrid_body.AppendHybridShape(body_point_feature)
    process_hybrid_body.AppendHybridShape(axis_point_feature)
    line_feature = factory.AddNewLinePtPt(
        create_reference(process_part, body_point_feature),
        create_reference(process_part, axis_point_feature),
    )
    set_if_possible(line_feature, "Name", line_name)
    process_hybrid_body.AppendHybridShape(line_feature)
    try:
        process_part.UpdateObject(line_feature)
    except Exception:
        try:
            process_part.Update()
        except Exception:
            pass

    return {
        "body_point_name": body_point_name,
        "axis_point_name": axis_point_name,
        "line_name": line_name,
        "body_point_local": round_vector(body_point_local),
        "axis_point_local": round_vector(axis_point_local),
        "body_point_world": round_vector(body_point_world),
        "axis_point_world": round_vector(axis_point_world),
        "body_point_process": round_vector(body_point_process),
        "axis_point_process": round_vector(axis_point_process),
    }


def measure_wheelhouse_to_axis_clearance(
    product_document: Any,
    process_part: Any,
    process_hybrid_body: Any,
    process_transform: Transform | None,
    section_document: Any,
    section_part: Any,
    wheelhouse_transform: Transform,
    target_rows: list[dict[str, Any]],
    feature_row: dict[str, Any],
) -> dict[str, Any]:
    """
    功能: 测量显示 Body 到对应车轮轴线的最小距离，并输出 Front-c 或 Rear-c。
    输入: ProductDocument、过程 Part/几何图形集、过程 Part 装配变换、轮罩截面工作 Part、装配变换、目标行和轴线数据。
    输出: Front-c 或 Rear-c 测量结果。
    """
    wheelhouse_label = str(feature_row.get("wheelhouse") or "")
    measurement_key = wheelhouse_measurement_key(wheelhouse_label, "c")
    axis_point_world = feature_row.get("axis_point_world")
    axis_direction_world = feature_row.get("axis_direction_world")
    if axis_point_world is None or axis_direction_world is None:
        return {
            "status": "failed",
            "measurement_key": measurement_key,
            "message": "缺少轴线世界坐标，无法测量轮罩到轴最小距离。",
        }
    axis_point_local = inverse_transform_point(wheelhouse_transform, axis_point_world)
    axis_direction_local = normalize_vector(
        (
            dot_product(axis_direction_world, wheelhouse_transform.x_axis),
            dot_product(axis_direction_world, wheelhouse_transform.y_axis),
            dot_product(axis_direction_world, wheelhouse_transform.z_axis),
        )
    )
    try:
        axis_line = create_local_axis_line_for_distance(
            section_part,
            "_Regulation_Axis_Clearance_Work",
            axis_point_local,
            axis_direction_local,
        )
    except Exception as exc:
        return {
            "status": "failed",
            "measurement_key": measurement_key,
            "message": f"创建测距轴线失败: {exc}",
        }

    best_row: dict[str, Any] | None = None
    body_distance_rows: list[dict[str, Any]] = []
    for target_row in target_rows:
        target_object = target_row.get("object")
        if target_object is None:
            body_distance_rows.append(
                {
                    "target_label": target_row.get("label"),
                    "status": "failed",
                    "message": "目标对象为空",
                }
            )
            continue
        try:
            measure_result = measure_distance_between(
                product_document.Application,
                target_object,
                axis_line,
                part=section_part,
                part_doc=section_document,
                return_points=True,
            )
            if measure_result is None:
                body_distance_rows.append(
                    {
                        "target_label": target_row.get("label"),
                        "target_name": target_row.get("name"),
                        "status": "failed",
                        "message": "无法测得有效距离",
                    }
                )
                print(f"[警告] Body 到轴测距失败，无有效距离: {target_row.get('label')}")
                continue
            distance_value = float(measure_result["distance"])
            info = measure_result
            row = {
                "distance": distance_value,
                "target_name": target_row.get("name"),
                "target_label": target_row.get("label"),
                "target_visibility": target_row.get("visibility"),
                "measurement_info": info,
            }
            body_distance_rows.append(
                {
                    "target_label": target_row.get("label"),
                    "target_name": target_row.get("name"),
                    "status": "success",
                    "distance": round(float(distance_value), 6),
                    "method": info.get("method"),
                    "measured_leaf_name": info.get("element1_name"),
                }
            )
            print(
                f"Body到轴测距: {measurement_key} target={target_row.get('label')} "
                f"distance={float(distance_value):.6f} method={info.get('method')}"
            )
            if best_row is None or row["distance"] < float(best_row["distance"]):
                best_row = row
        except Exception as exc:
            body_distance_rows.append(
                {
                    "target_label": target_row.get("label"),
                    "target_name": target_row.get("name"),
                    "status": "failed",
                    "message": str(exc),
                }
            )
            print(f"[警告] Body 到轴测距失败: {target_row.get('label')}: {exc}")

    print(f"\n-- {measurement_key} 可见Body到轴距离汇总 --")
    for row in body_distance_rows:
        if row.get("status") == "success":
            leaf_text = f" leaf={row.get('measured_leaf_name')}" if row.get("measured_leaf_name") else ""
            print(
                f"  - {row.get('target_label')}: "
                f"{row.get('distance')} mm method={row.get('method')}{leaf_text}"
            )
        else:
            print(f"  - {row.get('target_label')}: failed, {row.get('message')}")

    if best_row is None:
        return {
            "status": "failed",
            "measurement_key": measurement_key,
            "message": "没有得到显示 Body 到轴线的有效最小距离。",
            "body_distance_rows": body_distance_rows,
        }

    result = {
        "status": "success",
        "measurement_key": measurement_key,
        "measurement_value": round(float(best_row["distance"]), 6),
        "wheelhouse": wheelhouse_label,
        "target_name": best_row.get("target_name"),
        "target_label": best_row.get("target_label"),
        "target_visibility": best_row.get("target_visibility"),
        "measurement_info": best_row.get("measurement_info"),
        "body_distance_rows": body_distance_rows,
    }
    try:
        measurement_geometry = create_axis_clearance_measurement_geometry(
            process_part,
            process_hybrid_body,
            process_transform,
            wheelhouse_transform,
            measurement_key,
            best_row.get("measurement_info"),
        )
        if measurement_geometry is not None:
            result["measurement_geometry"] = measurement_geometry
            print(
                f"已创建{measurement_key}最小距离点和连线: "
                f"{measurement_geometry.get('body_point_name')}, "
                f"{measurement_geometry.get('axis_point_name')}, "
                f"{measurement_geometry.get('line_name')}"
            )
    except Exception as exc:
        result["measurement_geometry_error"] = str(exc)
        print(f"[警告] 创建{measurement_key}最小距离点线失败: {exc}")
    print(
        f"显示Body到轴最小距离: {measurement_key}={result['measurement_value']} mm "
        f"target={result.get('target_label')}"
    )
    return result


def create_section_intersection_curve_from_object(
    document: Any,
    part: Any,
    section_plane_name: str,
    target_row: dict[str, Any],
    result_name: str,
    target_body_name: str = "法规校核截面曲线",
    extend_mode: bool = False,
) -> dict[str, Any]:
    """
    功能: 使用目标 COM 对象直接创建截面相交曲线，避免同名目标解析到同一对象。
    输入: 轮罩 Document、Part、截面平面名称、目标行、结果名称、目标几何图形集和延伸模式。
    输出: 相交曲线创建结果字典。
    """
    target_body = get_or_create_hybrid_body(part, target_body_name)
    factory = part.HybridShapeFactory
    try:
        section_ref, target_ref, input_metadata = validate_section_intersection_input(
            document,
            part,
            section_plane_name,
            target_row,
        )
    except Exception as exc:
        return {
            "status": "failed",
            "message": str(exc),
            "created_features": [],
            "target_name": target_row.get("name"),
            "target_label": target_row.get("label"),
        }

    feature = None
    try:
        feature = factory.AddNewIntersection(section_ref, target_ref)
        if hasattr(feature, "ExtendMode"):
            feature.ExtendMode = bool(extend_mode)
        append_update_hybrid_shape(part, target_body, feature, result_name)
    except Exception as exc:
        return {
            "status": "failed",
            "message": f"相交特征创建或更新失败，等待流程结束后统一删除几何图形集: {exc}",
            "created_features": [],
            "target_name": target_row.get("name"),
            "target_label": target_row.get("label"),
        }

    measurement = measure_feature_basic(document, part, feature)
    if measurement.get("measurement_error") or not isinstance(measurement.get("length"), (int, float)):
        return {
            "status": "failed",
            "message": (
                f"相交结果不是有效曲线，等待流程结束后统一删除几何图形集: "
                f"{measurement.get('measurement_error') or '未测得曲线长度'}"
            ),
            "created_features": [],
            "target_name": target_row.get("name"),
            "target_label": target_row.get("label"),
            "measurement": measurement,
        }
    return {
        "status": "success",
        "message": f"Created section intersection curve '{result_name}'.",
        "created_features": [result_name],
        "feature_object": feature,
        "target_name": target_row.get("name"),
        "target_label": target_row.get("label"),
        "measurement": measurement,
        "input_metadata": input_metadata,
    }


def feature_name_indicates_axis(feature: Any) -> bool:
    """
    功能: 判断特征名称是否像车轮轴线。
    输入: CATIA 特征。
    输出: 是否命中轴线关键字。
    """
    name = safe_attr_text(feature, "Name").casefold()
    return any(keyword.casefold() in name for keyword in AXIS_NAME_KEYWORDS)


def axis_system_name_indicates_model_axis(axis_system: Any) -> bool:
    """
    功能: 判断 AxisSystem 名称是否明确表示车轮轴。
    输入: AxisSystem。
    输出: 是否参与轴线分析。
    """
    name = safe_attr_text(axis_system, "Name").casefold()
    return any(keyword.casefold() in name for keyword in AXIS_SYSTEM_SPECIFIC_KEYWORDS)


def iter_axis_feature_directions(context: WheelPartContext) -> Iterable[tuple[str, Vector]]:
    """
    功能: 从零件中寻找轴特征并读取局部方向。
    输入: 叶子零件上下文。
    输出: (特征名, 方向) 迭代器。
    """
    seen_feature_names: set[str] = set()

    for hybrid_body in iter_hybrid_bodies(context.part):
        for feature in iter_hybrid_shapes(hybrid_body):
            feature_name = safe_attr_text(feature, "Name")
            if not feature_name or not feature_name_indicates_axis(feature):
                continue
            if feature_name in seen_feature_names:
                continue
            try:
                direction = get_feature_direction(context.document, context.part, feature)
            except Exception:
                continue
            seen_feature_names.add(feature_name)
            yield feature_name, direction

    try:
        axis_systems = context.part.AxisSystems
    except Exception:
        axis_systems = None
    if axis_systems is None:
        return

    for axis_system in iter_collection(axis_systems):
        axis_name = safe_attr_text(axis_system, "Name", "AxisSystem")
        if not axis_system_name_indicates_model_axis(axis_system):
            continue
        if axis_name in seen_feature_names:
            continue
        for index, direction in enumerate(
            read_axis_system_directions(context.document, axis_system),
            start=1,
        ):
            yield f"{axis_name}_{index}", direction


def iter_axis_feature_records(context: WheelPartContext) -> Iterable[tuple[str, Vector, Vector]]:
    """
    功能: 从零件中寻找轴线候选。
    输入: 叶子零件上下文。
    输出: (特征名, 局部方向, 局部轴上一点) 迭代器。
    """
    seen_feature_names: set[str] = set()

    for hybrid_body in iter_hybrid_bodies(context.part):
        for feature in iter_hybrid_shapes(hybrid_body):
            feature_name = safe_attr_text(feature, "Name")
            if not feature_name or not feature_name_indicates_axis(feature):
                continue
            if feature_name in seen_feature_names:
                continue
            try:
                direction = get_feature_direction(context.document, context.part, feature)
                axis_point = get_feature_axis_point(context.document, context.part, feature)
            except Exception:
                continue
            seen_feature_names.add(feature_name)
            yield feature_name, direction, axis_point

    try:
        axis_systems = context.part.AxisSystems
    except Exception:
        axis_systems = None
    if axis_systems is None:
        return

    for axis_system in iter_collection(axis_systems):
        axis_name = safe_attr_text(axis_system, "Name", "AxisSystem")
        if not axis_system_name_indicates_model_axis(axis_system):
            continue
        if axis_name in seen_feature_names:
            continue
        origin = read_axis_system_origin(context.document, axis_system)
        if origin is None:
            continue
        for index, direction in enumerate(
            read_axis_system_directions(context.document, axis_system),
            start=1,
        ):
            yield f"{axis_name}_{index}", direction, origin


def select_unique_axis_parts(
    wheel_part_contexts: list[WheelPartContext],
    tolerance_degrees: float = AXIS_DIRECTION_TOLERANCE_DEGREES,
) -> tuple[list[AxisRecord], set[str], list[dict[str, Any]]]:
    """
    功能: 旧版按轴方向去重的辅助函数。
    输入: 车轮零件上下文和角度容差。
    输出: 唯一轴、可见路径集合和扫描结果。
    """
    unique_axes: list[AxisRecord] = []
    visible_component_paths: set[str] = set()
    scanned_axes: list[dict[str, Any]] = []

    for context in wheel_part_contexts:
        component_axis_count = 0
        for feature_name, direction in iter_axis_feature_directions(context):
            component_axis_count += 1
            is_duplicate = any(
                directions_match(direction, record.direction, tolerance_degrees)
                for record in unique_axes
            )
            scanned_axes.append(
                {
                    "component_path": context.component_path,
                    "component_name": context.component_name,
                    "component_part_number": context.component_part_number,
                    "feature_name": feature_name,
                    "direction": tuple(round(value, 8) for value in direction),
                    "duplicate": is_duplicate,
                }
            )
            if is_duplicate:
                continue
            unique_axes.append(
                AxisRecord(
                    direction=direction,
                    feature_name=feature_name,
                    component_path=context.component_path,
                    component_name=context.component_name,
                    component_part_number=context.component_part_number,
                )
            )
            visible_component_paths.add(context.component_path)
        if component_axis_count == 0:
            scanned_axes.append(
                {
                    "component_path": context.component_path,
                    "component_name": context.component_name,
                    "component_part_number": context.component_part_number,
                    "feature_name": None,
                    "direction": None,
                    "duplicate": None,
                    "warning": "未找到可测量轴特征",
                }
            )

    return unique_axes, visible_component_paths, scanned_axes


def round_vector(vector: Vector | None, digits: int = 6) -> tuple[float, float, float] | None:
    """
    功能: 对向量做小数位格式化。
    输入: 向量和位数。
    输出: 四舍五入后的向量或 None。
    """
    if vector is None:
        return None
    return tuple(round(value, digits) for value in vector)


def bounding_box_to_dict(bbox: BoundingBox | None) -> dict[str, Any] | None:
    """
    功能: 将 BoundingBox 转为 JSON 友好的字典。
    输入: BoundingBox 或 None。
    输出: 字典或 None。
    """
    if bbox is None:
        return None
    return {
        "min_point": round_vector(bbox.min_point),
        "max_point": round_vector(bbox.max_point),
        "size": round_vector(bbox.size),
        "diagonal": round(bbox.diagonal, 6),
        "volume": round(bbox.volume, 6),
        "source": bbox.source,
    }


def bbox_center_point(bbox: BoundingBox) -> Vector:
    """
    功能: 计算包围盒中心点。
    输入: BoundingBox。
    输出: 世界坐标中心点。
    """
    return (
        (bbox.min_point[0] + bbox.max_point[0]) / 2.0,
        (bbox.min_point[1] + bbox.max_point[1]) / 2.0,
        (bbox.min_point[2] + bbox.max_point[2]) / 2.0,
    )


def bbox_axis_from_largest_faces(bbox: BoundingBox) -> dict[str, Any]:
    """
    功能: 取包围盒面积最大的相对面中心点连线作为轮胎轴线。
    输入: BoundingBox。
    输出: 轴线方向、轴线上点、两侧面中心点和面信息。
    """
    min_x, min_y, min_z = bbox.min_point
    max_x, max_y, max_z = bbox.max_point
    center_x, center_y, center_z = bbox_center_point(bbox)
    size_x, size_y, size_z = bbox.size
    face_options = [
        {
            "axis": "X",
            "face_area": max(size_y, 0.0) * max(size_z, 0.0),
            "axis_length": max(size_x, 0.0),
            "negative_face_center": (min_x, center_y, center_z),
            "positive_face_center": (max_x, center_y, center_z),
            "direction": (1.0, 0.0, 0.0),
        },
        {
            "axis": "Y",
            "face_area": max(size_x, 0.0) * max(size_z, 0.0),
            "axis_length": max(size_y, 0.0),
            "negative_face_center": (center_x, min_y, center_z),
            "positive_face_center": (center_x, max_y, center_z),
            "direction": (0.0, 1.0, 0.0),
        },
        {
            "axis": "Z",
            "face_area": max(size_x, 0.0) * max(size_y, 0.0),
            "axis_length": max(size_z, 0.0),
            "negative_face_center": (center_x, center_y, min_z),
            "positive_face_center": (center_x, center_y, max_z),
            "direction": (0.0, 0.0, 1.0),
        },
    ]
    best_face = max(face_options, key=lambda item: (item["face_area"], item["axis_length"]))
    if float(best_face["axis_length"]) <= 1.0e-9:
        raise RuntimeError("包围盒最大面中心点连线长度为 0，无法生成轮胎轴线。")
    return {
        "axis": best_face["axis"],
        "axis_direction_world": normalized_axis_direction(best_face["direction"]),
        "axis_point_world": bbox_center_point(bbox),
        "negative_face_center": best_face["negative_face_center"],
        "positive_face_center": best_face["positive_face_center"],
        "face_area": best_face["face_area"],
        "axis_length": best_face["axis_length"],
    }


def bbox_size_signature(bbox: BoundingBox) -> tuple[float, float, float]:
    """
    功能: 生成忽略 X/Y/Z 顺序的包围盒尺寸签名。
    输入: BoundingBox。
    输出: 从小到大排序的尺寸元组。
    """
    size_values = sorted(float(value) for value in bbox.size)
    return size_values[0], size_values[1], size_values[2]


def bbox_sizes_match(
    first: BoundingBox,
    second: BoundingBox,
    tolerance: float = TIRE_BBOX_SIZE_EQUAL_TOLERANCE,
) -> bool:
    """
    功能: 判断两个轮胎包围盒尺寸是否在容差内一致。
    输入: 两个 BoundingBox 和尺寸容差。
    输出: 是否一致。
    """
    first_signature = bbox_size_signature(first)
    second_signature = bbox_size_signature(second)
    return all(
        abs(first_value - second_value) <= tolerance
        for first_value, second_value in zip(first_signature, second_signature)
    )


def wheel_context_group_key(context: WheelPartContext) -> str:
    """
    功能: 将叶子零件归到车轮装配的一级实例分组。
    输入: WheelPartContext。
    输出: 分组键。
    """
    parts = [part for part in context.component_path.split("/") if part]
    if len(parts) >= 2:
        return "/".join(parts[:2])
    return context.component_path


def measured_part_size_score(row: dict[str, Any]) -> tuple[float, float]:
    """
    功能: 生成包围盒候选排序分数，优先体积，其次对角线。
    输入: 测量行。
    输出: 排序元组。
    """
    bbox = row.get("bbox")
    if not isinstance(bbox, BoundingBox):
        return (0.0, 0.0)
    return bbox.volume, bbox.diagonal


def build_wheelhouse_near_part_rows(
    product_document: Any,
    wheel_part_contexts: list[WheelPartContext],
    wheelhouse_cog_items: list[tuple[str, Vector]] | None,
    near_distance: float = WHEELHOUSE_NEAR_DISTANCE,
) -> tuple[list[dict[str, Any]], set[str], list[str]]:
    """
    功能: 按零件重心到轮罩重心的距离阈值筛选轮罩附近零件。
    输入: ProductDocument、车轮叶子零件上下文、轮罩重心列表和近邻距离阈值。
    输出: 零件重心行、近邻零件路径集合和告警列表。
    """
    rows: list[dict[str, Any]] = []
    warnings: list[str] = []
    for context in wheel_part_contexts:
        try:
            component_cog = evaluate_product_cog(
                product_document,
                context.product,
                context.component_path,
            )
        except Exception as exc:
            warnings.append(f"{context.component_path}: 零件重心读取失败，无法参与轮罩近邻筛选: {exc}")
            continue
        wheelhouse_distances = {
            label: distance(component_cog, wheelhouse_cog)
            for label, wheelhouse_cog in (wheelhouse_cog_items or [])
        }
        nearest_label = None
        nearest_distance = None
        if wheelhouse_distances:
            nearest_label, nearest_distance = min(
                wheelhouse_distances.items(),
                key=lambda item: item[1],
            )
        rows.append(
            {
                "context": context,
                "component_cog_world": component_cog,
                "wheelhouse_distances": wheelhouse_distances,
                "nearest_wheelhouse": nearest_label,
                "nearest_wheelhouse_distance": nearest_distance,
                "near_wheelhouse_labels": [],
            }
        )

    if not wheelhouse_cog_items:
        for row in rows:
            row["near_wheelhouse_labels"] = [WHEEL_ASSEMBLY_LABEL]
        return rows, {row["context"].component_path for row in rows}, warnings

    near_component_paths: set[str] = set()
    for wheelhouse_label, _wheelhouse_cog in wheelhouse_cog_items:
        matched_count = 0
        for row in rows:
            wheelhouse_distance = row["wheelhouse_distances"].get(wheelhouse_label, math.inf)
            if wheelhouse_distance > near_distance:
                continue
            row["near_wheelhouse_labels"].append(wheelhouse_label)
            near_component_paths.add(row["context"].component_path)
            matched_count += 1
        if matched_count == 0:
            warnings.append(
                f"{wheelhouse_label}: {near_distance:.3f} mm 范围内没有找到车轮装配零件重心。"
            )
    return rows, near_component_paths, warnings


def candidate_name_text(candidate: WheelCandidate) -> str:
    """
    功能: 拼接候选件名称文本用于关键字判断。
    输入: WheelCandidate。
    输出: 小写文本。
    """
    parts = (
        candidate.context.component_name,
        candidate.context.component_part_number,
        candidate.context.component_path,
        candidate.feature_name or "",
    )
    return " ".join(parts).casefold()


def text_has_keyword(text: str, keywords: tuple[str, ...]) -> bool:
    """
    功能: 判断文本是否包含任一关键字。
    输入: 文本和关键字列表。
    输出: 是否命中。
    """
    return any(keyword.casefold() in text for keyword in keywords)


def radius_level_count(radii: list[float], tolerance: float = 2.0) -> int:
    """
    功能: 统计半径层级数量。
    输入: 半径列表和半径归并容差。
    输出: 半径层级数量。
    """
    levels: list[float] = []
    for radius in sorted(radii):
        if not any(abs(radius - level) <= tolerance for level in levels):
            levels.append(radius)
    return len(levels)


def group_topology_circles_by_axis(
    circles: list[tuple[TopologyCircleRecord, Vector, Vector]],
    axis_tolerance_degrees: float = TOPOLOGY_AXIS_DIRECTION_TOLERANCE_DEGREES,
    center_tolerance: float = TOPOLOGY_CENTER_CLUSTER_TOLERANCE,
) -> list[list[tuple[TopologyCircleRecord, Vector, Vector]]]:
    """
    功能: 按圆轴方向和圆心到轴线距离对圆拓扑记录分组。
    输入: 圆记录、方向容差和中心聚类容差。
    输出: 同轴圆分组列表。
    """
    groups: list[list[tuple[TopologyCircleRecord, Vector, Vector]]] = []
    for circle_item in circles:
        _circle, center, direction = circle_item
        matched_group: list[tuple[TopologyCircleRecord, Vector, Vector]] | None = None
        for group in groups:
            _seed_circle, seed_center, seed_direction = group[0]
            if not directions_match(direction, seed_direction, axis_tolerance_degrees):
                continue
            if point_to_axis_distance(center, seed_center, seed_direction) > center_tolerance:
                continue
            matched_group = group
            break
        if matched_group is None:
            groups.append([circle_item])
        else:
            matched_group.append(circle_item)
    return groups


def analyze_tire_topology(
    local_circles: list[TopologyCircleRecord],
    world_transform: Transform,
) -> TopologyAxisAnalysis | None:
    """
    功能: 根据圆线拓扑推导轮胎候选轴和轮胎形态评分。
    输入: 局部圆拓扑记录和装配世界变换。
    输出: 最优 TopologyAxisAnalysis，无法判断时为 None。
    """
    world_circles: list[tuple[TopologyCircleRecord, Vector, Vector]] = []
    for circle in local_circles:
        try:
            center_world = apply_transform_to_point(world_transform, circle.center_local)
            direction_world = normalized_axis_direction(
                apply_transform_to_direction(world_transform, circle.axis_direction_local)
            )
        except Exception:
            continue
        world_circles.append((circle, center_world, direction_world))
    if not world_circles:
        return None

    best_analysis: TopologyAxisAnalysis | None = None
    for group in group_topology_circles_by_axis(world_circles):
        if len(group) < MIN_TOPOLOGY_CIRCLE_COUNT:
            continue
        circles = [item[0] for item in group]
        centers = [item[1] for item in group]
        directions = [item[2] for item in group]
        radii = [circle.radius for circle in circles]
        levels = radius_level_count(radii)
        if levels < MIN_TOPOLOGY_RADIUS_LEVELS:
            continue

        axis_direction = normalized_axis_direction(average_vectors(directions))
        axis_point = average_vectors(centers)
        center_distances = [point_to_axis_distance(center, axis_point, axis_direction) for center in centers]
        center_cluster_radius = max(center_distances) if center_distances else 0.0
        min_radius = min(radii)
        max_radius = max(radii)
        radius_span = max_radius - min_radius
        score = (
            len(group) * 10.0
            + levels * 25.0
            + max(radius_span, 0.0) * 0.5
            + max(max_radius, 0.0) * 0.1
            - center_cluster_radius * 0.5
        )
        analysis = TopologyAxisAnalysis(
            axis_direction_world=axis_direction,
            axis_point_world=axis_point,
            circle_count=len(group),
            radius_levels=levels,
            center_cluster_radius=center_cluster_radius,
            min_radius=min_radius,
            max_radius=max_radius,
            score=score,
            feature_names=[circle.feature_name for circle in circles],
        )
        if best_analysis is None or analysis.score > best_analysis.score:
            best_analysis = analysis
    return best_analysis


def tire_priority_score(candidate: WheelCandidate) -> float:
    """
    功能: 计算候选件作为轮胎的优先级。
    输入: WheelCandidate。
    输出: 分数，越大越优先。
    """
    text = candidate_name_text(candidate)
    score = 0.0
    if candidate.bbox_world is not None:
        score += candidate.bbox_world.diagonal
        score += candidate.bbox_world.volume * 1e-9
    if candidate.volume is not None:
        score += max(candidate.volume, 0.0) * 1e-6
    score += candidate.topology_score
    score += candidate.topology_circle_count * 2.0
    score += candidate.topology_radius_levels * 20.0
    if candidate.axis_source == "topology":
        score += 500.0
    if text_has_keyword(text, TIRE_NAME_KEYWORDS):
        score += 1_000_000.0
    if text_has_keyword(text, HUB_NAME_KEYWORDS):
        score -= 1_000_000.0
    return score


def candidate_to_dict(candidate: WheelCandidate) -> dict[str, Any]:
    """
    功能: 将车轮候选件转为结果字典。
    输入: WheelCandidate。
    输出: JSON 友好的字典。
    """
    return {
        "component_path": candidate.context.component_path,
        "component_name": candidate.context.component_name,
        "component_part_number": candidate.context.component_part_number,
        "feature_name": candidate.feature_name,
        "axis_source": candidate.axis_source,
        "axis_direction_world": round_vector(candidate.axis_direction_world, 8),
        "axis_point_world": round_vector(candidate.axis_point_world),
        "component_cog_world": round_vector(candidate.component_cog_world),
        "bbox_world": bounding_box_to_dict(candidate.bbox_world),
        "mass": None if candidate.mass is None else round(candidate.mass, 6),
        "volume": None if candidate.volume is None else round(candidate.volume, 6),
        "topology_circle_count": candidate.topology_circle_count,
        "topology_radius_levels": candidate.topology_radius_levels,
        "topology_score": round(candidate.topology_score, 6),
        "tire_priority_score": round(tire_priority_score(candidate), 6),
        "axis_face_centers_world": (
            [round_vector(point) for point in candidate.axis_face_centers_world]
            if candidate.axis_face_centers_world
            else None
        ),
        "wheel_group_key": candidate.wheel_group_key,
        "warnings": candidate.warnings,
    }


def build_wheel_candidates(
    product_document: Any,
    root_product: Any,
    wheel_part_contexts: list[WheelPartContext],
    wheelhouse_cog_items: list[tuple[str, Vector]] | None = None,
) -> tuple[list[WheelCandidate], list[dict[str, Any]], list[str], list[dict[str, Any]], list[dict[str, Any]]]:
    """
    功能: 先按轮罩重心筛选近邻零件，再用装配后包围盒识别每个轮罩对应轮胎。
    输入: ProductDocument、根 Product、叶子零件上下文列表和可选轮罩重心列表。
    输出: 候选列表、非轮胎排除列表、告警列表、近邻筛选行和显示状态行。
    """
    candidates: list[WheelCandidate] = []
    excluded_no_axis_components: list[dict[str, Any]] = []
    warnings: list[str] = []
    measured_parts: list[dict[str, Any]] = []
    cog_rows, visible_component_paths, proximity_warnings = build_wheelhouse_near_part_rows(
        product_document,
        wheel_part_contexts,
        wheelhouse_cog_items,
    )
    warnings.extend(proximity_warnings)
    initially_hidden_component_paths: set[str] = set()
    for row in cog_rows:
        context = row["context"]
        visibility = get_object_visibility(product_document, context.product)
        row["initial_visibility"] = visibility
        if visibility.get("status") == "hidden" or visibility.get("visible") is False:
            initially_hidden_component_paths.add(context.component_path)
    if initially_hidden_component_paths:
        visible_component_paths = visible_component_paths - initially_hidden_component_paths
        warnings.append(
            f"车轮装配中有 {len(initially_hidden_component_paths)} 个原始隐藏零件，"
            "已排除包围盒测量。"
        )
    visibility_rows = apply_wheel_part_visibility(
        product_document,
        wheel_part_contexts,
        visible_component_paths,
    )
    cog_row_by_path = {
        row["context"].component_path: row
        for row in cog_rows
    }

    for row in cog_rows:
        context = row["context"]
        if context.component_path in visible_component_paths:
            continue
        initially_hidden = context.component_path in initially_hidden_component_paths
        excluded_no_axis_components.append(
            {
                "component_path": context.component_path,
                "component_name": context.component_name,
                "component_part_number": context.component_part_number,
                "component_cog_world": round_vector(row["component_cog_world"]),
                "wheelhouse_distances": {
                    label: round(value, 6)
                    for label, value in row["wheelhouse_distances"].items()
                },
                "nearest_wheelhouse": row["nearest_wheelhouse"],
                "nearest_wheelhouse_distance": (
                    None
                    if row["nearest_wheelhouse_distance"] is None
                    else round(row["nearest_wheelhouse_distance"], 6)
                ),
                "initial_visibility": row.get("initial_visibility"),
                "reason": "initially_hidden" if initially_hidden else "far_from_wheelhouse_cog",
                "warnings": (
                    ["零件原始状态为隐藏，未参与轮胎包围盒测量。"]
                    if initially_hidden
                    else []
                ),
            }
        )

    for context in visible_contexts(wheel_part_contexts, visible_component_paths):
        context_warnings: list[str] = []
        cog_row = cog_row_by_path.get(context.component_path)
        if cog_row is None:
            continue
        try:
            bbox_result = get_product_bounding_box_by_far_plane(
                context.product,
                root_product=root_product,
            )
        except Exception as exc:
            bbox_result = {"status": "failed", "message": str(exc)}
        bbox = bbox_result.get("bbox")
        bbox_source = "far_plane_axis_extreme"
        visibility_skipped = bbox_result.get("visibility_skipped") or []
        if visibility_skipped:
            context_warnings.append(f"包围盒测量跳过隐藏对象数量: {len(visibility_skipped)}。")
        if bbox is None:
            message = bbox_result.get("message") or "未知错误"
            context_warnings.append(f"远平面极值法包围盒测量失败: {message}")
            context_warnings.append("已禁用 Product.Analyze 包围盒回退，避免隐藏或辅助几何参与测量。")
        if bbox is None:
            excluded_no_axis_components.append(
                {
                    "component_path": context.component_path,
                    "component_name": context.component_name,
                    "component_part_number": context.component_part_number,
                    "reason": "bbox_measurement_failed",
                    "warnings": context_warnings.copy(),
                }
            )
            warnings.append(f"{context.component_path}: 包围盒测量失败，无法参与轮胎筛选。")
            continue
        try:
            axis_info = bbox_axis_from_largest_faces(bbox)
        except Exception as exc:
            excluded_no_axis_components.append(
                {
                    "component_path": context.component_path,
                    "component_name": context.component_name,
                    "component_part_number": context.component_part_number,
                    "bbox_world": bounding_box_to_dict(bbox),
                    "reason": "bbox_axis_failed",
                    "warnings": [*context_warnings, str(exc)],
                }
            )
            warnings.append(f"{context.component_path}: 包围盒最大面轴线生成失败: {exc}")
            continue
        mass = evaluate_product_scalar(context.product, "Mass")
        volume = evaluate_product_scalar(context.product, "Volume")
        if bbox_source != bbox.source:
            context_warnings.append(f"包围盒来源: {bbox.source}")
        measured_parts.append(
            {
                "context": context,
                "group_key": wheel_context_group_key(context),
                "near_wheelhouse_labels": list(cog_row["near_wheelhouse_labels"]),
                "wheelhouse_distances": dict(cog_row["wheelhouse_distances"]),
                "part_cog_world": cog_row["component_cog_world"],
                "bbox": bbox,
                "axis_info": axis_info,
                "component_center_world": axis_info["axis_point_world"],
                "mass": mass,
                "volume": volume,
                "warnings": context_warnings.copy(),
                "bbox_measurement_source": bbox.source,
            }
        )

    if not measured_parts:
        return candidates, excluded_no_axis_components, warnings, [
            {
                "component_path": row["context"].component_path,
                "component_name": row["context"].component_name,
                "component_part_number": row["context"].component_part_number,
                "component_cog_world": round_vector(row["component_cog_world"]),
                "wheelhouse_distances": {
                    label: round(value, 6)
                    for label, value in row["wheelhouse_distances"].items()
                },
                "near_wheelhouse_labels": row["near_wheelhouse_labels"],
                "selected_for_bbox": row["context"].component_path in visible_component_paths,
                "initial_visibility": row.get("initial_visibility"),
            }
            for row in cog_rows
        ], visibility_rows

    selected_rows: list[dict[str, Any]] = []
    if wheelhouse_cog_items:
        for wheelhouse_label, _wheelhouse_cog in wheelhouse_cog_items:
            near_rows = [
                row
                for row in measured_parts
                if wheelhouse_label in row["near_wheelhouse_labels"]
            ]
            if not near_rows:
                warnings.append(f"{wheelhouse_label}: 轮罩重心附近没有可测包围盒零件。")
                continue
            selected = max(near_rows, key=measured_part_size_score)
            selected["selection_reason"] = "largest_bbox_near_wheelhouse"
            selected["wheelhouse_label"] = wheelhouse_label
            selected_rows.append(selected)
    else:
        grouped_parts: dict[str, list[dict[str, Any]]] = {}
        for row in measured_parts:
            grouped_parts.setdefault(str(row["group_key"]), []).append(row)
        for group_key, group_rows in grouped_parts.items():
            selected = max(group_rows, key=measured_part_size_score)
            selected["selection_reason"] = "largest_bbox_in_wheel_group"
            selected["wheelhouse_label"] = group_key
            selected_rows.append(selected)

    expected_tire_count = len(wheelhouse_cog_items) if wheelhouse_cog_items else EXPECTED_TIRE_COUNT
    if len(selected_rows) != expected_tire_count:
        warnings.append(
            f"按轮罩近邻包围盒识别出的轮胎数量为 {len(selected_rows)}，"
            f"期望为 {expected_tire_count}。请检查车轮装配层级、轮罩重心或 {WHEELHOUSE_NEAR_DISTANCE:.3f} mm 分组阈值。"
        )
    if len(selected_rows) > 1:
        reference_bbox = selected_rows[0]["bbox"]
        mismatched_rows = [
            row
            for row in selected_rows[1:]
            if not bbox_sizes_match(row["bbox"], reference_bbox)
        ]
        if mismatched_rows:
            warnings.append(
                f"识别出的轮胎包围盒尺寸不一致，"
                f"尺寸容差为 {TIRE_BBOX_SIZE_EQUAL_TOLERANCE:.3f} mm。"
            )

    selected_ids = {id(row) for row in selected_rows}
    selected_component_paths = {row["context"].component_path for row in selected_rows}
    excluded_no_axis_components = [
        item
        for item in excluded_no_axis_components
        if item.get("component_path") not in selected_component_paths
    ]
    for row in measured_parts:
        if id(row) in selected_ids:
            continue
        context = row["context"]
        selected_for_labels = [
            selected_row.get("wheelhouse_label")
            for selected_row in selected_rows
            if selected_row.get("wheelhouse_label") in row["near_wheelhouse_labels"]
        ]
        excluded_no_axis_components.append(
            {
                "component_path": context.component_path,
                "component_name": context.component_name,
                "component_part_number": context.component_part_number,
                "component_cog_world": round_vector(row["component_center_world"]),
                "part_cog_world": round_vector(row["part_cog_world"]),
                "bbox_world": bounding_box_to_dict(row["bbox"]),
                "mass": None if row["mass"] is None else round(row["mass"], 6),
                "volume": None if row["volume"] is None else round(row["volume"], 6),
                "wheel_group_key": row["group_key"],
                "near_wheelhouse_labels": row["near_wheelhouse_labels"],
                "selected_for_wheelhouse_labels": selected_for_labels,
                "wheelhouse_distances": {
                    label: round(value, 6)
                    for label, value in row["wheelhouse_distances"].items()
                },
                "reason": "smaller_bbox_near_wheelhouse",
                "warnings": row["warnings"],
            }
        )

    for row in selected_rows:
        context = row["context"]
        bbox = row["bbox"]
        axis_info = row["axis_info"]
        selection_reason = row.get("selection_reason", "largest_bbox_in_wheel_group")
        candidate_warnings = row["warnings"].copy()
        candidate_warnings.append(
            f"轮胎识别来源={selection_reason}，"
            f"对应轮罩={row.get('wheelhouse_label')}，"
            f"最大面轴={axis_info['axis']}，"
            f"最大面面积={float(axis_info['face_area']):.3f}，"
            f"轴向长度={float(axis_info['axis_length']):.3f}。"
        )
        candidates.append(
            WheelCandidate(
                context=context,
                feature_name="BBox_Largest_Face_Centers",
                axis_source="bbox_largest_faces",
                axis_direction_world=axis_info["axis_direction_world"],
                axis_point_world=axis_info["axis_point_world"],
                component_cog_world=row["component_center_world"],
                bbox_world=bbox,
                mass=row["mass"],
                volume=row["volume"],
                topology_circle_count=0,
                topology_radius_levels=0,
                topology_score=0.0,
                warnings=candidate_warnings,
                axis_face_centers_world=(
                    axis_info["negative_face_center"],
                    axis_info["positive_face_center"],
                ),
                wheel_group_key=str(row.get("wheelhouse_label") or row["group_key"]),
            )
        )

    proximity_rows = [
        {
            "component_path": row["context"].component_path,
            "component_name": row["context"].component_name,
            "component_part_number": row["context"].component_part_number,
            "component_cog_world": round_vector(row["component_cog_world"]),
            "wheelhouse_distances": {
                label: round(value, 6)
                for label, value in row["wheelhouse_distances"].items()
            },
            "nearest_wheelhouse": row["nearest_wheelhouse"],
            "nearest_wheelhouse_distance": (
                None
                if row["nearest_wheelhouse_distance"] is None
                else round(row["nearest_wheelhouse_distance"], 6)
            ),
            "near_wheelhouse_labels": row["near_wheelhouse_labels"],
            "selected_for_bbox": row["context"].component_path in visible_component_paths,
            "initial_visibility": row.get("initial_visibility"),
        }
        for row in cog_rows
    ]
    return candidates, excluded_no_axis_components, warnings, proximity_rows, visibility_rows


def candidates_same_position(
    first: WheelCandidate,
    second: WheelCandidate,
    axis_tolerance_degrees: float,
    position_tolerance: float,
) -> bool:
    """
    功能: 判断两个候选件是否属于同一空间车轮位置。
    输入: 两个候选件、角度容差和位置容差。
    输出: 是否同组。
    """
    cog_distance = distance(first.component_cog_world, second.component_cog_world)
    if first.axis_direction_world is None or first.axis_point_world is None:
        return cog_distance <= position_tolerance
    if second.axis_direction_world is None or second.axis_point_world is None:
        return cog_distance <= position_tolerance
    if not directions_match(first.axis_direction_world, second.axis_direction_world, axis_tolerance_degrees):
        return False
    line_distance = axis_line_distance(
        first.axis_point_world,
        first.axis_direction_world,
        second.axis_point_world,
        second.axis_direction_world,
    )
    return line_distance <= position_tolerance and cog_distance <= position_tolerance


def cluster_wheel_candidates(
    candidates: list[WheelCandidate],
    axis_tolerance_degrees: float,
    position_tolerance: float,
) -> list[WheelPositionGroup]:
    """
    功能: 将有轴候选件按空间位置聚类。
    输入: 候选列表、轴线角度容差和位置容差。
    输出: 车轮位置组列表。
    """
    groups: list[WheelPositionGroup] = []
    for candidate in candidates:
        matched_group: WheelPositionGroup | None = None
        for group in groups:
            if any(
                candidates_same_position(
                    candidate,
                    existing,
                    axis_tolerance_degrees,
                    position_tolerance,
                )
                for existing in group.candidates
            ):
                matched_group = group
                break
        if matched_group is None:
            group_id = f"wheel_position_{len(groups) + 1:03d}"
            groups.append(WheelPositionGroup(group_id, [candidate], None, [], []))
        else:
            matched_group.candidates.append(candidate)
    return groups


def choose_group_tire_candidate(
    group: WheelPositionGroup,
    tire_hub_cog_tolerance: float,
) -> None:
    """
    功能: 在一个位置组内选择 Tire 代表件并记录排除项。
    输入: 位置组和轮胎/轮毂重心容差。
    输出: 直接修改 group。
    """
    unique_by_component: dict[str, WheelCandidate] = {}
    for candidate in group.candidates:
        existing = unique_by_component.get(candidate.context.component_path)
        if existing is None or tire_priority_score(candidate) > tire_priority_score(existing):
            unique_by_component[candidate.context.component_path] = candidate

    unique_candidates = list(unique_by_component.values())
    if not unique_candidates:
        group.tire_candidate = None
        group.excluded_candidates = []
        group.warnings.append("位置组内没有可用候选件。")
        return

    sorted_candidates = sorted(unique_candidates, key=tire_priority_score, reverse=True)
    tire_candidate = sorted_candidates[0]
    group.tire_candidate = tire_candidate
    group.excluded_candidates = [candidate for candidate in unique_candidates if candidate is not tire_candidate]

    if len(unique_candidates) > 1:
        for candidate in group.excluded_candidates:
            cog_distance = distance(tire_candidate.component_cog_world, candidate.component_cog_world)
            if cog_distance > tire_hub_cog_tolerance:
                group.warnings.append(
                    f"{candidate.context.component_path}: 与 Tire 代表件重心距离 {cog_distance:.3f} "
                    f"超过轮胎/轮毂重心容差 {tire_hub_cog_tolerance:.3f}。"
                )
            if (
                tire_candidate.bbox_world is not None
                and candidate.bbox_world is not None
                and tire_candidate.bbox_world.diagonal
                < candidate.bbox_world.diagonal * TIRE_SIZE_ADVANTAGE_RATIO
            ):
                group.warnings.append(
                    f"{candidate.context.component_path}: 包围盒尺寸接近或大于 Tire 代表件，"
                    "轮胎/轮毂区分可能不稳定。"
                )


def choose_tire_representatives(
    groups: list[WheelPositionGroup],
    tire_hub_cog_tolerance: float,
) -> list[WheelCandidate]:
    """
    功能: 为所有位置组选择 Tire 代表件。
    输入: 位置组列表和重心容差。
    输出: Tire 代表件列表。
    """
    representatives: list[WheelCandidate] = []
    for group in groups:
        choose_group_tire_candidate(group, tire_hub_cog_tolerance)
        if group.tire_candidate is not None:
            representatives.append(group.tire_candidate)
    return representatives


def wheel_position_group_to_dict(group: WheelPositionGroup) -> dict[str, Any]:
    """
    功能: 将车轮位置组转为结果字典。
    输入: WheelPositionGroup。
    输出: JSON 友好的字典。
    """
    return {
        "group_id": group.group_id,
        "candidate_count": len(group.candidates),
        "tire_candidate": candidate_to_dict(group.tire_candidate) if group.tire_candidate else None,
        "excluded_candidates": [candidate_to_dict(candidate) for candidate in group.excluded_candidates],
        "candidates": [candidate_to_dict(candidate) for candidate in group.candidates],
        "warnings": group.warnings,
    }


def set_product_visibility(document: Any, product: Any, visible: bool) -> None:
    """
    功能: 设置 Product 显示或隐藏。
    输入: Document、Product 和是否显示。
    输出: 无。
    """
    selection = document.Selection
    try:
        selection.Clear()
        selection.Add(product)
        selection.VisProperties.SetShow(0 if visible else 1)
    finally:
        selection.Clear()


def set_object_visibility(document: Any, feature: Any, visible: bool) -> None:
    """
    功能: 设置单个 CATIA 几何对象显示或隐藏。
    输入: Document、几何对象和是否显示。
    输出: 无。
    """
    selection = document.Selection
    try:
        selection.Clear()
        selection.Add(feature)
        selection.VisProperties.SetShow(0 if visible else 1)
    finally:
        selection.Clear()


def get_object_visibility(document: Any, feature: Any) -> dict[str, Any]:
    """
    功能: 读取单个 CATIA 几何对象的显示/隐藏状态。
    输入: Document 和几何对象。
    输出: 显隐状态字典，visible/hidden/unknown。
    """
    selection = document.Selection
    try:
        selection.Clear()
        selection.Add(feature)
        raw_show = selection.VisProperties.GetShow()
        if isinstance(raw_show, (tuple, list)) and raw_show:
            show_value = int(raw_show[-1])
        else:
            show_value = int(raw_show)
        if show_value == 1:
            return {"status": "hidden", "visible": False, "raw_show": show_value}
        if show_value == 0:
            return {"status": "visible", "visible": True, "raw_show": show_value}
        return {"status": "unknown", "visible": True, "raw_show": show_value}
    except Exception as exc:
        return {
            "status": "unknown",
            "visible": True,
            "message": str(exc),
        }
    finally:
        try:
            selection.Clear()
        except Exception:
            pass


def apply_wheel_part_visibility(
    product_document: Any,
    wheel_part_contexts: list[WheelPartContext],
    visible_component_paths: set[str],
) -> list[dict[str, Any]]:
    """
    功能: 按路径集合批量设置车轮零件显示状态。
    输入: ProductDocument、上下文列表和可见路径集合。
    输出: 显示状态列表。
    """
    visibility_rows: list[dict[str, Any]] = []
    for context in wheel_part_contexts:
        visible = context.component_path in visible_component_paths
        try:
            set_product_visibility(product_document, context.product, visible)
            status = "visible" if visible else "hidden"
            error = None
        except Exception as exc:
            status = "failed"
            error = str(exc)
        visibility_rows.append(
            {
                "component_path": context.component_path,
                "component_name": context.component_name,
                "component_part_number": context.component_part_number,
                "visible": visible,
                "status": status,
                "error": error,
            }
        )
    return visibility_rows


def hide_wheel_assembly_after_regulation_geometry(
    product_document: Any,
    wheel_assembly_component: Any,
) -> dict[str, Any]:
    """
    功能: 法规校核辅助几何创建完成后隐藏整套车轮装配。
    输入: ProductDocument 和车轮装配组件。
    输出: 显示状态结果字典。
    """
    row = {
        "component_path": WHEEL_ASSEMBLY_LABEL,
        "component_name": product_display_name(wheel_assembly_component),
        "component_part_number": product_part_number(wheel_assembly_component),
        "visible": False,
        "status": "hidden",
        "error": None,
    }
    try:
        set_product_visibility(product_document, wheel_assembly_component, False)
    except Exception as exc:
        row["status"] = "failed"
        row["error"] = str(exc)
    return row


def visible_contexts(
    wheel_part_contexts: list[WheelPartContext],
    visible_component_paths: set[str],
) -> list[WheelPartContext]:
    """
    功能: 筛选当前应显示的车轮零件上下文。
    输入: 上下文列表和可见路径集合。
    输出: 过滤后的上下文列表。
    """
    return [
        context
        for context in wheel_part_contexts
        if context.component_path in visible_component_paths
    ]


def product_cog_row(document: Any, context: WheelPartContext, label: str | None = None) -> dict[str, Any]:
    """
    功能: 读取组件重心并构造结果行。
    输入: Document、上下文和标签。
    输出: 包含重心的字典。
    """
    row_label = label or context.component_path
    cog = evaluate_product_cog(document, context.product, row_label)
    return {
        "label": row_label,
        "component_path": context.component_path,
        "component_name": context.component_name,
        "component_part_number": context.component_part_number,
        "cog": tuple(round(value, 6) for value in cog),
        "_raw_cog": cog,
    }


def component_context_from_product(product: Any, component_path: str) -> WheelPartContext:
    """
    功能: 根据 Product 构造临时上下文。
    输入: Product 和组件路径。
    输出: WheelPartContext。
    """
    part, document = get_part_and_document_from_product(product)
    return WheelPartContext(
        product=product,
        part=part,
        document=document,
        product_chain=[product],
        component_path=component_path,
        component_name=product_display_name(product),
        component_part_number=product_part_number(product),
    )


def match_wheelhouses_to_visible_wheels(
    product_document: Any,
    front_component: Any,
    rear_component: Any,
    wheel_contexts: list[WheelPartContext],
) -> tuple[list[dict[str, Any]], list[dict[str, Any]], list[dict[str, Any]]]:
    """
    功能: 旧版用显示车轮重心匹配轮罩。
    输入: ProductDocument、前后轮罩组件和车轮上下文。
    输出: 匹配、轮罩重心和车轮重心。
    """
    front_context = component_context_from_product(front_component, FRONT_WHEELHOUSE_LABEL)
    rear_context = component_context_from_product(rear_component, REAR_WHEELHOUSE_LABEL)
    cover_contexts = [
        (FRONT_WHEELHOUSE_LABEL, front_context),
        (REAR_WHEELHOUSE_LABEL, rear_context),
    ]

    cover_rows = [
        product_cog_row(product_document, front_context, FRONT_WHEELHOUSE_LABEL),
        product_cog_row(product_document, rear_context, REAR_WHEELHOUSE_LABEL),
    ]
    wheel_rows = [
        product_cog_row(product_document, context, context.component_path)
        for context in wheel_contexts
    ]

    if not wheel_rows:
        raise RuntimeError("车轮装配中没有可用于匹配的显示零件。")

    matches: list[dict[str, Any]] = []
    if len(wheel_rows) >= len(cover_rows):
        best_total = math.inf
        best_pairs: tuple[int, ...] | None = None
        for wheel_indexes in itertools.permutations(range(len(wheel_rows)), len(cover_rows)):
            total = 0.0
            for cover_index, wheel_index in enumerate(wheel_indexes):
                total += distance(
                    cover_rows[cover_index]["_raw_cog"],
                    wheel_rows[wheel_index]["_raw_cog"],
                )
            if total < best_total:
                best_total = total
                best_pairs = wheel_indexes

        if best_pairs is not None:
            for cover_index, wheel_index in enumerate(best_pairs):
                cover_label, _cover_context = cover_contexts[cover_index]
                wheel_row = wheel_rows[wheel_index]
                matches.append(
                    build_match_row(
                        cover_label,
                        cover_rows[cover_index],
                        wheel_row,
                    )
                )
    else:
        for cover_index, (cover_label, _cover_context) in enumerate(cover_contexts):
            nearest_wheel = min(
                wheel_rows,
                key=lambda row: distance(cover_rows[cover_index]["_raw_cog"], row["_raw_cog"]),
            )
            matches.append(build_match_row(cover_label, cover_rows[cover_index], nearest_wheel))

    for row in cover_rows + wheel_rows:
        row.pop("_raw_cog", None)
    return matches, cover_rows, wheel_rows


def build_match_row(cover_label: str, cover_row: dict[str, Any], wheel_row: dict[str, Any]) -> dict[str, Any]:
    """
    功能: 构造旧版车轮罩匹配结果。
    输入: 轮罩标签、轮罩行和车轮行。
    输出: 匹配字典。
    """
    match_distance = distance(cover_row["_raw_cog"], wheel_row["_raw_cog"])
    return {
        "wheelhouse": cover_label,
        "wheelhouse_cog": cover_row["cog"],
        "wheel_component_path": wheel_row["component_path"],
        "wheel_component_name": wheel_row["component_name"],
        "wheel_component_part_number": wheel_row["component_part_number"],
        "wheel_cog": wheel_row["cog"],
        "distance": round(match_distance, 6),
    }


def tire_match_score(cover_cog: Vector, tire_candidate: WheelCandidate) -> tuple[float, dict[str, float]]:
    """
    功能: 计算轮罩与 Tire 代表件的匹配分数。
    输入: 轮罩重心和 Tire 候选。
    输出: 总分和分项距离。
    """
    cog_distance = distance(cover_cog, tire_candidate.component_cog_world)
    axis_distance = 0.0
    axis_point_distance = 0.0
    if tire_candidate.axis_direction_world is not None and tire_candidate.axis_point_world is not None:
        axis_distance = point_to_axis_distance(
            cover_cog,
            tire_candidate.axis_point_world,
            tire_candidate.axis_direction_world,
        )
        axis_point_distance = distance(cover_cog, tire_candidate.axis_point_world)
    score = 0.65 * cog_distance + 0.25 * axis_distance + 0.10 * axis_point_distance
    return score, {
        "cog_distance": round(cog_distance, 6),
        "axis_distance": round(axis_distance, 6),
        "axis_point_distance": round(axis_point_distance, 6),
        "score": round(score, 6),
    }


def build_tire_match_row(
    wheelhouse_label: str,
    cover_cog: Vector,
    tire_candidate: WheelCandidate,
) -> dict[str, Any]:
    """
    功能: 构造轮罩到 Tire 的匹配结果。
    输入: 轮罩标签、轮罩重心和 Tire 候选。
    输出: 匹配字典。
    """
    _score, score_detail = tire_match_score(cover_cog, tire_candidate)
    return {
        "wheelhouse": wheelhouse_label,
        "wheelhouse_cog": round_vector(cover_cog),
        "tire_component_path": tire_candidate.context.component_path,
        "tire_component_name": tire_candidate.context.component_name,
        "tire_component_part_number": tire_candidate.context.component_part_number,
        "tire_cog": round_vector(tire_candidate.component_cog_world),
        "tire_axis_direction_world": round_vector(tire_candidate.axis_direction_world, 8),
        "tire_axis_point_world": round_vector(tire_candidate.axis_point_world),
        "tire_bbox_world": bounding_box_to_dict(tire_candidate.bbox_world),
        **score_detail,
    }


def evaluate_wheelhouse_cog_items(
    product_document: Any,
    front_component: Any,
    rear_component: Any | None = None,
) -> list[tuple[str, Vector]]:
    """
    功能: 读取前后轮罩重心，供近邻筛选和 Tire 匹配复用。
    输入: ProductDocument、前轮罩组件和后轮罩组件。
    输出: (轮罩标签, 世界坐标重心) 列表。
    """
    if isinstance(front_component, dict):
        component_items = list(front_component.items())
    else:
        component_items = [
            (FRONT_WHEELHOUSE_LABEL, front_component),
            (REAR_WHEELHOUSE_LABEL, rear_component),
        ]
    results: list[tuple[str, Vector]] = []
    for label, component in component_items:
        if component is None:
            continue
        context = component_context_from_product(component, str(label))
        results.append((str(label), evaluate_product_cog(product_document, context.product, str(label))))
    return results


def match_wheelhouses_to_tire_representatives(
    product_document: Any,
    front_component: Any,
    rear_component: Any | None = None,
    tire_candidates: list[WheelCandidate] | None = None,
    wheelhouse_cog_items: list[tuple[str, Vector]] | None = None,
) -> tuple[list[dict[str, Any]], list[dict[str, Any]], list[dict[str, Any]], list[str]]:
    """
    功能: 将前后轮罩匹配到 Tire 代表件。
    输入: ProductDocument、前后轮罩组件和 Tire 候选列表。
    输出: 匹配结果、轮罩重心、Tire 数据和告警。
    """
    warnings: list[str] = []
    if tire_candidates is None:
        raise RuntimeError("未提供 Tire 候选列表。")
    cover_items = wheelhouse_cog_items or evaluate_wheelhouse_cog_items(product_document, front_component, rear_component)
    cover_rows = [
        {
            "label": label,
            "component_path": label,
            "component_name": label,
            "component_part_number": label,
            "cog": round_vector(cog),
        }
        for label, cog in cover_items
    ]
    tire_rows = [candidate_to_dict(candidate) for candidate in tire_candidates]

    if not tire_candidates:
        raise RuntimeError("车轮装配中没有可用于匹配的 Tire 代表件。")

    matches: list[dict[str, Any]] = []
    label_candidates: dict[str, list[WheelCandidate]] = {}
    for candidate in tire_candidates:
        if candidate.wheel_group_key:
            label_candidates.setdefault(candidate.wheel_group_key, []).append(candidate)
    if all(label in label_candidates for label, _cover_cog in cover_items):
        for label, cover_cog in cover_items:
            selected_candidate = max(
                label_candidates[label],
                key=tire_priority_score,
            )
            matches.append(build_tire_match_row(label, cover_cog, selected_candidate))
        return matches, cover_rows, tire_rows, warnings

    if len(tire_candidates) >= len(cover_items):
        best_total = math.inf
        best_pairs: tuple[int, ...] | None = None
        for tire_indexes in itertools.permutations(range(len(tire_candidates)), len(cover_items)):
            total = 0.0
            for cover_index, tire_index in enumerate(tire_indexes):
                score, _score_detail = tire_match_score(
                    cover_items[cover_index][1],
                    tire_candidates[tire_index],
                )
                total += score
            if total < best_total:
                best_total = total
                best_pairs = tire_indexes
        if best_pairs is not None:
            for cover_index, tire_index in enumerate(best_pairs):
                label, cover_cog = cover_items[cover_index]
                matches.append(build_tire_match_row(label, cover_cog, tire_candidates[tire_index]))
    else:
        warnings.append("Tire 代表件数量不足，允许多个轮罩复用最近 Tire 候选。")
        for label, cover_cog in cover_items:
            nearest_tire = min(tire_candidates, key=lambda candidate: tire_match_score(cover_cog, candidate)[0])
            matches.append(build_tire_match_row(label, cover_cog, nearest_tire))

    return matches, cover_rows, tire_rows, warnings


def bounding_box_corners(bbox: BoundingBox) -> list[Vector]:
    """
    功能: 生成包围盒 8 个角点。
    输入: BoundingBox。
    输出: 角点坐标列表。
    """
    min_x, min_y, min_z = bbox.min_point
    max_x, max_y, max_z = bbox.max_point
    return [
        (x, y, z)
        for x in (min_x, max_x)
        for y in (min_y, max_y)
        for z in (min_z, max_z)
    ]


def bbox_length_along_axis(bbox: BoundingBox, axis_direction: Vector) -> float:
    """
    功能: 计算包围盒沿指定轴方向的投影长度。
    输入: 包围盒和世界坐标轴方向。
    输出: 投影长度。
    """
    direction = normalize_vector(axis_direction)
    projections = [dot_product(corner, direction) for corner in bounding_box_corners(bbox)]
    return max(projections) - min(projections)


def bbox_projection_range_along_axis(bbox: BoundingBox, axis_direction: Vector) -> tuple[float, float]:
    """
    功能: 计算包围盒角点沿指定轴方向的投影范围。
    输入: 包围盒和世界坐标轴方向。
    输出: 最小投影值和最大投影值。
    """
    direction = normalize_vector(axis_direction)
    projections = [dot_product(corner, direction) for corner in bounding_box_corners(bbox)]
    return min(projections), max(projections)


def tire_candidate_error_label(candidate: WheelCandidate) -> str:
    """
    功能: 生成用于报错定位的 Tire 候选零件说明。
    输入: WheelCandidate。
    输出: 包含路径、名称、零件编号和轴来源的文本。
    """
    return (
        f"component_path={candidate.context.component_path}, "
        f"component_name={candidate.context.component_name}, "
        f"component_part_number={candidate.context.component_part_number}, "
        f"axis_source={candidate.axis_source}, "
        f"feature_name={candidate.feature_name}"
    )


def build_regulation_axis_segment(
    wheelhouse_label: str,
    tire_candidate: WheelCandidate,
    half_length: float = REGULATION_AXIS_HALF_LENGTH,
    extrude_length: float = REGULATION_AXIS_EXTRUDE_LENGTH,
    rotation_angle: float = REGULATION_AXIS_ROTATION_ANGLE,
) -> RegulationAxisSegment:
    """
    功能: 根据 Tire 轴线和 Tire 重心在轴上的投影点计算法规校核轴线段。
    输入: 轮罩标签、Tire 候选、轴线两侧延伸长度、拉伸长度和旋转角度。
    输出: RegulationAxisSegment。
    """
    if tire_candidate.axis_direction_world is None or tire_candidate.axis_point_world is None:
        raise RuntimeError(
            f"{wheelhouse_label}: 匹配 Tire 缺少世界坐标轴线，"
            f"{tire_candidate_error_label(tire_candidate)}。"
        )
    direction = normalized_axis_direction(tire_candidate.axis_direction_world)
    axis_point = tire_candidate.axis_point_world
    cog_to_axis_offset = dot_product(subtract_vectors(tire_candidate.component_cog_world, axis_point), direction)
    center_point = add_vectors(axis_point, scale_vector(direction, cog_to_axis_offset))
    half_vector = scale_vector(direction, half_length)
    start_point = subtract_vectors(center_point, half_vector)
    end_point = add_vectors(center_point, half_vector)
    axis_length = half_length * 2.0
    definition = wheelhouse_definition_from_label(wheelhouse_label)
    geometry_set_name = definition["geometry_set_name"]
    line_name = definition["line_name"]
    return RegulationAxisSegment(
        wheelhouse_label=wheelhouse_label,
        geometry_set_name=geometry_set_name,
        line_name=line_name,
        start_point=start_point,
        end_point=end_point,
        axis_direction_world=direction,
        axis_point_world=axis_point,
        center_point_world=center_point,
        axis_length=axis_length,
        half_length=half_length,
        extrude_length=extrude_length,
        rotation_angle=rotation_angle,
        tire_component_path=tire_candidate.context.component_path,
    )


def regulation_axis_segment_to_dict(segment: RegulationAxisSegment) -> dict[str, Any]:
    """
    功能: 将法规校核轴线段转为结果字典。
    输入: RegulationAxisSegment。
    输出: JSON 友好的字典。
    """
    return {
        "wheelhouse": segment.wheelhouse_label,
        "geometry_set_name": segment.geometry_set_name,
        "line_name": segment.line_name,
        "start_point": round_vector(segment.start_point),
        "end_point": round_vector(segment.end_point),
        "axis_direction_world": round_vector(segment.axis_direction_world, 8),
        "axis_point_world": round_vector(segment.axis_point_world),
        "center_point_world": round_vector(segment.center_point_world),
        "axis_length": round(segment.axis_length, 6),
        "half_length": round(segment.half_length, 6),
        "extrude_length": round(segment.extrude_length, 6),
        "rotation_angle": round(segment.rotation_angle, 6),
        "tire_component_path": segment.tire_component_path,
    }


def add_new_part_to_product(root_product: Any, part_number: str, name: str) -> Any:
    """
    功能: 在根 Product 下新建一个 CATPart 组件。
    输入: 根 Product、英文 PartNumber 和英文名称。
    输出: 新增的 Part Product。
    """
    try:
        new_product = root_product.Products.AddNewComponent("Part", part_number)
    except Exception as exc:
        raise RuntimeError(f"无法新建 CATPart: {part_number}") from exc
    set_if_possible(new_product, "PartNumber", part_number)
    set_if_possible(new_product, "Name", name)
    return new_product


def add_saved_part_file_to_product(
    product_document: Any,
    root_product: Any,
    part_number: str,
    name: str,
    save_path: Path,
) -> tuple[Any, dict[str, Any]]:
    """
    功能: 先创建并保存独立 CATPart 文件，再把该文件装配到根 Product。
    输入: ProductDocument、根 Product、PartNumber、显示名称和 CATPart 保存路径。
    输出: (新增组件 Product, 初始 CATPart 保存结果)。
    """
    save_path = Path(save_path)
    save_path.parent.mkdir(parents=True, exist_ok=True)
    catia = product_document.Application
    part_document = None
    save_result: dict[str, Any] = {
        "status": "failed",
        "document": part_number,
        "path": str(save_path),
    }
    try:
        part_document = catia.Documents.Add("Part")
        part = part_document.Part
        set_if_possible(part, "PartNumber", part_number)
        try:
            set_if_possible(part, "Name", name)
        except Exception:
            pass
        try:
            part.Update()
        except Exception:
            pass
        part_document.SaveAs(str(save_path))
        save_result = {
            "status": "success",
            "message": "独立 CATPart 已预创建并保存。",
            "document": safe_attr_text(part_document, "Name", part_number),
            "path": str(save_path),
        }
    except Exception as exc:
        raise RuntimeError(f"无法预创建并保存 CATPart: {save_path}") from exc
    finally:
        if part_document is not None:
            try:
                part_document.Close()
            except Exception:
                pass

    new_product = add_component_from_file_to_product(
        product_document,
        root_product,
        save_path,
    )
    set_if_possible(new_product, "PartNumber", part_number)
    set_if_possible(new_product, "Name", name)
    return new_product, save_result


def append_axis_line_to_hybrid_body(
    document: Any,
    part: Any,
    hybrid_body: Any,
    segment: RegulationAxisSegment,
) -> dict[str, str]:
    """
    功能: 在指定几何图形集中创建轴线段、Z向拉伸、绕轴旋转和截面平面。
    输入: Document、Part、HybridBody 和轴线段数据。
    输出: 创建的特征名称字典。
    """
    factory = part.HybridShapeFactory
    start_point = factory.AddNewPointCoord(*segment.start_point)
    end_point = factory.AddNewPointCoord(*segment.end_point)
    set_if_possible(start_point, "Name", f"{segment.line_name}_Start")
    set_if_possible(end_point, "Name", f"{segment.line_name}_End")
    hybrid_body.AppendHybridShape(start_point)
    hybrid_body.AppendHybridShape(end_point)
    line = factory.AddNewLinePtPt(
        create_reference(part, start_point),
        create_reference(part, end_point),
    )
    set_if_possible(line, "Name", segment.line_name)
    hybrid_body.AppendHybridShape(line)
    z_direction = factory.AddNewDirectionByCoord(0.0, 0.0, 1.0)
    extrude = factory.AddNewExtrude(
        create_reference(part, line),
        0.0,
        # CATIA 该接口的偏移方向与 DirectionByCoord 表现相反，负值对应实际 Z 正向。
        -segment.extrude_length,
        z_direction,
    )
    extrude_name = f"{segment.line_name}_Z_Extrude"
    set_if_possible(extrude, "Name", extrude_name)
    hybrid_body.AppendHybridShape(extrude)
    rotation = factory.AddNewRotate(
        create_reference(part, extrude),
        create_reference(part, line),
        segment.rotation_angle,
    )
    rotation_name = f"{segment.line_name}_Rotation"
    set_if_possible(rotation, "Name", rotation_name)
    hybrid_body.AppendHybridShape(rotation)
    try:
        set_object_visibility(document, extrude, False)
        set_object_visibility(document, rotation, False)
    except Exception:
        pass

    z_offset = (0.0, 0.0, segment.extrude_length)
    section_0_point = factory.AddNewPointCoord(*add_vectors(segment.center_point_world, z_offset))
    section_30_offset = rotate_vector_around_axis(
        z_offset,
        segment.axis_direction_world,
        segment.rotation_angle,
    )
    section_30_point = factory.AddNewPointCoord(*add_vectors(segment.center_point_world, section_30_offset))
    set_if_possible(section_0_point, "Name", f"{segment.line_name}_Section_0_Point")
    set_if_possible(section_30_point, "Name", f"{segment.line_name}_Section_30_Point")
    hybrid_body.AppendHybridShape(section_0_point)
    hybrid_body.AppendHybridShape(section_30_point)

    section_prefix = wheelhouse_section_prefix(segment.wheelhouse_label)
    section_0_name = f"{section_prefix}截面0°"
    section_30_name = f"{section_prefix}截面30°"
    section_0_plane = factory.AddNewPlane3Points(
        create_reference(part, start_point),
        create_reference(part, end_point),
        create_reference(part, section_0_point),
    )
    section_30_plane = factory.AddNewPlane3Points(
        create_reference(part, start_point),
        create_reference(part, end_point),
        create_reference(part, section_30_point),
    )
    set_if_possible(section_0_plane, "Name", section_0_name)
    set_if_possible(section_30_plane, "Name", section_30_name)
    hybrid_body.AppendHybridShape(section_0_plane)
    hybrid_body.AppendHybridShape(section_30_plane)
    return {
        "start_point_name": f"{segment.line_name}_Start",
        "end_point_name": f"{segment.line_name}_End",
        "line_name": segment.line_name,
        "extrude_name": extrude_name,
        "rotation_name": rotation_name,
        "extrude_hidden": "true",
        "rotation_hidden": "true",
        "section_0_point_name": f"{segment.line_name}_Section_0_Point",
        "section_30_point_name": f"{segment.line_name}_Section_30_Point",
        "section_0_plane_name": section_0_name,
        "section_30_plane_name": section_30_name,
        "start_point_world": segment.start_point,
        "end_point_world": segment.end_point,
        "axis_point_world": segment.axis_point_world,
        "axis_direction_world": segment.axis_direction_world,
        "section_0_point_world": add_vectors(segment.center_point_world, z_offset),
        "section_30_point_world": add_vectors(segment.center_point_world, section_30_offset),
    }


def create_section_curves_from_regulation_planes(
    product_document: Any,
    root_product: Any,
    wheelhouse_components: dict[str, Any],
    process_part: Any,
    process_hybrid_bodies: dict[str, Any],
    process_transform: Transform | None,
    created_features: list[dict[str, Any]],
    output_dir: str | Path | None = None,
    course_dir: str | Path | None = None,
    run_timestamp: str | None = None,
    annotation_runtime_context: dict[str, Any] | None = None,
) -> list[dict[str, Any]]:
    """
    功能: 通过独立截面导出工具为前后轮罩创建四个截面结果 CATPart。
    输入: ProductDocument、根 Product、前后轮罩组件映射、过程 Part 信息、过程 Part 装配变换、已创建特征列表和输出目录。
    输出: 截面导出和装配结果列表。
    """
    results: list[dict[str, Any]] = []
    try:
        tool_path = resolve_section_curve_export_tool_path()
        spec = importlib.util.spec_from_file_location("section_curve_export_tool_embedded", tool_path)
        if spec is None or spec.loader is None:
            raise RuntimeError(f"无法加载工具文件: {tool_path}")
        section_tool = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(section_tool)
    except Exception as exc:
        return [
            {
                "status": "failed",
                "message": f"无法加载 section_curve_export_tool: {exc}",
            }
        ]

    export_root = Path(output_dir).expanduser().resolve() if output_dir else DEFAULT_OUTPUT_DIR
    export_root = export_root / "sec"
    export_root.mkdir(parents=True, exist_ok=True)
    section_work_dir = Path(course_dir).expanduser().resolve() if course_dir else export_root.parent / "section_work"
    section_work_dir.mkdir(parents=True, exist_ok=True)
    topology_curve_rows: list[dict[str, Any]] = []
    axis_clearance_results: list[dict[str, Any]] = []

    for feature_row in created_features:
        wheelhouse_label = feature_row.get("wheelhouse")
        wheelhouse_component = wheelhouse_components.get(str(wheelhouse_label))
        if wheelhouse_component is None:
            results.append(
                {
                    "status": "failed",
                    "message": f"未找到 {wheelhouse_label} 对应的轮罩 CATPart 组件。",
                    "feature_row": feature_row,
                }
            )
            continue
        section_document = None
        try:
            wheelhouse_part, wheelhouse_document = get_part_and_document_from_product(wheelhouse_component)
            wheelhouse_transform = product_position_transform(product_document, wheelhouse_component)
            wheelhouse_file_text = safe_attr_text(wheelhouse_document, "FullName")
            if not wheelhouse_file_text:
                raise RuntimeError(f"{wheelhouse_label}: 无法获取轮罩工作 CATPart 文件路径。")
            section_source_path = copy_part_to_section_work(
                Path(wheelhouse_file_text),
                section_work_dir,
                str(wheelhouse_label),
                timestamp=run_timestamp,
            )
            section_document = product_document.Application.Documents.Open(str(section_source_path))
            section_part = section_document.Part
            try:
                section_document.Activate()
            except Exception:
                pass
            body_candidates, surface_candidates = scan_section_target_objects(section_document, section_part, surface_limit=100)
            print(f"\n-- {wheelhouse_label} 截面目标扫描 --")
            print(f"轮罩文档: {safe_attr_text(section_document, 'Name', '<unknown>')}")
            print(f"轮罩零件: {safe_attr_text(section_part, 'Name', '<unknown>')}")
            print(f"截面工作文件: {section_source_path}")
            if body_candidates:
                print("Body 目标:")
                for target_row in body_candidates:
                    visibility = target_row.get("visibility") or {}
                    print(
                        f"  - {target_row.get('label') or target_row.get('name')} "
                        f"visibility={visibility.get('status', 'unknown')}"
                    )
            else:
                print("Body 目标: 未找到")
            if surface_candidates:
                print("Surface 目标:")
                for target_row in surface_candidates:
                    visibility = target_row.get("visibility") or {}
                    print(
                        f"  - {target_row.get('label') or target_row.get('name')} "
                        f"visibility={visibility.get('status', 'unknown')}"
                    )
            else:
                print("Surface 目标: 未找到")
            raw_target_rows = body_candidates if body_candidates else surface_candidates
            hidden_target_rows = [
                row for row in raw_target_rows
                if (row.get("visibility") or {}).get("status") == "hidden"
            ]
            target_rows = [
                row for row in raw_target_rows
                if (row.get("visibility") or {}).get("status") != "hidden"
            ]
            body_clearance_rows = [
                row for row in body_candidates
                if (row.get("visibility") or {}).get("status") != "hidden"
            ]
            target_labels = [str(row.get("label") or row.get("name")) for row in target_rows if row.get("name")]
            if not target_rows:
                raise RuntimeError(f"{wheelhouse_label}: 未找到显示状态可用于截面的 Body 或曲面。")
            if hidden_target_rows:
                print("隐藏目标已跳过:")
                for target_row in hidden_target_rows:
                    print(f"  - {target_row.get('label') or target_row.get('name')}")
            print("本次用于截面的目标:")
            for target_label in target_labels:
                print(f"  - {target_label}")
            target_name_counts: dict[str, int] = {}
            for target_row in target_rows:
                target_name = str(target_row.get("name") or "")
                target_name_counts[target_name] = target_name_counts.get(target_name, 0) + 1
            duplicate_target_names = [
                name for name, count in target_name_counts.items()
                if name and count > 1
            ]
            if duplicate_target_names:
                print("[警告] 存在同名截面目标，独立工具按名称解析时可能命中同名目标中的第一个:")
                for target_name in duplicate_target_names:
                    print(f"  - {target_name}")
            process_hybrid_body = process_hybrid_bodies.get(str(feature_row.get("geometry_set_name")))
            if process_hybrid_body is not None and body_clearance_rows:
                axis_clearance_result = measure_wheelhouse_to_axis_clearance(
                    product_document,
                    process_part,
                    process_hybrid_body,
                    process_transform,
                    section_document,
                    section_part,
                    wheelhouse_transform,
                    body_clearance_rows,
                    feature_row,
                )
                axis_clearance_results.append(axis_clearance_result)
            elif process_hybrid_body is not None:
                axis_clearance_results.append(
                    {
                        "status": "failed",
                        "measurement_key": wheelhouse_measurement_key(wheelhouse_label, "c"),
                        "message": f"{wheelhouse_label}: 未找到显示状态可用于测距的 Body。",
                    }
                )
            else:
                axis_clearance_results.append(
                    {
                        "status": "failed",
                        "measurement_key": wheelhouse_measurement_key(wheelhouse_label, "c"),
                        "message": f"未找到过程几何图形集: {feature_row.get('geometry_set_name')}",
                    }
                )
        except Exception as exc:
            results.append(
                {
                    "status": "failed",
                    "message": str(exc),
                    "feature_row": feature_row,
                }
            )
            continue
        for plane_key, angle_suffix in (
            ("section_0_plane_name", "0deg"),
            ("section_30_plane_name", "30deg"),
        ):
            plane_name = feature_row.get(plane_key)
            third_point_key = "section_0_point_world" if plane_key == "section_0_plane_name" else "section_30_point_world"
            if not plane_name:
                results.append(
                    {
                        "status": "failed",
                        "message": f"缺少 {plane_key}，无法创建截面曲线。",
                        "feature_row": feature_row,
                    }
                )
                continue
            try:
                first_local = inverse_transform_point(wheelhouse_transform, feature_row["start_point_world"])
                second_local = inverse_transform_point(wheelhouse_transform, feature_row["end_point_world"])
                third_local = inverse_transform_point(wheelhouse_transform, feature_row[third_point_key])
                equation = list(plane_equation_from_points(first_local, second_local, third_local))
                if plane_key == "section_30_plane_name":
                    extreme_world_direction = normalize_vector(
                        rotate_vector_around_axis(
                            (0.0, 0.0, 1.0),
                            feature_row["axis_direction_world"],
                            REGULATION_AXIS_ROTATION_ANGLE,
                        )
                    )
                else:
                    extreme_world_direction = (0.0, 0.0, 1.0)
                for target_index, target_row in enumerate(target_rows, start=1):
                    target_name = str(target_row.get("name") or "").strip()
                    target_label = str(target_row.get("label") or target_name or f"target#{target_index}")
                    target_visibility = target_row.get("visibility") or {}
                    if not target_name:
                        results.append(
                            {
                                "status": "failed",
                                "message": "截面目标缺少可传入独立工具的名称。",
                                "wheelhouse": wheelhouse_label,
                                "section_plane_name": plane_name,
                                "target_label": target_label,
                                "target_index": target_index,
                            }
                        )
                        continue
                    run_id = compact_section_run_id(
                        wheelhouse_label,
                        angle_suffix,
                        target_index,
                        target_name,
                        timestamp=run_timestamp,
                    )
                    unique_target_name = f"__main_target_{target_index:03d}_{short_stable_token(target_label)}"
                    run_output_dir = export_root / run_id
                    try:
                        section_document.Activate()
                    except Exception:
                        pass
                    print(f"\n-- 创建截面并导出CATPart: {plane_name} / {target_label} --")
                    print(f"轮罩文档: {safe_attr_text(section_document, 'Name', '<unknown>')}")
                    print(f"目标名称: {target_name}")
                    print(f"目标唯一键: {unique_target_name}")
                    print(f"平面方程: {format_plane_equation_for_cli(equation)}")
                    section_tool.get_active_catia_part = lambda: (
                        section_document.Application,
                        section_document,
                        section_part,
                    )
                    target_candidates_hook, reference_hook = make_embedded_section_target_hooks(
                        section_tool,
                        section_part,
                        target_row,
                        unique_target_name,
                    )
                    section_tool.target_candidates = target_candidates_hook
                    section_tool.reference_from_any_name = reference_hook
                    tool_args = argparse.Namespace(
                        target_name=unique_target_name,
                        section_plane="auto",
                        offset_distance=None,
                        reverse=False,
                        through_point=None,
                        normal=None,
                        plane_equation=format_plane_equation_for_cli(equation),
                        axis_name=None,
                        angle_deg=None,
                        angle_reverse=False,
                        run_id=run_id,
                        output_dir=str(run_output_dir),
                        min_length=0.001,
                        surface_limit=100,
                        work_body="wheelhouse_regulation_section_work",
                        export_body="SectionResult",
                        extend_mode=False,
                        color=True,
                        no_export=False,
                        close_exported_document=True,
                        dry_run=False,
                        user_confirmed=True,
                    )
                    try:
                        report_payload = section_tool.run_export(tool_args)
                    except SystemExit as exc:
                        report_payload = {
                            "status": "failed",
                            "message": f"截面工具触发 SystemExit，已捕获并继续: {exc}",
                        }
                    except Exception as exc:
                        report_payload = {
                            "status": "failed",
                            "message": f"截面工具运行异常，已捕获并继续: {exc}",
                        }
                    report_path = run_output_dir / "section_curve_export_report.json"
                    if not report_path.is_file() and isinstance(report_payload, dict):
                        try:
                            report_path = Path(str(report_payload.get("report_json") or report_path))
                        except Exception:
                            pass
                    exported_path_text = str(report_payload.get("exported_catpart") or "")
                    exported_path = Path(exported_path_text) if exported_path_text else run_output_dir / "SR.CATPart"
                    exported_component_info: dict[str, Any] | None = None
                    exported_curve_rows: list[dict[str, Any]] = []
                    section_extreme_processing: dict[str, Any] = {}
                    if report_payload.get("status") == "success" and exported_path.is_file():
                        component_name = f"{run_id}_SR"
                        exported_document = None
                        try:
                            exported_document = product_document.Application.Documents.Open(str(exported_path))
                            try:
                                exported_document.Activate()
                            except Exception:
                                pass
                            exported_part = exported_document.Part
                            try:
                                set_if_possible(exported_part, "PartNumber", component_name)
                            except Exception:
                                pass
                            try:
                                set_if_possible(exported_part, "Name", component_name)
                            except Exception:
                                pass
                            try:
                                set_if_possible(exported_document.Product, "PartNumber", component_name)
                            except Exception:
                                pass
                            try:
                                set_if_possible(exported_document.Product, "Name", component_name)
                            except Exception:
                                pass
                            exported_curve_rows = collect_section_result_curve_rows_from_part_document(
                                exported_document,
                                exported_part,
                                component_name,
                                component_name,
                                extreme_world_direction,
                                transform=identity_transform(),
                                section_body_name=str(tool_args.export_body),
                                section_curve_name=(
                                    str(report_payload.get("section_curve_name")).strip()
                                    if report_payload.get("section_curve_name")
                                    else None
                                ),
                            )
                            try:
                                exported_part.Update()
                            except Exception:
                                pass
                            save_result = save_document_if_modified(exported_document, exported_path)
                            section_extreme_processing = {
                                "status": "success" if exported_curve_rows and save_result.get("status") != "failed" else "failed",
                                "document": safe_attr_text(exported_document, "Name", "<unknown>"),
                                "curve_count": len(exported_curve_rows),
                                "save_result": save_result,
                                "processing_context": "standalone_catpart_before_product_assembly",
                            }
                            if not exported_curve_rows:
                                raise RuntimeError("单独打开截面 CATPart 后未能生成截面极值点。")
                            if save_result.get("status") == "failed":
                                raise RuntimeError(f"截面极值点处理后保存失败: {save_result.get('message')}")
                        except Exception as exc:
                            section_extreme_processing = {
                                **section_extreme_processing,
                                "status": "failed",
                                "message": str(exc),
                                "processing_context": "standalone_catpart_before_product_assembly",
                            }
                        finally:
                            if exported_document is not None:
                                try:
                                    exported_document.Close()
                                except Exception as exc:
                                    section_extreme_processing["close_warning"] = str(exc)
                        if section_extreme_processing.get("status") == "success":
                            exported_component = add_component_from_file_to_product(
                                product_document,
                                root_product,
                                exported_path,
                            )
                            set_if_possible(exported_component, "Name", component_name)
                            set_if_possible(exported_component, "PartNumber", component_name)
                            exported_component_info = {
                                "component_name": product_display_name(exported_component),
                                "component_part_number": product_part_number(exported_component),
                            }
                            for curve_row in exported_curve_rows:
                                curve_row["component_name"] = product_display_name(exported_component)
                                curve_row["component_part_number"] = product_part_number(exported_component)
                                curve_row["wheelhouse"] = wheelhouse_label
                                curve_row["geometry_set_name"] = feature_row.get("geometry_set_name")
                                curve_row["section_plane_name"] = plane_name
                                curve_row["target_name"] = target_name
                                curve_row["target_label"] = target_label
                                curve_row["target_index"] = target_index
                                curve_row["run_id"] = run_id
                                curve_row["extreme_world_direction"] = extreme_world_direction
                            topology_curve_rows.extend(exported_curve_rows)
                            status = "success"
                            message = "截面目标结果 CATPart 已单独处理极值并装配到总成。"
                        else:
                            status = "failed"
                            message = str(section_extreme_processing.get("message") or "截面极值点处理失败。")
                    else:
                        status = "failed"
                        message = str(report_payload.get("message") or "独立截面工具执行失败。")
                    result = {
                        "status": status,
                        "message": message,
                        "wheelhouse": wheelhouse_label,
                        "geometry_set_name": feature_row.get("geometry_set_name"),
                        "wheelhouse_document": safe_attr_text(section_document, "Name", "<unknown>"),
                        "wheelhouse_part": safe_attr_text(section_part, "Name", "<unknown>"),
                        "section_source_path": str(section_source_path),
                        "source_wheelhouse_work_path": wheelhouse_file_text,
                        "section_plane_name": plane_name,
                        "section_curve_name": report_payload.get("section_curve_name"),
                        "target_name": target_name,
                        "tool_target_name": report_payload.get("target_name") or unique_target_name,
                        "target_label": target_label,
                        "target_index": target_index,
                        "target_visibility": target_visibility,
                        "selected_target": report_payload.get("selected_target"),
                        "run_id": run_id,
                        "run_output_dir": str(run_output_dir),
                        "tool_path": str(tool_path),
                        "tool_return_code": 0 if report_payload.get("status") == "success" else 1,
                        "tool_stdout": "",
                        "tool_stderr": "",
                        "plane_equation_local": equation,
                        "world_points": [
                            feature_row["start_point_world"],
                            feature_row["end_point_world"],
                            feature_row[third_point_key],
                        ],
                        "exported_catpart": str(exported_path) if exported_path.is_file() else None,
                        "exported_component": exported_component_info,
                        "section_extreme_processing": section_extreme_processing,
                        "topology_curve_count": len(exported_curve_rows) if status == "success" else 0,
                        "report_json": str(report_path) if report_path.is_file() else None,
                        "raw_report": report_payload,
                    }
                    results.append(result)
                    if status == "success":
                        print(f"截面导出装配成功: {plane_name} / {target_label} -> {exported_path}")
                    else:
                        print(f"截面导出装配失败: {plane_name} / {target_label}: {message}")
            except Exception as exc:
                results.append(
                    {
                        "status": "failed",
                        "message": f"截面工具运行失败，已继续后续目标: {exc}",
                        "wheelhouse": wheelhouse_label,
                        "section_plane_name": plane_name,
                        "feature_row": feature_row,
                    }
                )
        if section_document is not None:
            save_result = save_document_if_modified(section_document)
            if save_result.get("status") == "failed":
                print(f"[警告] 截面工作文件保存失败: {save_result.get('document')}: {save_result.get('message')}")
            try:
                section_document.Close()
            except Exception as exc:
                print(f"[警告] 截面工作文件关闭失败: {safe_attr_text(section_document, 'Name', '<unknown>')}: {exc}")

    if annotation_runtime_context is not None:
        annotation_runtime_context.setdefault("regulation_axis_part_info", {})
        annotation_runtime_context["regulation_axis_part_info"]["section_curve_results"] = results
        for clearance_result in axis_clearance_results:
            measurement_geometry = clearance_result.get("measurement_geometry") or {}
            axis_item = build_wheelhouse_annotation_item(
                clearance_result.get("measurement_key"),
                clearance_result.get("measurement_value"),
                measurement_geometry.get("body_point_world"),
                measurement_geometry.get("axis_point_world"),
                source="wheelhouse_to_axis_min_distance",
            )
            if axis_item is None:
                continue
            axis_item["wheelhouse"] = clearance_result.get("wheelhouse")
            runtime_capture_wheelhouse_annotation_item(
                product_document.Application,
                product_document,
                root_product,
                annotation_runtime_context,
                axis_item,
                category="axis_clearance",
                regulation_axis_part=annotation_runtime_context["regulation_axis_part_info"],
            )
    topology_group_results = analyze_section_result_topology_groups(
        topology_curve_rows,
    )
    regulation_distance_measurements = {
        str(group["measurement_key"]): group.get("measurement_value")
        for group in topology_group_results
        if group.get("measurement_key") and group.get("measurement_value") is not None
    }
    for clearance_result in axis_clearance_results:
        if clearance_result.get("measurement_key") and clearance_result.get("measurement_value") is not None:
            regulation_distance_measurements[str(clearance_result["measurement_key"])] = clearance_result.get("measurement_value")
    if annotation_runtime_context is not None:
        annotation_runtime_context.setdefault("regulation_axis_part_info", {})
        annotation_runtime_context["regulation_axis_part_info"]["section_curve_results"] = results
        for group in topology_group_results:
            section_item = build_wheelhouse_annotation_item(
                group.get("measurement_key"),
                group.get("measurement_value"),
                group.get("high_point_world"),
                group.get("low_point_world"),
                source="section_extreme_points",
                annotation_text=format_section_z_annotation_text(
                    group.get("high_point_world"),
                    group.get("low_point_world"),
                ),
            )
            if section_item is None:
                continue
            section_item["wheelhouse"] = group.get("wheelhouse")
            runtime_capture_wheelhouse_annotation_item(
                product_document.Application,
                product_document,
                root_product,
                annotation_runtime_context,
                section_item,
                category="section",
                regulation_axis_part=annotation_runtime_context["regulation_axis_part_info"],
            )
    results.append(
        {
            "status": "success",
            "message": f"已完成截面拓扑分组: {len(topology_group_results)} 组。",
            "section_topology_groups": topology_group_results,
            "axis_clearance_results": axis_clearance_results,
            "regulation_distance_measurements": regulation_distance_measurements,
        }
    )
    return results


def create_bounding_box_wireframe_in_part(
    min_point_world: Vector,
    max_point_world: Vector,
    target_part: Any,
    geometry_set_name: str,
    process_transform: Transform | None = None,
    target_document: Any | None = None,
    save_path: str | Path | None = None,
) -> dict[str, Any]:
    """
    功能: 在目标 Part 的几何集中创建包围盒八个点和十二条边线。
    输入: 包围盒世界坐标最小点/最大点、目标 Part、几何集名称、目标 Part 装配变换和可选保存路径。
    输出: 创建结果，包含点名、线名、尺寸和最大尺寸。
    """
    min_point = as_vector(min_point_world, "包围盒最小点")
    max_point = as_vector(max_point_world, "包围盒最大点")
    min_x, min_y, min_z = (min(a, b) for a, b in zip(min_point, max_point))
    max_x, max_y, max_z = (max(a, b) for a, b in zip(min_point, max_point))
    corners_world: list[Vector] = [
        (min_x, min_y, min_z),
        (max_x, min_y, min_z),
        (max_x, max_y, min_z),
        (min_x, max_y, min_z),
        (min_x, min_y, max_z),
        (max_x, min_y, max_z),
        (max_x, max_y, max_z),
        (min_x, max_y, max_z),
    ]
    corners_local = [
        inverse_transform_point(process_transform, point)
        if process_transform is not None
        else point
        for point in corners_world
    ]
    size = (max_x - min_x, max_y - min_y, max_z - min_z)
    max_size = max(size)
    hybrid_body = get_or_create_hybrid_body(target_part, geometry_set_name)
    factory = target_part.HybridShapeFactory
    point_features: list[Any] = []
    point_names: list[str] = []
    for index, point in enumerate(corners_local, start=1):
        feature = factory.AddNewPointCoord(*point)
        feature_name = f"{geometry_set_name}_P{index:02d}"
        set_if_possible(feature, "Name", feature_name)
        hybrid_body.AppendHybridShape(feature)
        point_features.append(feature)
        point_names.append(feature_name)
    edge_indexes = (
        (0, 1), (1, 2), (2, 3), (3, 0),
        (4, 5), (5, 6), (6, 7), (7, 4),
        (0, 4), (1, 5), (2, 6), (3, 7),
    )
    longest_edge_candidates: list[dict[str, Any]] = []
    y_edge_candidates: list[dict[str, Any]] = []
    selected_longest_edge: dict[str, Any] | None = None
    selected_y_edge: dict[str, Any] | None = None
    line_names: list[str] = []
    for index, (start_index, end_index) in enumerate(edge_indexes, start=1):
        line = factory.AddNewLinePtPt(
            create_reference(target_part, point_features[start_index]),
            create_reference(target_part, point_features[end_index]),
        )
        line_name = f"{geometry_set_name}_E{index:02d}"
        set_if_possible(line, "Name", line_name)
        hybrid_body.AppendHybridShape(line)
        line_names.append(line_name)
        start_world = corners_world[start_index]
        end_world = corners_world[end_index]
        start_local = corners_local[start_index]
        end_local = corners_local[end_index]
        edge_vector = subtract_vectors(end_world, start_world)
        edge_length = vector_length(edge_vector)
        axis_components = tuple(abs(value) for value in edge_vector)
        edge_axis = max(
            range(3),
            key=lambda axis_index: axis_components[axis_index],
        )
        edge_axis_name = ("X", "Y", "Z")[edge_axis]
        midpoint_world = average_vectors([start_world, end_world])
        candidate = {
            "edge_index": index,
            "line_name": line_name,
            "start_corner_index": start_index + 1,
            "end_corner_index": end_index + 1,
            "point1_world": round_vector(start_world),
            "point2_world": round_vector(end_world),
            "point1_process": round_vector(start_local),
            "point2_process": round_vector(end_local),
            "midpoint_world": round_vector(midpoint_world),
            "length": round(float(edge_length), 6),
            "axis": edge_axis_name,
            "selection_score": tuple(round(value, 6) for value in midpoint_world),
        }
        if abs(edge_length - max_size) <= 1.0e-6:
            longest_edge_candidates.append(candidate)
        if edge_axis_name == "Y" and abs(edge_length - size[1]) <= 1.0e-6:
            y_edge_candidates.append(candidate)
    if longest_edge_candidates:
        selected_longest_edge = max(
            longest_edge_candidates,
            key=lambda row: row["selection_score"],
        )
    if y_edge_candidates:
        selected_y_edge = max(
            y_edge_candidates,
            key=lambda row: row["selection_score"],
        )
    try:
        target_part.Update()
    except Exception as exc:
        print(f"[警告] 包围盒线框创建后 Part 更新失败: {geometry_set_name}: {exc}")
    save_result: dict[str, Any] | None = None
    if target_document is not None:
        save_result = save_document_if_modified(
            target_document,
            Path(save_path) if save_path is not None else None,
        )
    return {
        "status": "success",
        "geometry_set_name": geometry_set_name,
        "point_names": point_names,
        "line_names": line_names,
        "min_point_world": round_vector((min_x, min_y, min_z)),
        "max_point_world": round_vector((max_x, max_y, max_z)),
        "size": round_vector(size),
        "max_size": round(float(max_size), 6),
        "y_size": round(float(size[1]), 6),
        "longest_edge_candidates": longest_edge_candidates,
        "selected_longest_edge": selected_longest_edge,
        "y_edge_candidates": y_edge_candidates,
        "selected_y_edge": selected_y_edge,
        "annotation_points": (
            {
                "point1": selected_y_edge.get("point1_world"),
                "point2": selected_y_edge.get("point2_world"),
                "label_source": "bbox_y_edge_max_xyz_midpoint",
            }
            if selected_y_edge is not None
            else None
        ),
        "save_result": save_result,
    }


def create_wheelhouse_bounding_box_wireframes(
    part: Any,
    part_document: Any,
    process_transform: Transform | None,
    wheelhouse_bounding_boxes: dict[str, Any] | None,
    save_path: str | Path | None = None,
) -> tuple[list[dict[str, Any]], dict[str, float]]:
    """
    功能: 在过程 Part 中创建前/后轮罩包围盒线框，并计算 Front-q、Rear-q。
    输入: 过程 Part、过程 PartDocument、过程 Part 装配变换、前后轮罩包围盒结果和保存路径。
    输出: 线框创建结果列表和 q 值字典。
    """
    results: list[dict[str, Any]] = []
    measurements: dict[str, float] = {}
    configs = [
        (
            str(wheelhouse_label),
            f"{wheelhouse_definition_from_label(wheelhouse_label)['geometry_set_name']}包围盒",
            wheelhouse_measurement_key(wheelhouse_label, "q"),
        )
        for wheelhouse_label in (wheelhouse_bounding_boxes or {}).keys()
    ]
    for wheelhouse_label, geometry_set_name, measurement_key in configs:
        bbox_row = (wheelhouse_bounding_boxes or {}).get(wheelhouse_label) or {}
        bbox = bbox_row.get("bbox_world") or {}
        min_point = bbox.get("min_point")
        max_point = bbox.get("max_point")
        if not min_point or not max_point:
            result = {
                "status": "failed",
                "wheelhouse": wheelhouse_label,
                "measurement_key": measurement_key,
                "message": "缺少包围盒最大点或最小点，无法创建线框。",
            }
            print(f"[警告] {wheelhouse_label} 包围盒线框创建失败: {result['message']}")
            results.append(result)
            continue
        try:
            result = create_bounding_box_wireframe_in_part(
                as_vector(min_point, f"{wheelhouse_label} bbox min"),
                as_vector(max_point, f"{wheelhouse_label} bbox max"),
                part,
                geometry_set_name,
                process_transform=process_transform,
                target_document=part_document,
                save_path=save_path,
            )
            result["wheelhouse"] = wheelhouse_label
            result["measurement_key"] = measurement_key
            measurements[measurement_key] = float(result.get("y_size") or result["size"][1])
            print(f"{measurement_key}: {measurements[measurement_key]} mm")
            results.append(result)
        except Exception as exc:
            result = {
                "status": "failed",
                "wheelhouse": wheelhouse_label,
                "measurement_key": measurement_key,
                "message": str(exc),
            }
            print(f"[警告] {wheelhouse_label} 包围盒线框创建失败: {exc}")
            results.append(result)
    return results, measurements


def create_regulation_axis_part(
    product_document: Any,
    root_product: Any,
    front_component: Any,
    rear_component: Any | None,
    segments: list[RegulationAxisSegment],
    wheelhouse_bounding_boxes: dict[str, Any] | None = None,
    output_dir: str | Path | None = None,
    course_dir: str | Path | None = None,
    run_timestamp: str | None = None,
    annotation_runtime_context: dict[str, Any] | None = None,
) -> dict[str, Any]:
    """
    功能: 新建法规校核 CATPart，先写入轴线和截面切割面，再测包围盒，最后创建截面结果。
    输入: ProductDocument、根 Product、前后轮罩组件、轴线段列表、可选前后轮罩包围盒结果和输出目录。
    输出: 新 CATPart 的结果信息字典。
    """
    if isinstance(front_component, dict):
        wheelhouse_components = front_component
    else:
        wheelhouse_components = {
            FRONT_WHEELHOUSE_LABEL: front_component,
            REAR_WHEELHOUSE_LABEL: rear_component,
        }
    axis_part_number = f"{REGULATION_AXIS_PART_NUMBER}_{run_timestamp or build_timestamp()}"
    axis_part_save_path = build_process_part_save_path(axis_part_number, output_dir)
    axis_product, initial_axis_part_save_result = add_saved_part_file_to_product(
        product_document,
        root_product,
        axis_part_number,
        axis_part_number,
        axis_part_save_path,
    )
    part, _part_document = get_part_and_document_from_product(axis_product)
    try:
        set_if_possible(part, "PartNumber", axis_part_number)
    except Exception:
        pass

    hybrid_bodies = part.HybridBodies
    body_by_name: dict[str, Any] = {}
    created_features: list[dict[str, Any]] = []
    for segment in segments:
        hybrid_body = body_by_name.get(segment.geometry_set_name)
        if hybrid_body is None:
            hybrid_body = hybrid_bodies.Add()
            set_if_possible(hybrid_body, "Name", segment.geometry_set_name)
            body_by_name[segment.geometry_set_name] = hybrid_body
        created_feature = append_axis_line_to_hybrid_body(_part_document, part, hybrid_body, segment)
        created_feature["wheelhouse"] = segment.wheelhouse_label
        created_feature["geometry_set_name"] = segment.geometry_set_name
        created_features.append(created_feature)

    try:
        part.Update()
    except Exception as exc:
        raise RuntimeError("法规校核轴线段创建后 Part 更新失败。") from exc
    axis_part_save_result = save_document_if_modified(_part_document, axis_part_save_path)
    if axis_part_save_result.get("status") == "failed":
        print(f"[警告] 过程 CATPart 保存失败: {axis_part_save_result.get('message')}")
    process_transform = product_position_transform(product_document, axis_product)

    if wheelhouse_bounding_boxes is None:
        print("\n-- 轴线和切割面创建完成，开始测量前后轮罩包围盒 --")
        wheelhouse_bounding_boxes = evaluate_wheelhouse_extreme_bounding_boxes(
            root_product,
            wheelhouse_components,
        )

    bbox_wireframe_results, bbox_measurements = create_wheelhouse_bounding_box_wireframes(
        part,
        _part_document,
        process_transform,
        wheelhouse_bounding_boxes,
        save_path=axis_part_save_path,
    )
    if annotation_runtime_context is not None:
        annotation_runtime_context["regulation_axis_part_info"] = {
            "component_name": product_display_name(axis_product),
            "component_part_number": product_part_number(axis_product),
            "section_curve_results": [],
        }
        for bbox_row in bbox_wireframe_results:
            if bbox_row.get("status") != "success":
                continue
            annotation_points = bbox_row.get("annotation_points") or {}
            bbox_item = build_wheelhouse_annotation_item(
                bbox_row.get("measurement_key"),
                bbox_row.get("y_size"),
                annotation_points.get("point1"),
                annotation_points.get("point2"),
                source=str(annotation_points.get("label_source") or "bbox_y_edge"),
            )
            if bbox_item is None:
                continue
            bbox_item["wheelhouse"] = bbox_row.get("wheelhouse")
            bbox_item["text_offset_direction"] = WHEELHOUSE_ANNOTATION_BBOX_OFFSET_DIRECTION
            runtime_capture_wheelhouse_annotation_item(
                product_document.Application,
                product_document,
                root_product,
                annotation_runtime_context,
                bbox_item,
                category="bbox",
                regulation_axis_part=annotation_runtime_context["regulation_axis_part_info"],
            )
    section_curve_results = create_section_curves_from_regulation_planes(
        product_document,
        root_product,
        wheelhouse_components,
        part,
        body_by_name,
        process_transform,
        created_features,
        output_dir,
        course_dir=course_dir,
        run_timestamp=run_timestamp,
        annotation_runtime_context=annotation_runtime_context,
    )
    try:
        part.Update()
    except Exception as exc:
        print(f"[警告] 过程 CATPart 截面结果创建后更新失败: {exc}")
    final_axis_part_save_result = save_document_if_modified(_part_document, axis_part_save_path)
    if final_axis_part_save_result.get("status") == "failed":
        print(f"[警告] 过程 CATPart 最终保存失败: {final_axis_part_save_result.get('message')}")

    return {
        "component_name": product_display_name(axis_product),
        "component_part_number": product_part_number(axis_product),
        "part_save_path": str(axis_part_save_path),
        "initial_part_save_result": initial_axis_part_save_result,
        "part_save_result": axis_part_save_result,
        "final_part_save_result": final_axis_part_save_result,
        "geometry_set_names": list(body_by_name.keys()),
        "segments": [regulation_axis_segment_to_dict(segment) for segment in segments],
        "created_features": created_features,
        "wheelhouse_extreme_bounding_boxes": wheelhouse_bounding_boxes,
        "bbox_wireframe_results": bbox_wireframe_results,
        "bbox_measurements": bbox_measurements,
        "section_curve_results": section_curve_results,
    }


def format_wheelhouse_annotation_text(value: Any) -> str:
    return f"{float(value):.3f} mm"


def coerce_annotation_point(value: Any, label: str) -> tuple[float, float, float] | None:
    try:
        return round_vector(as_vector(value, label))
    except Exception:
        return None


def build_wheelhouse_annotation_item(
    measurement_key: Any,
    measurement_value: Any,
    point1: Any,
    point2: Any,
    *,
    source: str,
    annotation_text: Any | None = None,
) -> dict[str, Any] | None:
    key = str(measurement_key or "").strip()
    if measurement_value is None:
        return None
    p1 = coerce_annotation_point(point1, f"{key} annotation point1")
    p2 = coerce_annotation_point(point2, f"{key} annotation point2")
    if p1 is None or p2 is None:
        return None
    try:
        text_value = (
            str(annotation_text)
            if annotation_text is not None
            else format_wheelhouse_annotation_text(measurement_value)
        )
    except Exception:
        return None
    return {
        "measurement_key": key,
        "measurement_value": round(float(measurement_value), 6),
        "annotation_name": key.replace("-", "_"),
        "annotation_text": text_value,
        "point1": p1,
        "point2": p2,
        "source": source,
    }


def format_section_z_annotation_text(point1: Any, point2: Any) -> str:
    p1 = as_vector(point1, "section annotation point1")
    p2 = as_vector(point2, "section annotation point2")
    z_distance = abs(float(p2[2]) - float(p1[2]))
    if abs(z_distance - round(z_distance)) <= 1.0e-6:
        value_text = str(int(round(z_distance)))
    else:
        value_text = f"{z_distance:.3f}".rstrip("0").rstrip(".")
    return f"Z={value_text}mm"


def collect_wheelhouse_regulation_annotation_items(
    regulation_axis_part: dict[str, Any] | None,
) -> list[dict[str, Any]]:
    rows_by_key: dict[str, dict[str, Any]] = {}
    regulation_axis_part = regulation_axis_part or {}

    for row in regulation_axis_part.get("bbox_wireframe_results") or []:
        annotation_points = row.get("annotation_points") or {}
        item = build_wheelhouse_annotation_item(
            row.get("measurement_key"),
            row.get("y_size"),
            annotation_points.get("point1"),
            annotation_points.get("point2"),
            source=str(annotation_points.get("label_source") or "bbox_y_edge"),
        )
        if item is not None:
            item["wheelhouse"] = row.get("wheelhouse")
            item["text_offset_direction"] = WHEELHOUSE_ANNOTATION_BBOX_OFFSET_DIRECTION
            rows_by_key[item["measurement_key"]] = item

    for section_result in regulation_axis_part.get("section_curve_results") or []:
        for group in section_result.get("section_topology_groups") or []:
            item = build_wheelhouse_annotation_item(
                group.get("measurement_key"),
                group.get("measurement_value"),
                group.get("high_point_world"),
                group.get("low_point_world"),
                source="section_extreme_points",
                annotation_text=format_section_z_annotation_text(
                    group.get("high_point_world"),
                    group.get("low_point_world"),
                ),
            )
            if item is not None:
                item["wheelhouse"] = group.get("wheelhouse")
                rows_by_key[item["measurement_key"]] = item
        for clearance_result in section_result.get("axis_clearance_results") or []:
            measurement_geometry = clearance_result.get("measurement_geometry") or {}
            item = build_wheelhouse_annotation_item(
                clearance_result.get("measurement_key"),
                clearance_result.get("measurement_value"),
                measurement_geometry.get("body_point_world"),
                measurement_geometry.get("axis_point_world"),
                source="wheelhouse_to_axis_min_distance",
            )
            if item is not None:
                item["wheelhouse"] = clearance_result.get("wheelhouse")
                rows_by_key[item["measurement_key"]] = item

    order_index = {key: index for index, key in enumerate(WHEELHOUSE_ANNOTATION_KEYS)}
    return sorted(
        rows_by_key.values(),
        key=lambda row: order_index.get(str(row.get("measurement_key")), 999),
    )


def wheelhouse_annotation_item_to_tool_kwargs(item: dict[str, Any]) -> dict[str, Any]:
    kwargs = {
        "point1": item["point1"],
        "point2": item["point2"],
        "annotation_name": item["annotation_name"],
        "annotation_text": item["annotation_text"],
    }
    if item.get("text_offset_direction") is not None:
        kwargs["text_offset_direction"] = item["text_offset_direction"]
    return kwargs


def create_wheelhouse_annotation_runtime_context(result_dir: str | Path) -> dict[str, Any]:
    screenshot_dir = Path(result_dir) / "screenshots"
    screenshot_dir.mkdir(parents=True, exist_ok=True)
    process_part_save_path = Path(result_dir) / f"{WHEELHOUSE_ANNOTATION_PART_NUMBER}.CATPart"
    return {
        "process_part": None,
        "process_part_document": None,
        "process_component": None,
        "process_part_save_path": process_part_save_path,
        "annotation_results": [],
        "annotation_items": [],
        "screenshots": [],
        "scene_visibility_rows": [],
        "state_rows": {},
        "screenshot_dir": screenshot_dir,
        "regulation_axis_part_info": None,
    }


def wheelhouse_runtime_annotation_result(runtime_context: dict[str, Any]) -> dict[str, Any]:
    annotation_results = list(runtime_context.get("annotation_results") or [])
    annotation_items = list(runtime_context.get("annotation_items") or [])
    screenshots = list(runtime_context.get("screenshots") or [])
    success_count = sum(1 for row in annotation_results if row.get("status") == "success")
    return {
        "status": "success" if not annotation_results or success_count == len(annotation_results) else "partial_failed",
        "count": len(annotation_results),
        "success_count": success_count,
        "annotation_items": annotation_items,
        "results": annotation_results,
        "process_part_save_path": str(runtime_context.get("process_part_save_path") or ""),
        "screenshots": screenshots,
        "scene_visibility_rows": list(runtime_context.get("scene_visibility_rows") or []),
    }


def wheelhouse_runtime_screenshot_result(runtime_context: dict[str, Any]) -> dict[str, Any]:
    screenshots = list(runtime_context.get("screenshots") or [])
    success_count = sum(1 for row in screenshots if row.get("status") == "success")
    return {
        "status": "success" if not screenshots or success_count == len(screenshots) else "partial_failed",
        "count": len(screenshots),
        "success_count": success_count,
        "screenshot_dir": str(runtime_context.get("screenshot_dir") or ""),
        "screenshots": screenshots,
        "scene_visibility_rows": list(runtime_context.get("scene_visibility_rows") or []),
    }


def runtime_clear_annotation_visuals(
    product_document: Any,
    root_product: Any,
    annotation_part: Any | None,
    annotation_result: dict[str, Any] | None,
    state_rows: dict[int, dict[str, Any]],
) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    if annotation_part is None or not annotation_result:
        return rows
    geometry = annotation_result.get("_annotation_geometry") or annotation_result.get("geometry") or {}
    geometry_document = annotation_result.get("_process_part_document") or product_document
    for feature_key in ("point1_feature", "point2_feature", "line_feature"):
        feature = geometry.get(feature_key)
        if feature is None:
            continue
        label = f"annotation_geometry:{annotation_result.get('annotation_name')}:{feature_key}"
        try:
            rows.append(
                remember_and_set_visibility(
                    geometry_document,
                    feature,
                    False,
                    state_rows,
                    kind="object",
                    label=label,
                )
            )
        except Exception as exc:
            rows.append(
                {
                    "status": "failed",
                    "kind": "object",
                    "label": label,
                    "visible": False,
                    "message": str(exc),
                    "document": safe_attr_text(geometry_document, "Name", ""),
                }
            )
    marker = annotation_result.get("_text_marker")
    if marker is not None:
        rows.append(
            remember_and_set_visibility(
                product_document,
                marker,
                False,
                state_rows,
                kind="object",
                label=f"annotation_text:{annotation_result.get('annotation_name')}",
            )
        )
    try:
        annotation_part.Update()
    except Exception:
        pass
    try:
        product_document.Application.ActiveWindow.ActiveViewer.Update()
    except Exception:
        pass
    return rows


def runtime_capture_wheelhouse_annotation_item(
    catia: Any,
    product_document: Any,
    root_product: Any,
    runtime_context: dict[str, Any],
    annotation_item: dict[str, Any],
    *,
    category: str,
    regulation_axis_part: dict[str, Any] | None,
) -> dict[str, Any]:
    try:
        from catia_annotation_tools import create_distance_annotation
    except Exception:
        tool_path = Path(__file__).resolve().with_name("catia_annotation_tools.py")
        spec = importlib.util.spec_from_file_location("catia_annotation_tools_runtime", tool_path)
        if spec is None or spec.loader is None:
            raise RuntimeError(f"无法加载工具文件: {tool_path}")
        annotation_tool = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(annotation_tool)
        create_distance_annotation = annotation_tool.create_distance_annotation

    process_part = runtime_context.get("process_part")
    process_part_document = runtime_context.get("process_part_document")
    process_component = runtime_context.get("process_component")
    create_kwargs = wheelhouse_annotation_item_to_tool_kwargs(annotation_item)
    create_kwargs.update(
        {
            "hybrid_body_name": WHEELHOUSE_ANNOTATION_BODY_NAME,
            "text_offset_distance": WHEELHOUSE_ANNOTATION_OFFSET_DISTANCE,
            "feature_color": WHEELHOUSE_ANNOTATION_COLOR,
            "text_color": WHEELHOUSE_ANNOTATION_COLOR,
            "line_width": WHEELHOUSE_ANNOTATION_LINE_WIDTH,
            "text_size": WHEELHOUSE_ANNOTATION_TEXT_SIZE,
            "process_part": process_part,
            "process_part_document": process_part_document,
            "process_component": process_component,
            "process_part_number": WHEELHOUSE_ANNOTATION_PART_NUMBER,
            "process_part_name": WHEELHOUSE_ANNOTATION_PART_NAME,
            "process_part_save_path": runtime_context.get("process_part_save_path"),
            "create_process_part_if_missing": True,
            "reopen_product_after_create": False,
        }
    )
    result = create_distance_annotation(catia, **create_kwargs)
    runtime_context["process_part"] = result.get("_process_part") or runtime_context.get("process_part")
    runtime_context["process_part_document"] = result.get("_process_part_document") or runtime_context.get("process_part_document")
    runtime_context["process_component"] = result.get("_process_component") or runtime_context.get("process_component")
    runtime_context.setdefault("annotation_results", []).append(strip_private_result_fields(result))
    runtime_context.setdefault("annotation_items", []).append(annotation_item)

    components = collect_reloaded_wheelhouse_components(root_product, regulation_axis_part)
    annotation_component = runtime_context.get("process_component")
    if annotation_component is not None:
        components["annotation"] = annotation_component
    annotation_part = runtime_context.get("process_part")
    annotation_part_result = result
    state_rows = runtime_context.setdefault("state_rows", {})
    scene_rows = set_wheelhouse_stage_visibility(
        product_document,
        root_product,
        components,
        annotation_part,
        str(annotation_item.get("wheelhouse") or ""),
        str(annotation_item.get("measurement_key") or ""),
        category,
        state_rows,
    )
    runtime_context.setdefault("scene_visibility_rows", []).extend(scene_rows)

    capture_tool = load_capture_document_view()
    screenshot_dir = Path(runtime_context.get("screenshot_dir") or (Path.cwd() / "screenshots"))
    screenshot_dir.mkdir(parents=True, exist_ok=True)
    spec_list = build_wheelhouse_screenshot_specs([annotation_item], components, screenshot_dir)
    if spec_list:
        spec = spec_list[0]
        try:
            root_product.Update()
        except Exception:
            pass
        shot_result = capture_tool(
            catia,
            product_document,
            spec["output_path"],
            mode="fixed",
            view_point=spec["view_point"],
            view_distance=spec["view_distance"],
            sight_direction=spec["sight_direction"],
            up_direction=spec["up_direction"],
            image_format=WHEELHOUSE_SCREENSHOT_IMAGE_FORMAT,
            wait_seconds=0.5,
        )
        runtime_context.setdefault("screenshots", []).append(
            {
                "status": "success",
                "measurement_key": spec["measurement_key"],
                "category": spec["category"],
                "annotation_name": spec["annotation_name"],
                "path": shot_result.get("screenshot_path"),
                "view": shot_result.get("view"),
            }
        )
    else:
        runtime_context.setdefault("screenshots", []).append(
            {
                "status": "failed",
                "measurement_key": annotation_item.get("measurement_key"),
                "category": category,
                "annotation_name": annotation_item.get("annotation_name"),
                "path": None,
                "message": "未生成有效截图配置。",
            }
        )
    runtime_context.setdefault("scene_visibility_rows", []).extend(
        runtime_clear_annotation_visuals(
            product_document,
            root_product,
            annotation_part,
            annotation_part_result,
            state_rows,
        )
    )
    return result


def strip_private_result_fields(value: Any) -> Any:
    if value is None or isinstance(value, (str, int, float, bool)):
        return value
    if isinstance(value, Path):
        return str(value)
    if isinstance(value, tuple):
        return [strip_private_result_fields(item) for item in value]
    if isinstance(value, list):
        return [strip_private_result_fields(item) for item in value]
    if isinstance(value, dict):
        return {
            str(key): strip_private_result_fields(item)
            for key, item in value.items()
            if not str(key).startswith("_")
        }
    if hasattr(value, "Name"):
        return safe_attr_text(value, "Name", repr(value))
    return repr(value)


def create_wheelhouse_regulation_annotations(
    catia: Any,
    product_document: Any,
    regulation_axis_part: dict[str, Any] | None,
    result_dir: str | Path,
) -> dict[str, Any]:
    annotation_items = collect_wheelhouse_regulation_annotation_items(regulation_axis_part)
    process_part_save_path = Path(result_dir) / f"{WHEELHOUSE_ANNOTATION_PART_NUMBER}.CATPart"
    if not annotation_items:
        return {
            "status": "skipped",
            "message": "未找到可用于创建标注的法规点。",
            "annotation_items": [],
            "process_part_save_path": str(process_part_save_path),
        }
    try:
        from catia_annotation_tools import create_distance_annotations
    except Exception as exc:
        try:
            tool_path = Path(__file__).resolve().with_name("catia_annotation_tools.py")
            spec = importlib.util.spec_from_file_location("catia_annotation_tools_local", tool_path)
            if spec is None or spec.loader is None:
                raise RuntimeError(f"无法加载工具文件: {tool_path}")
            annotation_tool = importlib.util.module_from_spec(spec)
            spec.loader.exec_module(annotation_tool)
            create_distance_annotations = annotation_tool.create_distance_annotations
        except Exception as fallback_exc:
            return {
                "status": "failed",
                "message": f"无法加载通用标注工具: {exc}; fallback: {fallback_exc}",
                "annotation_items": annotation_items,
                "process_part_save_path": str(process_part_save_path),
            }
    try:
        product_document.Activate()
    except Exception:
        pass
    tool_annotations = [
        wheelhouse_annotation_item_to_tool_kwargs(item)
        for item in annotation_items
    ]
    result = create_distance_annotations(
        catia,
        annotations=tool_annotations,
        hybrid_body_name=WHEELHOUSE_ANNOTATION_BODY_NAME,
        text_offset_direction=WHEELHOUSE_ANNOTATION_OFFSET_DIRECTION,
        text_offset_distance=WHEELHOUSE_ANNOTATION_OFFSET_DISTANCE,
        feature_color=WHEELHOUSE_ANNOTATION_COLOR,
        text_color=WHEELHOUSE_ANNOTATION_COLOR,
        line_width=WHEELHOUSE_ANNOTATION_LINE_WIDTH,
        text_size=WHEELHOUSE_ANNOTATION_TEXT_SIZE,
        process_part_number=WHEELHOUSE_ANNOTATION_PART_NUMBER,
        process_part_name=WHEELHOUSE_ANNOTATION_PART_NAME,
        process_part_save_path=process_part_save_path,
    )
    public_result = strip_private_result_fields(result)
    public_result["annotation_items"] = annotation_items
    public_result["process_part_save_path"] = str(process_part_save_path)
    if public_result.get("success_count") != len(annotation_items):
        public_result["status"] = "partial_failed"
    return public_result


def reopen_wheelhouse_product_for_annotation_refresh(
    catia: Any,
    product_document: Any,
    product_save_path: str | Path,
) -> dict[str, Any]:
    try:
        from catia_annotation_tools import reopen_product_document_for_marker_refresh
    except Exception as exc:
        try:
            tool_path = Path(__file__).resolve().with_name("catia_annotation_tools.py")
            spec = importlib.util.spec_from_file_location("catia_annotation_tools_local_reopen", tool_path)
            if spec is None or spec.loader is None:
                raise RuntimeError(f"无法加载工具文件: {tool_path}")
            annotation_tool = importlib.util.module_from_spec(spec)
            spec.loader.exec_module(annotation_tool)
            reopen_product_document_for_marker_refresh = annotation_tool.reopen_product_document_for_marker_refresh
        except Exception as fallback_exc:
            return {
                "status": "failed",
                "message": f"无法加载标注Product重开刷新工具: {exc}; fallback: {fallback_exc}",
            }
    result = reopen_product_document_for_marker_refresh(
        catia,
        product_document,
        product_save_path=product_save_path,
    )
    return result


def iter_product_tree(product: Any) -> Iterable[Any]:
    yield product
    try:
        products = product.Products
    except Exception:
        return
    for child in iter_collection(products):
        yield from iter_product_tree(child)


def find_product_component_by_tokens(root_product: Any, tokens: Iterable[str]) -> Any | None:
    normalized_tokens = {str(token).strip() for token in tokens if str(token).strip()}
    if not normalized_tokens:
        return None
    for product in iter_product_tree(root_product):
        if product is root_product:
            continue
        values = {
            product_display_name(product),
            product_part_number(product),
            safe_attr_text(product, "Name"),
            safe_attr_text(product, "PartNumber"),
        }
        if any(value in normalized_tokens for value in values if value):
            return product
    return None


def direct_child_products_by_visibility(root_product: Any) -> list[Any]:
    try:
        return list(iter_collection(root_product.Products))
    except Exception:
        return []


def remember_and_set_visibility(
    product_document: Any,
    target: Any,
    visible: bool,
    state_rows: dict[int, dict[str, Any]],
    *,
    kind: str,
    label: str,
) -> dict[str, Any]:
    key = id(target)
    if key not in state_rows:
        state_rows[key] = {
            "target": target,
            "kind": kind,
            "label": label,
            "visibility": get_object_visibility(product_document, target),
        }
    try:
        if kind == "product":
            set_product_visibility(product_document, target, visible)
        else:
            set_object_visibility(product_document, target, visible)
        return {
            "status": "success",
            "kind": kind,
            "label": label,
            "visible": bool(visible),
        }
    except Exception as exc:
        return {
            "status": "failed",
            "kind": kind,
            "label": label,
            "visible": bool(visible),
            "message": str(exc),
        }


def restore_recorded_visibility(product_document: Any, state_rows: dict[int, dict[str, Any]]) -> list[dict[str, Any]]:
    restore_rows: list[dict[str, Any]] = []
    for row in reversed(list(state_rows.values())):
        original_visibility = row.get("visibility") or {}
        visible = bool(original_visibility.get("visible", True))
        target = row.get("target")
        kind = str(row.get("kind") or "object")
        label = str(row.get("label") or "")
        try:
            if kind == "product":
                set_product_visibility(product_document, target, visible)
            else:
                set_object_visibility(product_document, target, visible)
            restore_rows.append(
                {
                    "status": "success",
                    "kind": kind,
                    "label": label,
                    "restored_visible": visible,
                }
            )
        except Exception as exc:
            restore_rows.append(
                {
                    "status": "failed",
                    "kind": kind,
                    "label": label,
                    "restored_visible": visible,
                    "message": str(exc),
                }
            )
    return restore_rows


def record_visibility_state(
    product_document: Any,
    target: Any,
    state_rows: dict[int, dict[str, Any]],
    *,
    kind: str,
    label: str,
) -> None:
    key = id(target)
    if key in state_rows:
        return
    state_rows[key] = {
        "target": target,
        "kind": kind,
        "label": label,
        "visibility": get_object_visibility(product_document, target),
    }


def get_marker3ds_from_product(root_product: Any) -> Any | None:
    try:
        return root_product.GetTechnologicalObject("Marker3Ds")
    except Exception:
        return None


def collect_marker3d_by_name(root_product: Any) -> dict[str, Any]:
    marker3ds = get_marker3ds_from_product(root_product)
    if marker3ds is None:
        return {}
    markers: dict[str, Any] = {}
    for marker in iter_collection(marker3ds):
        name = safe_attr_text(marker, "Name")
        if name:
            markers[name] = marker
    return markers


def collect_annotation_features_by_name(annotation_part: Any) -> dict[str, Any]:
    body = find_hybrid_body_by_name(annotation_part, WHEELHOUSE_ANNOTATION_BODY_NAME)
    if body is None:
        body = find_hybrid_body_by_name(annotation_part, "距离标注")
    if body is None:
        return {}
    features: dict[str, Any] = {}
    for feature in iter_hybrid_shapes(body):
        name = safe_attr_text(feature, "Name")
        if name:
            features[name] = feature
    return features


def show_only_annotation_elements(
    product_document: Any,
    root_product: Any,
    annotation_part: Any,
    annotation_names: Iterable[str],
    state_rows: dict[int, dict[str, Any]],
) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    selected_names = {str(name) for name in annotation_names if str(name)}
    selected_feature_prefixes = tuple(f"{name}_" for name in selected_names)
    selected_text_names = {f"{name}_Text" for name in selected_names}

    for feature_name, feature in collect_annotation_features_by_name(annotation_part).items():
        visible = feature_name.startswith(selected_feature_prefixes) if selected_feature_prefixes else False
        rows.append(
            remember_and_set_visibility(
                product_document,
                feature,
                visible,
                state_rows,
                kind="object",
                label=f"annotation_feature:{feature_name}",
            )
        )

    for marker_name, marker in collect_marker3d_by_name(root_product).items():
        if not marker_name.endswith("_Text"):
            continue
        visible = marker_name in selected_text_names
        rows.append(
            remember_and_set_visibility(
                product_document,
                marker,
                visible,
                state_rows,
                kind="object",
                label=f"annotation_text:{marker_name}",
            )
        )
    return rows


def average_annotation_points(item: dict[str, Any]) -> Vector:
    return average_vectors([
        as_vector(item["point1"], "annotation point1"),
        as_vector(item["point2"], "annotation point2"),
    ])


def measurement_key_to_annotation_name(measurement_key: str) -> str:
    return str(measurement_key).replace("-", "_")


def collect_reloaded_wheelhouse_components(
    root_product: Any,
    regulation_axis_part: dict[str, Any] | None,
) -> dict[str, Any]:
    components: dict[str, Any] = {}
    wheelhouse_components: dict[str, Any] = {}
    for definition in WHEELHOUSE_SLOT_DEFINITIONS.values():
        component = find_product_component_by_tokens(
            root_product,
            [
                definition["label"],
                definition["component_part_number"],
                definition["measurement_prefix"],
            ],
        )
        if component is not None:
            wheelhouse_components[definition["label"]] = component
    legacy_front = find_product_component_by_tokens(root_product, [FRONT_WHEELHOUSE_LABEL, "Front_Wheelhouse"])
    legacy_rear = find_product_component_by_tokens(root_product, [REAR_WHEELHOUSE_LABEL, "Rear_Wheelhouse"])
    if legacy_front is not None:
        wheelhouse_components[FRONT_WHEELHOUSE_LABEL] = legacy_front
    if legacy_rear is not None:
        wheelhouse_components[REAR_WHEELHOUSE_LABEL] = legacy_rear
    components["wheelhouses"] = wheelhouse_components
    components["front"] = legacy_front or wheelhouse_components.get(LEFT_FRONT_WHEELHOUSE_LABEL) or wheelhouse_components.get(RIGHT_FRONT_WHEELHOUSE_LABEL)
    components["rear"] = legacy_rear or wheelhouse_components.get(LEFT_REAR_WHEELHOUSE_LABEL) or wheelhouse_components.get(RIGHT_REAR_WHEELHOUSE_LABEL)
    components["wheel"] = find_product_component_by_tokens(root_product, [WHEEL_ASSEMBLY_LABEL, "Wheel_Assembly"])
    axis_tokens = [REGULATION_AXIS_PART_NUMBER]
    if regulation_axis_part:
        axis_tokens.extend(
            [
                str(regulation_axis_part.get("component_name") or ""),
                str(regulation_axis_part.get("component_part_number") or ""),
            ]
        )
    components["axis"] = find_product_component_by_tokens(root_product, axis_tokens)
    components["annotation"] = find_product_component_by_tokens(
        root_product,
        [WHEELHOUSE_ANNOTATION_PART_NUMBER, WHEELHOUSE_ANNOTATION_PART_NAME],
    )

    section_components_by_key: dict[str, list[Any]] = {}
    for item in (regulation_axis_part or {}).get("section_curve_results") or []:
        if item.get("status") != "success" or not item.get("exported_component"):
            continue
        measurement_key = section_distance_measurement_key(item.get("wheelhouse"), item.get("section_plane_name"))
        if not measurement_key:
            continue
        exported_component = item.get("exported_component") or {}
        component = find_product_component_by_tokens(
            root_product,
            [
                str(exported_component.get("component_name") or ""),
                str(exported_component.get("component_part_number") or ""),
            ],
        )
        if component is not None:
            section_components_by_key.setdefault(measurement_key, []).append(component)
    components["sections_by_key"] = section_components_by_key
    components["all_sections"] = [
        component
        for section_components in section_components_by_key.values()
        for component in section_components
    ]
    return components


def build_wheelhouse_screenshot_specs(
    annotation_items: list[dict[str, Any]],
    components: dict[str, Any],
    screenshot_dir: Path,
) -> list[dict[str, Any]]:
    specs: list[dict[str, Any]] = []
    by_key = {str(item.get("measurement_key")): item for item in annotation_items}
    for key in WHEELHOUSE_ANNOTATION_KEYS:
        item = by_key.get(key)
        if not item:
            continue
        annotation_name = measurement_key_to_annotation_name(key)
        wheelhouse_label = str(item.get("wheelhouse") or "")
        wheelhouse_component = (components.get("wheelhouses") or {}).get(wheelhouse_label)
        if key.endswith("-q"):
            direction = WHEELHOUSE_SCREENSHOT_DIRECTIONS["bbox"]
            category = "bbox"
            view_distance = WHEELHOUSE_SCREENSHOT_BBOX_VIEW_DISTANCE
        elif key.endswith("-c"):
            direction = WHEELHOUSE_SCREENSHOT_DIRECTIONS["axis_clearance"]
            category = "axis_clearance"
            view_distance = WHEELHOUSE_SCREENSHOT_AXIS_CLEARANCE_VIEW_DISTANCE
        else:
            direction = WHEELHOUSE_SCREENSHOT_DIRECTIONS["section"]
            category = "section"
            view_distance = WHEELHOUSE_SCREENSHOT_SECTION_VIEW_DISTANCE
        specs.append(
            {
                "measurement_key": key,
                "annotation_name": annotation_name,
                "category": category,
                "wheelhouse": wheelhouse_label,
                "wheelhouse_component": wheelhouse_component,
                "view_point": average_annotation_points(item),
                "sight_direction": direction["sight_direction"],
                "up_direction": direction["up_direction"],
                "view_distance": view_distance,
                "output_path": str(screenshot_dir / f"{annotation_name}.png"),
            }
        )
    order_index = {key: index for index, key in enumerate(WHEELHOUSE_SCREENSHOT_SEQUENCE)}
    return sorted(specs, key=lambda row: order_index.get(str(row.get("measurement_key")), 999))


def get_annotation_target_names(measurement_key: str) -> list[str]:
    key = str(measurement_key)
    if key.endswith("-q") or key.endswith("-c"):
        return [measurement_key_to_annotation_name(key)]
    if "-p" in key:
        return [measurement_key_to_annotation_name(key)]
    return []


def set_common_wheelhouse_visibility(
    product_document: Any,
    root_product: Any,
    components: dict[str, Any],
    state_rows: dict[int, dict[str, Any]],
) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    for label, component in (components.get("wheelhouses") or {}).items():
        if component is None:
            continue
        rows.append(
            remember_and_set_visibility(
                product_document,
                component,
                True,
                state_rows,
                kind="product",
                label=f"component:wheelhouse:{label}",
            )
        )
    common_visible_keys = ("axis", "annotation")
    for key in common_visible_keys:
        component = components.get(key)
        if component is None:
            continue
        rows.append(
            remember_and_set_visibility(
                product_document,
                component,
                True,
                state_rows,
                kind="product",
                label=f"component:{key}",
            )
        )
    for key in ("wheel",):
        component = components.get(key)
        if component is None:
            continue
        rows.append(
            remember_and_set_visibility(
                product_document,
                component,
                False,
                state_rows,
                kind="product",
                label=f"component:{key}",
            )
        )
    for section_component in components.get("all_sections") or []:
        rows.append(
            remember_and_set_visibility(
                product_document,
                section_component,
                False,
                state_rows,
                kind="product",
                label=f"section:{product_part_number(section_component)}",
            )
        )
    return rows


def set_annotation_group_visibility(
    product_document: Any,
    root_product: Any,
    annotation_part: Any | None,
    annotation_names: Iterable[str],
    state_rows: dict[int, dict[str, Any]],
) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    selected_names = {str(name) for name in annotation_names if str(name)}
    selected_feature_prefixes = tuple(f"{name}_" for name in selected_names)
    selected_text_names = {f"{name}_Text" for name in selected_names}
    if annotation_part is None:
        return rows
    for feature_name, feature in collect_annotation_features_by_name(annotation_part).items():
        visible = feature_name.startswith(selected_feature_prefixes) if selected_feature_prefixes else False
        rows.append(
            remember_and_set_visibility(
                product_document,
                feature,
                visible,
                state_rows,
                kind="object",
                label=f"annotation_feature:{feature_name}",
            )
        )
    for marker_name, marker in collect_marker3d_by_name(root_product).items():
        if not marker_name.endswith("_Text"):
            continue
        visible = marker_name in selected_text_names
        rows.append(
            remember_and_set_visibility(
                product_document,
                marker,
                visible,
                state_rows,
                kind="object",
                label=f"annotation_text:{marker_name}",
            )
        )
    return rows


def set_section_screenshot_visibility(
    product_document: Any,
    components: dict[str, Any],
    measurement_key: str,
    state_rows: dict[int, dict[str, Any]],
) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    selected_sections = components.get("sections_by_key") or {}
    for section_component in components.get("all_sections") or []:
        visible = section_component in selected_sections.get(measurement_key, [])
        rows.append(
            remember_and_set_visibility(
                product_document,
                section_component,
                visible,
                state_rows,
                kind="product",
                label=f"section:{product_part_number(section_component)}",
            )
        )
    return rows


def set_wheelhouse_stage_visibility(
    product_document: Any,
    root_product: Any,
    components: dict[str, Any],
    annotation_part: Any | None,
    wheelhouse_label: str,
    measurement_key: str,
    category: str,
    state_rows: dict[int, dict[str, Any]],
) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    for label, component in (components.get("wheelhouses") or {}).items():
        visible = str(label) == str(wheelhouse_label) and str(category) != "section"
        rows.append(
            remember_and_set_visibility(
                product_document,
                component,
                visible,
                state_rows,
                kind="product",
                label=f"component:wheelhouse:{label}",
            )
        )
    annotation_component = components.get("annotation")
    if annotation_component is not None:
        rows.append(
            remember_and_set_visibility(
                product_document,
                annotation_component,
                True,
                state_rows,
                kind="product",
                label="component:annotation",
            )
        )
    axis_visible = str(category) == "bbox"
    for key in ("axis", "wheel"):
        component = components.get(key)
        if component is None:
            continue
        rows.append(
            remember_and_set_visibility(
                product_document,
                component,
                axis_visible if key == "axis" else False,
                state_rows,
                kind="product",
                label=f"component:{key}:stage:{measurement_key}",
            )
        )
    annotation_name = measurement_key_to_annotation_name(measurement_key)
    rows.extend(
        set_annotation_group_visibility(
            product_document,
            root_product,
            annotation_part,
            [annotation_name],
            state_rows,
        )
    )
    rows.extend(
        set_section_screenshot_visibility(
            product_document,
            components,
            measurement_key,
            state_rows,
        )
    )
    return rows


def set_wheelhouse_final_visibility(
    product_document: Any,
    root_product: Any,
    components: dict[str, Any],
    annotation_part: Any | None,
    annotation_part_document: Any | None,
    state_rows: dict[int, dict[str, Any]],
) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    for label, component in (components.get("wheelhouses") or {}).items():
        rows.append(
            remember_and_set_visibility(
                product_document,
                component,
                True,
                state_rows,
                kind="product",
                label=f"component:wheelhouse:final:{label}",
            )
        )
    for key in ("axis", "annotation"):
        component = components.get(key)
        if component is None:
            continue
        rows.append(
            remember_and_set_visibility(
                product_document,
                component,
                True,
                state_rows,
                kind="product",
                label=f"component:{key}:final",
            )
        )
    for section_component in components.get("all_sections") or []:
        rows.append(
            remember_and_set_visibility(
                product_document,
                section_component,
                True,
                state_rows,
                kind="product",
                label=f"section:final:{product_part_number(section_component)}",
            )
        )
    if annotation_part is not None:
        annotation_document = annotation_part_document or product_document
        for feature_name, feature in collect_annotation_features_by_name(annotation_part).items():
            try:
                rows.append(
                    remember_and_set_visibility(
                        annotation_document,
                        feature,
                        True,
                        state_rows,
                        kind="object",
                        label=f"annotation_feature:final:{feature_name}",
                    )
                )
            except Exception as exc:
                rows.append(
                    {
                        "status": "failed",
                        "kind": "object",
                        "label": f"annotation_feature:final:{feature_name}",
                        "visible": True,
                        "message": str(exc),
                    }
                )
    for marker_name, marker in collect_marker3d_by_name(root_product).items():
        if marker_name.endswith("_Text"):
            rows.append(
                remember_and_set_visibility(
                    product_document,
                    marker,
                    True,
                    state_rows,
                    kind="object",
                    label=f"annotation_text:final:{marker_name}",
                )
            )
    return rows


def reset_annotation_text_visibility(
    product_document: Any,
    root_product: Any,
    annotation_part: Any | None,
    state_rows: dict[int, dict[str, Any]],
) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    if annotation_part is None:
        return rows
    for feature_name, feature in collect_annotation_features_by_name(annotation_part).items():
        rows.append(
            remember_and_set_visibility(
                product_document,
                feature,
                False,
                state_rows,
                kind="object",
                label=f"annotation_feature:hidden:{feature_name}",
            )
        )
    for marker_name, marker in collect_marker3d_by_name(root_product).items():
        if marker_name.endswith("_Text"):
            rows.append(
                remember_and_set_visibility(
                    product_document,
                    marker,
                    False,
                    state_rows,
                    kind="object",
                    label=f"annotation_text:hidden:{marker_name}",
                )
            )
    return rows


def refresh_and_capture_document_view(
    catia: Any,
    product_document: Any,
    output_path: str | Path,
    *,
    view_point: Vector,
    sight_direction: Vector,
    up_direction: Vector,
    view_distance: float,
) -> dict[str, Any]:
    capture_document_view = load_capture_document_view()
    shot_result = capture_document_view(
        catia,
        product_document,
        output_path,
        mode="fixed",
        view_point=view_point,
        view_distance=view_distance,
        sight_direction=sight_direction,
        up_direction=up_direction,
        image_format=WHEELHOUSE_SCREENSHOT_IMAGE_FORMAT,
        wait_seconds=0.5,
    )
    return shot_result


def load_capture_document_view() -> Any:
    try:
        from catia_picture_capture import capture_document_view

        return capture_document_view
    except Exception:
        tool_path = Path(__file__).resolve().with_name("catia_picture_capture.py")
        spec = importlib.util.spec_from_file_location("catia_picture_capture_local", tool_path)
        if spec is None or spec.loader is None:
            raise RuntimeError(f"无法加载截图工具文件: {tool_path}")
        capture_tool = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(capture_tool)
        return capture_tool.capture_document_view


def load_reframe_document_root() -> Any:
    try:
        from catia_picture_capture import reframe_document_root

        return reframe_document_root
    except Exception:
        tool_path = Path(__file__).resolve().with_name("catia_picture_capture.py")
        spec = importlib.util.spec_from_file_location("catia_picture_capture_local", tool_path)
        if spec is None or spec.loader is None:
            raise RuntimeError(f"无法加载截图工具文件: {tool_path}")
        capture_tool = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(capture_tool)
        return capture_tool.reframe_document_root


def center_wheelhouse_product_root(
    catia: Any,
    product_document: Any,
) -> dict[str, Any]:
    try:
        reframe_document_root = load_reframe_document_root()
        return reframe_document_root(
            catia,
            product_document,
            clear_selection_after=True,
            wait_seconds=0.5,
        )
    except Exception as exc:
        return {
            "status": "failed",
            "message": str(exc),
        }


def capture_wheelhouse_regulation_screenshots(
    catia: Any,
    product_document: Any,
    root_product: Any,
    regulation_axis_part: dict[str, Any] | None,
    annotation_result: dict[str, Any] | None,
    result_dir: str | Path,
) -> dict[str, Any]:
    screenshot_dir = Path(result_dir) / "screenshots"
    screenshot_dir.mkdir(parents=True, exist_ok=True)
    annotation_items = list((annotation_result or {}).get("annotation_items") or [])
    if not annotation_items:
        return {
            "status": "skipped",
            "message": "未找到法规标注点，跳过截图。",
            "screenshot_dir": str(screenshot_dir),
            "screenshots": [],
        }
    state_rows: dict[int, dict[str, Any]] = {}
    screenshots: list[dict[str, Any]] = []
    scene_rows: list[dict[str, Any]] = []
    restore_rows: list[dict[str, Any]] = []
    post_restore_save_result: dict[str, Any] | None = None
    try:
        product_document.Activate()
    except Exception:
        pass

    components = collect_reloaded_wheelhouse_components(root_product, regulation_axis_part)
    annotation_component = components.get("annotation")
    annotation_part = None
    if annotation_component is not None:
        try:
            annotation_part, _annotation_part_document = get_part_and_document_from_product(annotation_component)
        except Exception:
            annotation_part = None
    specs = build_wheelhouse_screenshot_specs(annotation_items, components, screenshot_dir)
    if not specs:
        return {
            "status": "skipped",
            "message": "未生成有效截图任务。",
            "screenshot_dir": str(screenshot_dir),
            "screenshots": [],
        }

    try:
        for key in ("front", "rear", "axis", "wheel", "annotation"):
            component = components.get(key)
            if component is not None:
                record_visibility_state(
                    product_document,
                    component,
                    state_rows,
                    kind="product",
                    label=f"component:{key}",
                )
        for section_component in components.get("all_sections") or []:
            record_visibility_state(
                product_document,
                section_component,
                state_rows,
                kind="product",
                label=f"section:{product_part_number(section_component)}",
            )
        if annotation_part is not None:
            for feature_name, feature in collect_annotation_features_by_name(annotation_part).items():
                record_visibility_state(
                    product_document,
                    feature,
                    state_rows,
                    kind="object",
                    label=f"annotation_feature:{feature_name}",
                )
        for marker_name, marker in collect_marker3d_by_name(root_product).items():
            if marker_name.endswith("_Text"):
                record_visibility_state(
                    product_document,
                    marker,
                    state_rows,
                    kind="object",
                    label=f"annotation_text:{marker_name}",
                )

        specs_by_key = {str(spec.get("measurement_key")): spec for spec in specs}

        def run_capture_for_spec(spec: dict[str, Any]) -> None:
            try:
                try:
                    root_product.Update()
                except Exception:
                    pass
                shot_result = refresh_and_capture_document_view(
                    catia,
                    product_document,
                    spec["output_path"],
                    view_point=spec["view_point"],
                    view_distance=spec["view_distance"],
                    sight_direction=spec["sight_direction"],
                    up_direction=spec["up_direction"],
                )
                screenshots.append(
                    {
                        "status": "success",
                        "measurement_key": spec["measurement_key"],
                        "category": spec["category"],
                        "annotation_name": spec["annotation_name"],
                        "path": shot_result.get("screenshot_path"),
                        "view": shot_result.get("view"),
                    }
                )
            except Exception as exc:
                screenshots.append(
                    {
                        "status": "failed",
                        "measurement_key": spec["measurement_key"],
                        "category": spec["category"],
                        "annotation_name": spec["annotation_name"],
                        "path": spec["output_path"],
                        "message": str(exc),
                    }
                )

        for spec in specs:
            measurement_key = str(spec["measurement_key"])
            wheelhouse_label = str(spec.get("wheelhouse") or "")
            scene_rows.extend(
                set_wheelhouse_stage_visibility(
                    product_document,
                    root_product,
                    components,
                    annotation_part,
                    wheelhouse_label,
                    measurement_key,
                    str(spec.get("category") or ""),
                    state_rows,
                )
            )
            run_capture_for_spec(spec)
            if annotation_part is not None:
                scene_rows.extend(
                    reset_annotation_text_visibility(
                        product_document,
                        root_product,
                        annotation_part,
                        state_rows,
                    )
                )
    finally:
        restore_rows = restore_recorded_visibility(product_document, state_rows)
        try:
            root_product.Update()
        except Exception:
            pass
        try:
            post_restore_save_result = save_document_if_modified(product_document)
        except Exception as exc:
            post_restore_save_result = {"status": "failed", "message": str(exc)}

    success_count = sum(1 for row in screenshots if row.get("status") == "success")
    return {
        "status": "success" if success_count == len(screenshots) else "partial_failed",
        "screenshot_dir": str(screenshot_dir),
        "count": len(screenshots),
        "success_count": success_count,
        "screenshots": screenshots,
        "scene_visibility_rows": scene_rows,
        "restore_rows": restore_rows,
        "post_restore_save_result": post_restore_save_result,
    }


def axis_records_to_dicts(records: list[AxisRecord]) -> list[dict[str, Any]]:
    """
    功能: 将旧版轴记录转为字典。
    输入: AxisRecord 列表。
    输出: 字典列表。
    """
    return [
        {
            "direction": tuple(round(value, 8) for value in record.direction),
            "feature_name": record.feature_name,
            "component_path": record.component_path,
            "component_name": record.component_name,
            "component_part_number": record.component_part_number,
        }
        for record in records
    ]


def run_wheelhouse_regulation_verification(
    front_wheelhouse_part_path: str | Path | None = None,
    rear_wheelhouse_part_path: str | Path | None = None,
    wheel_assembly_path: str | Path | None = None,
    output_dir: str | Path | None = None,
    product_part_number: str = DEFAULT_PRODUCT_PART_NUMBER,
    product_name: str = DEFAULT_PRODUCT_NAME,
    axis_direction_tolerance_degrees: float = AXIS_DIRECTION_TOLERANCE_DEGREES,
    wheel_position_tolerance: float = WHEEL_POSITION_TOLERANCE,
    tire_hub_cog_tolerance: float = TIRE_HUB_COG_TOLERANCE,
    save_product: bool = True,
    left_front_wheelhouse_part_path: str | Path | None = None,
    right_front_wheelhouse_part_path: str | Path | None = None,
    left_rear_wheelhouse_part_path: str | Path | None = None,
    right_rear_wheelhouse_part_path: str | Path | None = None,
) -> dict[str, Any]:
    """
    功能: 执行完整车轮罩法规检测流程。
    输入: 1-4 个轮罩、车轮装配路径和配置。
    输出: 检测结果字典。
    """
    run_output_paths = create_run_output_paths(output_dir)
    print(f"运行输出目录: {run_output_paths.run_root_dir}")
    print(f"结果输出目录: {run_output_paths.result_dir}")
    print(f"过程输出目录: {run_output_paths.course_dir}")

    if wheel_assembly_path is None:
        raise RuntimeError("未提供车轮装配路径。")
    wheel_path = validate_existing_catia_file(wheel_assembly_path, WHEEL_ASSEMBLY_LABEL)

    wheelhouse_input_specs: list[WheelhouseInput] = []
    if any(
        value is not None
        for value in (
            left_front_wheelhouse_part_path,
            right_front_wheelhouse_part_path,
            left_rear_wheelhouse_part_path,
            right_rear_wheelhouse_part_path,
        )
    ):
        multi_inputs = [
            ("left_front", left_front_wheelhouse_part_path),
            ("right_front", right_front_wheelhouse_part_path),
            ("left_rear", left_rear_wheelhouse_part_path),
            ("right_rear", right_rear_wheelhouse_part_path),
        ]
        for slot, path in multi_inputs:
            if path is None:
                continue
            validated_path = validate_existing_catia_file(path, wheelhouse_definition_from_label(WHEELHOUSE_SLOT_DEFINITIONS[slot]["label"])["label"])
            wheelhouse_input_specs.append(
                wheelhouse_input_from_path(
                    slot,
                    validated_path,
                    copy_wheelhouse_to_temp(
                        validated_path,
                        run_output_paths.result_dir,
                        WHEELHOUSE_SLOT_DEFINITIONS[slot]["label"],
                        timestamp=run_output_paths.timestamp,
                    ),
                )
            )
    else:
        if front_wheelhouse_part_path is not None:
            front_path = validate_existing_catia_file(front_wheelhouse_part_path, FRONT_WHEELHOUSE_LABEL)
            wheelhouse_input_specs.append(
                legacy_wheelhouse_input(
                    "front",
                    FRONT_WHEELHOUSE_LABEL,
                    front_path,
                    copy_wheelhouse_to_temp(
                        front_path,
                        run_output_paths.result_dir,
                        FRONT_WHEELHOUSE_LABEL,
                        timestamp=run_output_paths.timestamp,
                    ),
                )
            )
        if rear_wheelhouse_part_path is not None:
            rear_path = validate_existing_catia_file(rear_wheelhouse_part_path, REAR_WHEELHOUSE_LABEL)
            wheelhouse_input_specs.append(
                legacy_wheelhouse_input(
                    "rear",
                    REAR_WHEELHOUSE_LABEL,
                    rear_path,
                    copy_wheelhouse_to_temp(
                        rear_path,
                        run_output_paths.result_dir,
                        REAR_WHEELHOUSE_LABEL,
                        timestamp=run_output_paths.timestamp,
                    ),
                )
            )
    if not wheelhouse_input_specs:
        raise RuntimeError("未提供可用于校核的轮罩零件。")

    catia = start_or_connect_catia()
    (
        product_document,
        root_product,
        wheelhouse_components,
        wheel_assembly_component,
    ) = create_wheelhouse_product_from_inputs(
        catia,
        wheelhouse_input_specs,
        wheel_path,
        product_part_number=product_part_number,
        product_name=product_name,
    )

    initial_root_reframe_result = center_wheelhouse_product_root(
        catia,
        product_document,
    )

    product_save_path = build_product_save_path(
        run_output_paths.result_dir,
        timestamp=run_output_paths.timestamp,
    )
    hidden_wheel_part_rows: list[dict[str, Any]] = []
    wheel_part_contexts = iter_leaf_part_contexts(
        wheel_assembly_component,
        WHEEL_ASSEMBLY_LABEL,
        visibility_document=product_document,
        visible_only=True,
        visibility_skipped=hidden_wheel_part_rows,
    )
    if not wheel_part_contexts:
        raise RuntimeError("车轮装配中未找到可分析的零件。")

    warnings: list[str] = []
    if hidden_wheel_part_rows:
        warnings.append(
            f"遍历车轮装配时跳过隐藏零件/子装配 {len(hidden_wheel_part_rows)} 个，"
            "这些对象未参与重心、轮罩近邻分组和包围盒测量。"
        )
    wheelhouse_cog_items = evaluate_wheelhouse_cog_items(
        product_document,
        wheelhouse_components,
    )
    (
        wheel_candidates,
        excluded_no_axis_components,
        candidate_warnings,
        wheelhouse_part_proximity,
        near_part_visibility_rows,
    ) = build_wheel_candidates(
        product_document,
        root_product,
        wheel_part_contexts,
        wheelhouse_cog_items=wheelhouse_cog_items,
    )
    warnings.extend(candidate_warnings)
    if not wheel_candidates:
        raise RuntimeError("车轮装配中未找到可用于筛选的车轮候选件。")

    wheel_position_groups = cluster_wheel_candidates(
        wheel_candidates,
        axis_tolerance_degrees=axis_direction_tolerance_degrees,
        position_tolerance=wheel_position_tolerance,
    )
    tire_representatives = choose_tire_representatives(
        wheel_position_groups,
        tire_hub_cog_tolerance=tire_hub_cog_tolerance,
    )
    for group in wheel_position_groups:
        warnings.extend(f"{group.group_id}: {warning}" for warning in group.warnings)
    if not tire_representatives:
        raise RuntimeError("车轮装配中未筛选出 Tire 代表件。")

    matches, wheelhouse_cogs, tire_cogs, match_warnings = match_wheelhouses_to_tire_representatives(
        product_document,
        wheelhouse_components,
        None,
        tire_representatives,
        wheelhouse_cog_items=wheelhouse_cog_items,
    )
    warnings.extend(match_warnings)

    tire_candidate_by_path = {
        candidate.context.component_path: candidate
        for candidate in tire_representatives
    }
    regulation_axis_segments: list[RegulationAxisSegment] = []
    for match in matches:
        tire_component_path = match.get("tire_component_path")
        tire_candidate = tire_candidate_by_path.get(tire_component_path)
        if tire_candidate is None:
            warnings.append(f"{match.get('wheelhouse')}: 未找到匹配 Tire 候选对象，无法创建法规校核轴线段。")
            continue
        regulation_axis_segments.append(
            build_regulation_axis_segment(
                match["wheelhouse"],
                tire_candidate,
                REGULATION_AXIS_HALF_LENGTH,
                REGULATION_AXIS_EXTRUDE_LENGTH,
                REGULATION_AXIS_ROTATION_ANGLE,
            ),
        )
    annotation_runtime_context = create_wheelhouse_annotation_runtime_context(run_output_paths.result_dir)
    regulation_axis_part = create_regulation_axis_part(
        product_document=product_document,
        root_product=root_product,
        front_component=wheelhouse_components,
        rear_component=None,
        segments=regulation_axis_segments,
        wheelhouse_bounding_boxes=None,
        output_dir=run_output_paths.result_dir,
        course_dir=run_output_paths.course_dir,
        run_timestamp=run_output_paths.timestamp,
        annotation_runtime_context=annotation_runtime_context,
    )
    wheelhouse_extreme_bounding_boxes = (
        regulation_axis_part.get("wheelhouse_extreme_bounding_boxes") or {}
    )

    visibility_rows = [
        *near_part_visibility_rows,
        hide_wheel_assembly_after_regulation_geometry(
            product_document,
            wheel_assembly_component,
        )
    ]
    annotation_result = wheelhouse_runtime_annotation_result(annotation_runtime_context)
    screenshot_result = wheelhouse_runtime_screenshot_result(annotation_runtime_context)
    print("\n[9] 创建标注")
    print(
        f"[9] 创建标注完成 status={annotation_result.get('status')} "
        f"success={annotation_result.get('success_count', 0)}/"
        f"{len(annotation_result.get('annotation_items') or [])}"
    )
    print("\n[10] 截图")
    print(
        f"[10] 截图完成 status={screenshot_result.get('status')} "
        f"success={screenshot_result.get('success_count', 0)}/"
        f"{screenshot_result.get('count', 0)}"
    )
    if annotation_result.get("status") in {"failed", "partial_failed"}:
        warnings.append(
            f"法规标注创建未完全成功: {annotation_result.get('status')}"
        )

    component_save_results: list[dict[str, Any]] = []
    annotation_product_reopen_result: dict[str, Any] | None = None
    try:
        final_components = collect_reloaded_wheelhouse_components(root_product, regulation_axis_part)
        final_annotation_part = None
        final_annotation_part_document = None
        final_annotation_component = final_components.get("annotation")
        if final_annotation_component is not None:
            try:
                final_annotation_part, final_annotation_part_document = get_part_and_document_from_product(final_annotation_component)
            except Exception:
                final_annotation_part = None
                final_annotation_part_document = None
        final_visibility_rows = set_wheelhouse_final_visibility(
            product_document,
            root_product,
            final_components,
            final_annotation_part,
            final_annotation_part_document,
            {},
        )
        visibility_rows.extend(final_visibility_rows)
        center_wheelhouse_product_root(catia, product_document)
    except Exception as exc:
        warnings.append(f"最终显隐恢复失败: {exc}")

    if save_product:
        try:
            component_save_results = save_modified_component_documents(
                root_product,
                exclude_documents=[product_document],
            )
            product_document.Activate()
            product_document.SaveAs(str(product_save_path))
            saved_product_path: str | None = str(product_save_path)
        except Exception as exc:
            warnings.append(f"CATProduct 保存失败，流程结果仍返回: {exc}")
            saved_product_path = None
    else:
        saved_product_path = None

    try:
        product_document.Activate()
        product_document.Application.ActiveWindow.ActiveViewer.Reframe()
        product_document.Application.ActiveWindow.ActiveViewer.Update()
    except Exception:
        pass

    return {
        "success": True,
        "product_path": saved_product_path,
        "root_product_part_number": product_part_number,
        "root_product_name": product_name,
        "input_paths": {
            item.label: str(item.path) for item in wheelhouse_input_specs
        } | {
            "wheel_assembly": str(wheel_path),
        },
        "work_paths": {
            item.label: str(item.work_path) if item.work_path is not None else None
            for item in wheelhouse_input_specs
        },
        "run_output_paths": {
            "timestamp": run_output_paths.timestamp,
            "run_root_dir": str(run_output_paths.run_root_dir),
            "result_dir": str(run_output_paths.result_dir),
            "course_dir": str(run_output_paths.course_dir),
            "default_json_result_path": str(run_output_paths.json_result_path),
        },
        "wheelhouse_inputs": [
            {
                "slot": item.slot,
                "label": item.label,
                "measurement_prefix": item.measurement_prefix,
                "component_part_number": item.component_part_number,
                "geometry_set_name": item.geometry_set_name,
                "line_name": item.line_name,
                "section_prefix": item.section_prefix,
                "position": item.position,
                "side": item.side,
                "path": str(item.path),
                "work_path": str(item.work_path) if item.work_path is not None else None,
            }
            for item in wheelhouse_input_specs
        ],
        "initial_root_reframe_result": initial_root_reframe_result,
        "wheel_part_count": len(wheel_part_contexts),
        "hidden_wheel_part_count": len(hidden_wheel_part_rows),
        "hidden_wheel_parts_skipped_before_cog": hidden_wheel_part_rows,
        "wheelhouse_near_distance": WHEELHOUSE_NEAR_DISTANCE,
        "wheelhouse_part_proximity": wheelhouse_part_proximity,
        "axis_direction_tolerance_degrees": axis_direction_tolerance_degrees,
        "wheel_position_tolerance": wheel_position_tolerance,
        "tire_hub_cog_tolerance": tire_hub_cog_tolerance,
        "excluded_no_axis_count": len(excluded_no_axis_components),
        "excluded_no_axis_components": excluded_no_axis_components,
        "excluded_non_tire_components": excluded_no_axis_components,
        "wheel_candidate_count": len(wheel_candidates),
        "wheel_candidates": [candidate_to_dict(candidate) for candidate in wheel_candidates],
        "wheel_position_group_count": len(wheel_position_groups),
        "wheel_position_groups": [
            wheel_position_group_to_dict(group)
            for group in wheel_position_groups
        ],
        "tire_representative_count": len(tire_representatives),
        "tire_representatives": [
            candidate_to_dict(candidate)
            for candidate in tire_representatives
        ],
        "excluded_hub_candidates": [
            candidate_to_dict(candidate)
            for group in wheel_position_groups
            for candidate in group.excluded_candidates
        ],
        "visibility": visibility_rows,
        "annotation_result": annotation_result,
        "annotation_product_reopen_result": strip_private_result_fields(annotation_product_reopen_result),
        "screenshot_result": screenshot_result,
        "component_save_results": component_save_results,
        "wheelhouse_cogs": wheelhouse_cogs,
        "wheelhouse_extreme_bounding_boxes": wheelhouse_extreme_bounding_boxes,
        "tire_cogs": tire_cogs,
        "matches": matches,
        "regulation_axis_part": regulation_axis_part,
        "regulation_axis_segments": [
            regulation_axis_segment_to_dict(segment)
            for segment in regulation_axis_segments
        ],
        "warnings": warnings,
    }


def print_result_summary(result: dict[str, Any]) -> None:
    """
    功能: 打印检测结果摘要。
    输入: run_wheelhouse_regulation_verification 返回的结果字典。
    输出: 终端文本。
    """
    print("========== 车轮罩法规校核 ==========")
    print(f"CATProduct: {result.get('product_path') or '未保存'}")
    print(f"车轮装配零件数量: {result.get('wheel_part_count')}")
    print(f"非轮胎排除零件数量: {result.get('excluded_no_axis_count')}")
    print(f"车轮候选数量: {result.get('wheel_candidate_count')}")
    print(f"车轮位置组数量: {result.get('wheel_position_group_count')}")
    print(f"Tire代表件数量: {result.get('tire_representative_count')}")
    print(f"筛除Hub/Rim候选数量: {len(result.get('excluded_hub_candidates', []))}")

    wheelhouse_bboxes = result.get("wheelhouse_extreme_bounding_boxes") or {}
    if wheelhouse_bboxes:
        print("\n-- 轮罩极值包围盒 --")
        for label, row in wheelhouse_bboxes.items():
            bbox = row.get("bbox_world") or {}
            if row.get("status") != "success" or not bbox:
                print(f"{label}: failed")
                continue
            print(
                f"{label}: min={bbox.get('min_point')}, "
                f"max={bbox.get('max_point')}, size={bbox.get('size')}, "
                f"diagonal={bbox.get('diagonal')}"
            )

    warnings = result.get("warnings") or []
    for warning in warnings:
        print(f"[警告] {warning}")

    print("\n-- Tire代表件 --")
    for index, tire in enumerate(result.get("tire_representatives", []), start=1):
        print(
            f"{index}. {tire.get('component_path')} / {tire.get('feature_name')} "
            f"轴来源={tire.get('axis_source')}, "
            f"重心={tire.get('component_cog_world')}, "
            f"包围盒对角线={((tire.get('bbox_world') or {}).get('diagonal'))}, "
            f"拓扑圆={tire.get('topology_circle_count')}, "
            f"半径层级={tire.get('topology_radius_levels')}"
        )
    if not result.get("tire_representatives"):
        print("无")

    print("\n-- 非轮胎排除零件 --")
    no_axis_items = result.get("excluded_no_axis_components", [])
    for index, item in enumerate(no_axis_items, start=1):
        print(
            f"{index}. {item.get('component_path')} "
            f"PartNumber={item.get('component_part_number')}"
        )
    if not no_axis_items:
        print("无")

    print("\n-- 筛除Hub/Rim候选 --")
    excluded = result.get("excluded_hub_candidates", [])
    for index, candidate in enumerate(excluded, start=1):
        print(f"{index}. {candidate.get('component_path')}")
    if not excluded:
        print("无")

    print("\n-- 显示状态 --")
    for row in result.get("visibility", []):
        state = "显示" if row.get("visible") else "隐藏"
        print(f"{state}: {row.get('component_path')}")

    print("\n-- 车轮罩与Tire匹配结果 --")
    for match in result.get("matches", []):
        print(
            f"{match['wheelhouse']} -> {match['tire_component_path']} "
            f"score={match['score']}, cog_distance={match['cog_distance']} mm"
        )

    print("\n-- 法规校核轴线段 --")
    regulation_axis_part = result.get("regulation_axis_part") or {}
    if regulation_axis_part:
        print(
            f"CATPart: {regulation_axis_part.get('component_part_number')} "
            f"几何图形集={regulation_axis_part.get('geometry_set_names')}"
        )
    for segment in result.get("regulation_axis_segments", []):
        print(
            f"{segment.get('wheelhouse')} / {segment.get('geometry_set_name')} / "
            f"{segment.get('line_name')} length={segment.get('axis_length')} mm"
        )
    if not result.get("regulation_axis_segments"):
        print("无")
    bbox_wireframe_results = (regulation_axis_part or {}).get("bbox_wireframe_results") or []
    if bbox_wireframe_results:
        print("\n-- 轮罩包围盒线框 --")
        for row in bbox_wireframe_results:
            if row.get("status") == "success":
                print(
                    f"{row.get('wheelhouse')} -> {row.get('geometry_set_name')} "
                    f"size={row.get('size')} max={row.get('max_size')} mm"
                )
            else:
                print(f"{row.get('wheelhouse')} failed: {row.get('message')}")

    print("\n-- 法规校核截面曲线 --")
    section_curve_results = (regulation_axis_part or {}).get("section_curve_results", [])
    section_topology_groups: list[dict[str, Any]] = []
    regulation_distance_measurements: dict[str, Any] = dict(
        (regulation_axis_part or {}).get("bbox_measurements") or {}
    )
    for item in section_curve_results:
        if item.get("section_topology_groups"):
            section_topology_groups.extend(item.get("section_topology_groups") or [])
            regulation_distance_measurements.update(item.get("regulation_distance_measurements") or {})
            continue
        print(
            f"{item.get('section_plane_name')} -> {item.get('section_curve_name')} "
            f"target={item.get('target_name')} status={item.get('status')}"
        )
        if item.get("exported_catpart"):
            print(f"  导出: {item.get('exported_catpart')}")
        exported_component = item.get("exported_component") or {}
        if exported_component:
            print(f"  装配: {exported_component.get('component_part_number')}")
        if item.get("section_source_path"):
            print(f"  截面工作文件: {item.get('section_source_path')}")
        if item.get("status") != "success" and item.get("message"):
            print(f"  原因: {item.get('message')}")
    if not section_curve_results:
        print("无")

    print("\n-- 法规校核截面拓扑分组 --")
    if section_topology_groups:
        for group in section_topology_groups:
            print(
                f"{group.get('section_plane_name')} group={group.get('group_index')} "
                f"curves={group.get('curve_count')} "
                f"direction_span={group.get('direction_projection_span', group.get('z_height'))} mm "
                f"distance={group.get('direction_distance', group.get('z_distance'))} mm"
            )
    else:
        print("无")

    print("\n-- 法规校核距离测量 --")
    if regulation_distance_measurements:
        for key in WHEELHOUSE_ANNOTATION_KEYS:
            value = regulation_distance_measurements.get(key)
            if value is not None:
                try:
                    print(f"{key}: {float(value):.2f} mm")
                except Exception:
                    print(f"{key}: {value} mm")
    else:
        print("无")

    annotation_result = result.get("annotation_result") or {}
    if annotation_result:
        print("\n-- 法规标注 --")
        print(
            f"status={annotation_result.get('status')} "
            f"success={annotation_result.get('success_count', 0)}/"
            f"{len(annotation_result.get('annotation_items') or [])} "
            f"part={annotation_result.get('process_part_save_path')}"
        )
    annotation_product_reopen_result = result.get("annotation_product_reopen_result") or {}
    if annotation_product_reopen_result:
        print(
            "标注Product重开刷新: "
            f"status={annotation_product_reopen_result.get('status')} "
            f"path={annotation_product_reopen_result.get('path')}"
        )
    screenshot_result = result.get("screenshot_result") or {}
    if screenshot_result:
        print(
            "法规截图: "
            f"status={screenshot_result.get('status')} "
            f"success={screenshot_result.get('success_count', 0)}/"
            f"{screenshot_result.get('count', 0)} "
            f"dir={screenshot_result.get('screenshot_dir')}"
        )


def write_json_result(result_path: str | Path | None, result: dict[str, Any]) -> None:
    """
    功能: 写入 JSON 结果文件。
    输入: 输出路径和结果字典。
    输出: JSON 文件，路径为空时不写。
    """
    if not result_path:
        return
    json_path = Path(result_path).expanduser().resolve()
    json_path.parent.mkdir(parents=True, exist_ok=True)
    json_path.write_text(
        json.dumps(make_json_safe(result), ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


def make_json_safe(value: Any) -> Any:
    """
    功能: 将结果数据转换为 JSON 可序列化对象。
    输入: 任意结果值。
    输出: 可 json.dumps 的值。
    """
    if value is None or isinstance(value, (str, int, float, bool)):
        return value
    if isinstance(value, Path):
        return str(value)
    if isinstance(value, tuple):
        return [make_json_safe(item) for item in value]
    if isinstance(value, list):
        return [make_json_safe(item) for item in value]
    if isinstance(value, dict):
        return {
            str(key): make_json_safe(item)
            for key, item in value.items()
            if key not in {"object", "element1", "element2", "source_element1", "source_element2"}
        }
    if hasattr(value, "Name"):
        return safe_attr_text(value, "Name", repr(value))
    return repr(value)


def resolve_json_result_path(result_path: str | Path | None, result: dict[str, Any]) -> str | Path | None:
    """
    功能: 获取 JSON 结果输出路径，未指定时使用本次运行 result 目录默认路径。
    输入: 用户指定路径和流程结果。
    输出: JSON 路径或 None。
    """
    if result_path:
        return result_path
    run_output_paths = result.get("run_output_paths") or {}
    return run_output_paths.get("default_json_result_path")


def main(
    front_wheelhouse_part_path: str | Path = FRONT_WHEELHOUSE_PART_PATH,
    rear_wheelhouse_part_path: str | Path = REAR_WHEELHOUSE_PART_PATH,
    wheel_assembly_path: str | Path = WHEEL_ASSEMBLY_PATH,
    output_dir: str | Path | None = OUTPUT_DIR,
    json_result_path: str | Path | None = JSON_RESULT_PATH,
    product_part_number: str = PRODUCT_PART_NUMBER,
    product_name: str = PRODUCT_NAME,
    axis_direction_tolerance_degrees: float = AXIS_TOLERANCE_DEGREES,
    wheel_position_tolerance: float = WHEEL_POSITION_CLUSTER_TOLERANCE,
    tire_hub_cog_tolerance: float = TIRE_HUB_CENTER_TOLERANCE,
    save_product: bool = SAVE_PRODUCT_FILE,
    left_front_wheelhouse_part_path: str | Path | None = None,
    right_front_wheelhouse_part_path: str | Path | None = None,
    left_rear_wheelhouse_part_path: str | Path | None = None,
    right_rear_wheelhouse_part_path: str | Path | None = None,
) -> dict[str, Any]:
    """
    功能: 轮罩法规检测主函数。
    输入: 前轮罩路径、后轮罩路径、车轮装配路径、输出路径和筛选容差配置。
    输出: 检测结果字典，并在 CATIA 中生成/保存校核 CATProduct。
    """
    configure_console_encoding()
    result = run_wheelhouse_regulation_verification(
        front_wheelhouse_part_path=front_wheelhouse_part_path,
        rear_wheelhouse_part_path=rear_wheelhouse_part_path,
        wheel_assembly_path=wheel_assembly_path,
        output_dir=output_dir,
        product_part_number=product_part_number,
        product_name=product_name,
        axis_direction_tolerance_degrees=axis_direction_tolerance_degrees,
        wheel_position_tolerance=wheel_position_tolerance,
        tire_hub_cog_tolerance=tire_hub_cog_tolerance,
        save_product=save_product,
        left_front_wheelhouse_part_path=left_front_wheelhouse_part_path,
        right_front_wheelhouse_part_path=right_front_wheelhouse_part_path,
        left_rear_wheelhouse_part_path=left_rear_wheelhouse_part_path,
        right_rear_wheelhouse_part_path=right_rear_wheelhouse_part_path,
    )
    print_result_summary(result)
    write_json_result(resolve_json_result_path(json_result_path, result), result)
    return result


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    """
    功能: 解析命令行参数。
    输入: argv 或 None。
    输出: argparse.Namespace。
    """
    parser = argparse.ArgumentParser(
        description=(
            "创建车轮罩法规校核 CATProduct，并按车轮轴线位置聚类、筛选Tire代表件、匹配车轮罩。"
            "不传命令行参数时，将使用 main.py 顶部固定路径变量。"
        ),
    )
    parser.add_argument("front_wheelhouse_part_path", nargs="?", default=None, help="前轮罩 CATPart 路径（兼容旧模式）")
    parser.add_argument("rear_wheelhouse_part_path", nargs="?", default=None, help="后轮罩 CATPart 路径（兼容旧模式）")
    parser.add_argument("wheel_assembly_path", nargs="?", default=None, help="车轮装配 CATProduct/CATPart 路径（兼容旧模式）")
    parser.add_argument("--wheel-assembly-path", dest="wheel_assembly_path_flag", default=None, help="车轮装配 CATProduct/CATPart 路径")
    parser.add_argument("--left-front-wheelhouse-part-path", dest="left_front_wheelhouse_part_path", default=None, help="左前轮罩 CATPart 路径")
    parser.add_argument("--right-front-wheelhouse-part-path", dest="right_front_wheelhouse_part_path", default=None, help="右前轮罩 CATPart 路径")
    parser.add_argument("--left-rear-wheelhouse-part-path", dest="left_rear_wheelhouse_part_path", default=None, help="左后轮罩 CATPart 路径")
    parser.add_argument("--right-rear-wheelhouse-part-path", dest="right_rear_wheelhouse_part_path", default=None, help="右后轮罩 CATPart 路径")
    parser.add_argument("--output-dir", default=None, help="CATProduct 输出目录")
    parser.add_argument(
        "--axis-tolerance-deg",
        type=float,
        default=AXIS_DIRECTION_TOLERANCE_DEGREES,
        help="轴线平行判断角度容差，默认 2 度",
    )
    parser.add_argument(
        "--wheel-position-tolerance",
        type=float,
        default=WHEEL_POSITION_TOLERANCE,
        help="同一车轮位置聚类距离容差，默认 50 mm",
    )
    parser.add_argument(
        "--tire-hub-cog-tolerance",
        type=float,
        default=TIRE_HUB_COG_TOLERANCE,
        help="轮胎和轮毂重心一致性判断容差，默认 30 mm",
    )
    parser.add_argument("--json-result", default=None, help="可选：将结果写入 JSON 文件")
    parser.add_argument("--no-save", action="store_true", help="只运行分析，不保存 CATProduct")
    return parser.parse_args(argv)


def cli_main(argv: list[str] | None = None) -> int:
    """
    功能: 命令行入口。
    输入: argv 或 None。
    输出: 进程退出码。
    """
    configure_console_encoding()
    effective_argv = sys.argv[1:] if argv is None else argv
    try:
        if not effective_argv:
            main()
            return 0

        args = parse_args(effective_argv)
        wheel_assembly_path = args.wheel_assembly_path_flag or args.wheel_assembly_path
        if wheel_assembly_path is None:
            raise ValueError("未提供车轮装配路径。")
        result = run_wheelhouse_regulation_verification(
            front_wheelhouse_part_path=args.front_wheelhouse_part_path,
            rear_wheelhouse_part_path=args.rear_wheelhouse_part_path,
            wheel_assembly_path=wheel_assembly_path,
            output_dir=args.output_dir,
            axis_direction_tolerance_degrees=args.axis_tolerance_deg,
            wheel_position_tolerance=args.wheel_position_tolerance,
            tire_hub_cog_tolerance=args.tire_hub_cog_tolerance,
            save_product=not args.no_save,
            left_front_wheelhouse_part_path=args.left_front_wheelhouse_part_path,
            right_front_wheelhouse_part_path=args.right_front_wheelhouse_part_path,
            left_rear_wheelhouse_part_path=args.left_rear_wheelhouse_part_path,
            right_rear_wheelhouse_part_path=args.right_rear_wheelhouse_part_path,
        )
        print_result_summary(result)
        write_json_result(resolve_json_result_path(args.json_result, result), result)
        return 0
    except Exception as exc:
        print(f"车轮罩法规校核失败: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(cli_main())
