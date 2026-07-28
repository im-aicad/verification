from __future__ import annotations

import argparse
import itertools
import json
import math
import sys
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any, Iterable

import win32com.client


Vector = tuple[float, float, float]


DEFAULT_PRODUCT_PART_NUMBER = "Wheelhouse_Regulation_Verification"
DEFAULT_PRODUCT_NAME = "Wheelhouse_Regulation_Verification"
DEFAULT_OUTPUT_DIR = Path(__file__).resolve().parent / "catia_project"
AXIS_DIRECTION_TOLERANCE_DEGREES = 2.0
WHEEL_POSITION_TOLERANCE = 50.0
TIRE_HUB_COG_TOLERANCE = 30.0
TIRE_SIZE_ADVANTAGE_RATIO = 1.05

# ==================== 变量命名区 ====================
# 直接运行本文件时，修改下面三个固定路径即可执行轮罩法规检测主程序。
FRONT_WHEELHOUSE_PART_PATH: str | Path = r"C:\\Users\\Administrator\\Desktop\\catia_project_test\\wheel_house\\11954872_05.CATPart"
REAR_WHEELHOUSE_PART_PATH: str | Path = r"C:\\Users\\Administrator\\Desktop\\catia_project_test\\wheel_house\\12128341_06.CATPart"
WHEEL_ASSEMBLY_PATH: str | Path = r"C:\\Users\\Administrator\\Desktop\\catia_project_test\\wheel\\11940666_04.CATProduct"

OUTPUT_DIR: str | Path | None = DEFAULT_OUTPUT_DIR
JSON_RESULT_PATH: str | Path | None = None
SAVE_PRODUCT_FILE = True

PRODUCT_PART_NUMBER = DEFAULT_PRODUCT_PART_NUMBER
PRODUCT_NAME = DEFAULT_PRODUCT_NAME
AXIS_TOLERANCE_DEGREES = AXIS_DIRECTION_TOLERANCE_DEGREES
WHEEL_POSITION_CLUSTER_TOLERANCE = WHEEL_POSITION_TOLERANCE
TIRE_HUB_CENTER_TOLERANCE = TIRE_HUB_COG_TOLERANCE
# ================== 变量命名区结束 ==================

FRONT_WHEELHOUSE_LABEL = "Front_Wheelhouse"
REAR_WHEELHOUSE_LABEL = "Rear_Wheelhouse"
WHEEL_ASSEMBLY_LABEL = "Wheel_Assembly"

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
    axis_direction_world: Vector | None
    axis_point_world: Vector | None
    component_cog_world: Vector
    bbox_world: BoundingBox | None
    mass: float | None
    volume: float | None
    warnings: list[str]


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


def build_product_save_path(output_dir: str | Path | None = None) -> Path:
    """
    功能: 构造校核 CATProduct 保存路径。
    输入: 可选输出目录。
    输出: CATProduct 文件路径。
    """
    target_dir = Path(output_dir).expanduser().resolve() if output_dir else DEFAULT_OUTPUT_DIR
    target_dir.mkdir(parents=True, exist_ok=True)
    return target_dir / f"wheelhouse_regulation_verification_{build_timestamp()}.CATProduct"


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


def create_wheelhouse_product(
    catia: Any,
    front_wheelhouse_path: Path,
    rear_wheelhouse_path: Path,
    wheel_assembly_path: Path,
    product_part_number: str = DEFAULT_PRODUCT_PART_NUMBER,
    product_name: str = DEFAULT_PRODUCT_NAME,
) -> tuple[Any, Any, Any, Any, Any]:
    """
    功能: 创建校核 CATProduct 并装配前轮罩、后轮罩和车轮装配。
    输入: CATIA 和三个文件路径。
    输出: ProductDocument、根 Product 和三个组件。
    """
    product_document = catia.Documents.Add("Product")
    root_product = product_document.Product
    set_if_possible(root_product, "PartNumber", product_part_number)
    set_if_possible(root_product, "Name", product_name)

    front_component = add_component_from_file_to_product(
        product_document,
        root_product,
        front_wheelhouse_path,
    )
    rear_component = add_component_from_file_to_product(
        product_document,
        root_product,
        rear_wheelhouse_path,
    )
    wheel_assembly_component = add_component_from_file_to_product(
        product_document,
        root_product,
        wheel_assembly_path,
    )

    set_if_possible(front_component, "Name", FRONT_WHEELHOUSE_LABEL)
    set_if_possible(front_component, "PartNumber", "Front_Wheelhouse")
    set_if_possible(rear_component, "Name", REAR_WHEELHOUSE_LABEL)
    set_if_possible(rear_component, "PartNumber", "Rear_Wheelhouse")
    set_if_possible(wheel_assembly_component, "Name", WHEEL_ASSEMBLY_LABEL)
    set_if_possible(wheel_assembly_component, "PartNumber", "Wheel_Assembly")

    return (
        product_document,
        root_product,
        front_component,
        rear_component,
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


def iter_leaf_part_contexts(
    root_product: Any,
    root_path: str | None = None,
    product_chain: list[Any] | None = None,
) -> list[WheelPartContext]:
    """
    功能: 递归遍历装配树并收集叶子零件上下文。
    输入: 根 Product、路径和装配链。
    输出: WheelPartContext 列表。
    """
    contexts: list[WheelPartContext] = []
    current_name = product_display_name(root_product)
    current_path = root_path or current_name
    current_chain = [*(product_chain or []), root_product]
    children = get_child_products(root_product)

    if children:
        for index, child in enumerate(children, start=1):
            child_name = product_display_name(child)
            contexts.extend(
                iter_leaf_part_contexts(
                    child,
                    f"{current_path}/{index:03d}_{child_name}",
                    current_chain,
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
        "axis_direction_world": round_vector(candidate.axis_direction_world, 8),
        "axis_point_world": round_vector(candidate.axis_point_world),
        "component_cog_world": round_vector(candidate.component_cog_world),
        "bbox_world": bounding_box_to_dict(candidate.bbox_world),
        "mass": None if candidate.mass is None else round(candidate.mass, 6),
        "volume": None if candidate.volume is None else round(candidate.volume, 6),
        "tire_priority_score": round(tire_priority_score(candidate), 6),
        "warnings": candidate.warnings,
    }


def build_wheel_candidates(
    product_document: Any,
    wheel_part_contexts: list[WheelPartContext],
) -> tuple[list[WheelCandidate], list[dict[str, Any]], list[str]]:
    """
    功能: 构建有轴车轮候选并排除无轴零件。
    输入: ProductDocument 和叶子零件上下文列表。
    输出: 候选列表、无轴排除列表和告警列表。
    """
    candidates: list[WheelCandidate] = []
    excluded_no_axis_components: list[dict[str, Any]] = []
    warnings: list[str] = []

    for context in wheel_part_contexts:
        context_warnings: list[str] = []
        try:
            component_cog = evaluate_product_cog(
                product_document,
                context.product,
                context.component_path,
            )
        except Exception as exc:
            warnings.append(f"{context.component_path}: {exc}")
            continue

        bbox, bbox_warnings = evaluate_product_bounding_box(product_document, context.product)
        for warning in bbox_warnings:
            context_warnings.append(warning)

        mass = evaluate_product_scalar(context.product, "Mass")
        volume = evaluate_product_scalar(context.product, "Volume")
        world_transform = product_chain_world_transform(product_document, context.product_chain)
        axis_records = list(iter_axis_feature_records(context))

        if not axis_records:
            excluded_no_axis_components.append(
                {
                    "component_path": context.component_path,
                    "component_name": context.component_name,
                    "component_part_number": context.component_part_number,
                    "component_cog_world": round_vector(component_cog),
                    "bbox_world": bounding_box_to_dict(bbox),
                    "mass": None if mass is None else round(mass, 6),
                    "volume": None if volume is None else round(volume, 6),
                    "reason": "no_axis_feature",
                    "warnings": context_warnings.copy(),
                }
            )
            warnings.append(f"{context.component_path}: 未找到可测量轴线，已排除，不参与车轮位置聚类。")
            continue

        for feature_name, local_direction, local_axis_point in axis_records:
            try:
                world_direction = normalized_axis_direction(
                    apply_transform_to_direction(world_transform, local_direction)
                )
                world_axis_point = apply_transform_to_point(world_transform, local_axis_point)
            except Exception as exc:
                warnings.append(f"{context.component_path}/{feature_name}: 轴线坐标转换失败: {exc}")
                continue
            candidates.append(
                WheelCandidate(
                    context=context,
                    feature_name=feature_name,
                    axis_direction_world=world_direction,
                    axis_point_world=world_axis_point,
                    component_cog_world=component_cog,
                    bbox_world=bbox,
                    mass=mass,
                    volume=volume,
                    warnings=context_warnings.copy(),
                )
            )

    return candidates, excluded_no_axis_components, warnings


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


def match_wheelhouses_to_tire_representatives(
    product_document: Any,
    front_component: Any,
    rear_component: Any,
    tire_candidates: list[WheelCandidate],
) -> tuple[list[dict[str, Any]], list[dict[str, Any]], list[dict[str, Any]], list[str]]:
    """
    功能: 将前后轮罩匹配到 Tire 代表件。
    输入: ProductDocument、前后轮罩组件和 Tire 候选列表。
    输出: 匹配结果、轮罩重心、Tire 数据和告警。
    """
    warnings: list[str] = []
    front_context = component_context_from_product(front_component, FRONT_WHEELHOUSE_LABEL)
    rear_context = component_context_from_product(rear_component, REAR_WHEELHOUSE_LABEL)
    cover_items = [
        (FRONT_WHEELHOUSE_LABEL, evaluate_product_cog(product_document, front_context.product, FRONT_WHEELHOUSE_LABEL)),
        (REAR_WHEELHOUSE_LABEL, evaluate_product_cog(product_document, rear_context.product, REAR_WHEELHOUSE_LABEL)),
    ]
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
        warnings.append("Tire 代表件数量不足，允许前/后轮罩复用最近 Tire 候选。")
        for label, cover_cog in cover_items:
            nearest_tire = min(tire_candidates, key=lambda candidate: tire_match_score(cover_cog, candidate)[0])
            matches.append(build_tire_match_row(label, cover_cog, nearest_tire))

    return matches, cover_rows, tire_rows, warnings


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
    front_wheelhouse_part_path: str | Path,
    rear_wheelhouse_part_path: str | Path,
    wheel_assembly_path: str | Path,
    output_dir: str | Path | None = None,
    product_part_number: str = DEFAULT_PRODUCT_PART_NUMBER,
    product_name: str = DEFAULT_PRODUCT_NAME,
    axis_direction_tolerance_degrees: float = AXIS_DIRECTION_TOLERANCE_DEGREES,
    wheel_position_tolerance: float = WHEEL_POSITION_TOLERANCE,
    tire_hub_cog_tolerance: float = TIRE_HUB_COG_TOLERANCE,
    save_product: bool = True,
) -> dict[str, Any]:
    """
    功能: 执行完整车轮罩法规检测流程。
    输入: 前轮罩、后轮罩、车轮装配路径和配置。
    输出: 检测结果字典。
    """
    front_path = validate_existing_catia_file(front_wheelhouse_part_path, FRONT_WHEELHOUSE_LABEL)
    rear_path = validate_existing_catia_file(rear_wheelhouse_part_path, REAR_WHEELHOUSE_LABEL)
    wheel_path = validate_existing_catia_file(wheel_assembly_path, WHEEL_ASSEMBLY_LABEL)

    catia = start_or_connect_catia()
    (
        product_document,
        root_product,
        front_component,
        rear_component,
        wheel_assembly_component,
    ) = create_wheelhouse_product(
        catia,
        front_path,
        rear_path,
        wheel_path,
        product_part_number=product_part_number,
        product_name=product_name,
    )

    product_save_path = build_product_save_path(output_dir)
    wheel_part_contexts = iter_leaf_part_contexts(
        wheel_assembly_component,
        WHEEL_ASSEMBLY_LABEL,
    )
    if not wheel_part_contexts:
        raise RuntimeError("车轮装配中未找到可分析的零件。")

    warnings: list[str] = []
    wheel_candidates, excluded_no_axis_components, candidate_warnings = build_wheel_candidates(
        product_document,
        wheel_part_contexts,
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
        front_component,
        rear_component,
        tire_representatives,
    )
    warnings.extend(match_warnings)

    visible_component_paths = {
        match["tire_component_path"]
        for match in matches
        if match.get("tire_component_path")
    }
    visibility_rows = apply_wheel_part_visibility(
        product_document,
        wheel_part_contexts,
        visible_component_paths,
    )

    if save_product:
        product_document.SaveAs(str(product_save_path))
        saved_product_path: str | None = str(product_save_path)
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
            "front_wheelhouse": str(front_path),
            "rear_wheelhouse": str(rear_path),
            "wheel_assembly": str(wheel_path),
        },
        "wheel_part_count": len(wheel_part_contexts),
        "axis_direction_tolerance_degrees": axis_direction_tolerance_degrees,
        "wheel_position_tolerance": wheel_position_tolerance,
        "tire_hub_cog_tolerance": tire_hub_cog_tolerance,
        "excluded_no_axis_count": len(excluded_no_axis_components),
        "excluded_no_axis_components": excluded_no_axis_components,
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
        "wheelhouse_cogs": wheelhouse_cogs,
        "tire_cogs": tire_cogs,
        "matches": matches,
        "topology_target_tires": [
            match["tire_component_path"]
            for match in matches
            if match.get("tire_component_path")
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
    print(f"无轴排除零件数量: {result.get('excluded_no_axis_count')}")
    print(f"车轮候选数量: {result.get('wheel_candidate_count')}")
    print(f"车轮位置组数量: {result.get('wheel_position_group_count')}")
    print(f"Tire代表件数量: {result.get('tire_representative_count')}")
    print(f"筛除Hub/Rim候选数量: {len(result.get('excluded_hub_candidates', []))}")

    warnings = result.get("warnings") or []
    for warning in warnings:
        print(f"[警告] {warning}")

    print("\n-- Tire代表件 --")
    for index, tire in enumerate(result.get("tire_representatives", []), start=1):
        print(
            f"{index}. {tire.get('component_path')} / {tire.get('feature_name')} "
            f"重心={tire.get('component_cog_world')}, "
            f"包围盒对角线={((tire.get('bbox_world') or {}).get('diagonal'))}"
        )
    if not result.get("tire_representatives"):
        print("无")

    print("\n-- 无轴排除零件 --")
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

    print("\n-- 后续拓扑提取目标轮胎 --")
    for tire_path in result.get("topology_target_tires", []):
        print(tire_path)


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
        json.dumps(result, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


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
    )
    print_result_summary(result)
    write_json_result(json_result_path, result)
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
    parser.add_argument("front_wheelhouse_part_path", help="前轮罩 CATPart 路径")
    parser.add_argument("rear_wheelhouse_part_path", help="后轮罩 CATPart 路径")
    parser.add_argument("wheel_assembly_path", help="车轮装配 CATProduct/CATPart 路径")
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
        result = run_wheelhouse_regulation_verification(
            front_wheelhouse_part_path=args.front_wheelhouse_part_path,
            rear_wheelhouse_part_path=args.rear_wheelhouse_part_path,
            wheel_assembly_path=args.wheel_assembly_path,
            output_dir=args.output_dir,
            axis_direction_tolerance_degrees=args.axis_tolerance_deg,
            wheel_position_tolerance=args.wheel_position_tolerance,
            tire_hub_cog_tolerance=args.tire_hub_cog_tolerance,
            save_product=not args.no_save,
        )
        print_result_summary(result)
        write_json_result(args.json_result, result)
        return 0
    except Exception as exc:
        print(f"车轮罩法规校核失败: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(cli_main())
