from __future__ import annotations

import math
import os
import struct
import sys
import threading
import time
import zipfile
from datetime import datetime
from html import escape as escape_xml
from itertools import combinations, product
from pathlib import Path
from typing import Any, Iterable

import win32con
import win32com.client
import win32gui


# ==================== 变量命名区 ====================
# 修改为需要读取的 CATPart 文件路径。
READ_FILE_PATH = Path(__file__).resolve().with_name("Outside_Mirror_Regulation_Check.CATPart")

INPUT_PARAMETER_GEO_SET_NAME = "输入参数"
REGULATION_LINE_GEO_SET_NAME = "法规线"
PARAMETRIC_REARVIEW_MIRROR_GEO_SET_NAME = "参数化后视镜"
REGULATION_REFLECTION_POINT_GEO_SET_NAME = "法规反射取点"
GAP_CHECK_GEO_SET_NAME = "测量点间隙校验"

LEFT_MIRROR_FEATURE_NAME = "镜片 左"
RIGHT_MIRROR_FEATURE_NAME = "镜片 右"
LEFT_EYE_POINT_FEATURE_NAME = "左眼点"
RIGHT_EYE_POINT_FEATURE_NAME = "右眼点"
GROUND_FEATURE_NAME = "空载地面"
LEFT_VEHICLE_WIDTH_LINE_FEATURE_NAME = "车宽线-左"
RIGHT_VEHICLE_WIDTH_LINE_FEATURE_NAME = "车宽线-右"

SCREENSHOT_VIEW_DISTANCE = 644.0
REGULATION_VISION_ESTIMATED_MAX_DISTANCE = 24000.0
REGULATION_VISION_SCREENSHOT_MARGIN_FACTOR = 1.25
REGULATION_VISION_SCREENSHOT_VIEW_DISTANCE = (
    REGULATION_VISION_ESTIMATED_MAX_DISTANCE
    * REGULATION_VISION_SCREENSHOT_MARGIN_FACTOR
)
ANNOTATION_TEXT_SIZE = 5.0
ANNOTATION_COLOR = (0, 0, 0)
ANNOTATION_PART_NUMBER = "Rearview_Distance_Annotations"
AXIS_SNAP_ANGLE_DEGREES = 5.0
VEHICLE_WIDTH_LINE_EXTENSION_DISTANCE = 24000.0
REGULATION_IMAGE_PATH = Path(__file__).resolve().parent / "resources" / "Rearview_mirror_regulations.jpg"
# ================== 变量命名区结束 ==================


Vector = tuple[float, float, float]


def configure_console_encoding() -> None:
    """功能: 尽量让 Windows 终端正确输出中文；输入: 无；输出: 无。"""
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass


def normalized_path(path: str | Path) -> str:
    """功能: 生成用于路径比较的规范化小写路径；输入: 路径；输出: casefold 后的绝对路径字符串。"""
    return str(Path(path).expanduser().resolve()).casefold()


def get_output_dir() -> Path:
    """功能: 获取并确保算法 output 目录存在；输入: 无；输出: output 目录 Path。"""
    configured_output_dir = os.environ.get("REARVIEW_OUTPUT_DIR")
    output_dir = Path(configured_output_dir) if configured_output_dir else Path(__file__).resolve().parent / "output"
    output_dir.mkdir(parents=True, exist_ok=True)
    return output_dir


def build_timestamp() -> str:
    """功能: 生成文件命名时间戳；输入: 无；输出: yyyyMMdd_HHmmss 字符串。"""
    return datetime.now().strftime("%Y%m%d_%H%M%S")


def build_result_save_path(source_path: str | Path, output_dir: Path | None = None) -> Path:
    """功能: 构造结果 CATPart 另存路径；输入: 源文件路径和可选输出目录；输出: 结果文件 Path。"""
    source = Path(source_path).expanduser().resolve()
    target_dir = output_dir or get_output_dir()
    return target_dir / f"rearview_result_{build_timestamp()}{source.suffix}"


def build_screenshot_save_path(
    mirror_side_name: str,
    output_dir: Path | None = None,
    extension: str = ".png",
) -> Path:
    """功能: 构造后视镜截图路径；输入: 左右侧名称、可选输出目录和扩展名；输出: 截图文件 Path。"""
    target_dir = output_dir or get_output_dir()
    mirror_side_map = {"左": "left", "右": "right"}
    mirror_side = mirror_side_map.get(mirror_side_name, str(mirror_side_name).lower())
    return target_dir / f"{mirror_side}_mirror_screenshot_{build_timestamp()}{extension}"


def build_named_screenshot_save_path(
    screenshot_name: str,
    output_dir: Path | None = None,
    extension: str = ".png",
) -> Path:
    """功能: 构造指定名称截图路径；输入: 截图名称、可选输出目录和扩展名；输出: 截图文件 Path。"""
    target_dir = output_dir or get_output_dir()
    screenshot_name_map = {"法规视野截图": "regulation_vision_screenshot"}
    safe_name = screenshot_name_map.get(screenshot_name, "screenshot")
    return target_dir / f"{safe_name}_{build_timestamp()}{extension}"


def build_report_save_path(output_dir: Path | None = None) -> Path:
    """功能: 构造 Word 报告保存路径；输入: 可选输出目录；输出: 报告文件 Path。"""
    target_dir = output_dir or get_output_dir()
    return target_dir / f"rearview_inspection_report_{build_timestamp()}.docx"


def build_annotation_product_save_path(output_dir: Path | None = None) -> Path:
    """功能: 构造标注 CATProduct 保存路径；输入: 可选输出目录；输出: CATProduct 保存路径。"""
    target_dir = output_dir or get_output_dir()
    return target_dir / f"label_{build_timestamp()}.CATProduct"


def build_annotation_part_save_path(output_dir: Path | None = None) -> Path:
    """功能: 构造标注承载 CATPart 保存路径；输入: 可选输出目录；输出: CATPart 保存路径。"""
    target_dir = output_dir or get_output_dir()
    return target_dir / f"Rearview_Distance_Annotations_{build_timestamp()}.CATPart"


def get_image_size(image_path: str | Path) -> tuple[int, int] | None:
    """功能: 读取 PNG/JPEG/BMP 图片像素尺寸；输入: 图片路径；输出: (宽, 高) 或 None。"""
    path = Path(image_path)
    data = path.read_bytes()
    if data.startswith(b"\x89PNG\r\n\x1a\n") and len(data) >= 24:
        return struct.unpack(">II", data[16:24])
    if data.startswith(b"BM") and len(data) >= 26:
        return struct.unpack("<II", data[18:26])
    if data.startswith(b"\xff\xd8"):
        index = 2
        while index + 9 < len(data):
            if data[index] != 0xFF:
                index += 1
                continue
            marker = data[index + 1]
            index += 2
            if marker in (0xD8, 0xD9):
                continue
            if index + 2 > len(data):
                break
            segment_length = struct.unpack(">H", data[index : index + 2])[0]
            if marker in range(0xC0, 0xC4) and index + 7 < len(data):
                height, width = struct.unpack(">HH", data[index + 3 : index + 7])
                return width, height
            index += segment_length
    return None


def scaled_image_size(image_path: str | Path, target_width_px: int) -> tuple[int, int]:
    """功能: 按指定宽度等比缩放图片尺寸；输入: 图片路径和目标宽度；输出: 缩放后的宽高。"""
    original_size = get_image_size(image_path)
    if not original_size:
        return target_width_px, int(target_width_px * 0.5625)
    width, height = original_size
    if width <= 0 or height <= 0:
        return target_width_px, int(target_width_px * 0.5625)
    return target_width_px, int(round(target_width_px * height / width))


def scaled_image_ratio(image_path: str | Path, ratio: float) -> tuple[int, int]:
    """功能: 按比例等比缩放图片尺寸；输入: 图片路径和缩放比例；输出: 缩放后的宽高。"""
    original_size = get_image_size(image_path)
    if not original_size:
        return int(round(600 * ratio)), int(round(338 * ratio))
    width, height = original_size
    return int(round(width * ratio)), int(round(height * ratio))


def get_collection_item(collection: Any, name_or_index: str | int) -> Any:
    """功能: 从 CATIA 集合中读取对象；输入: 集合和名称/序号；输出: CATIA 对象。"""
    try:
        return collection.Item(name_or_index)
    except Exception as exc:
        raise LookupError(f"未找到对象: {name_or_index}") from exc


def iter_collection(collection: Any) -> Iterable[Any]:
    """功能: 遍历 CATIA 1 基集合；输入: CATIA 集合；输出: 元素迭代器。"""
    for index in range(1, int(collection.Count) + 1):
        yield collection.Item(index)


def start_or_connect_catia() -> Any:
    """功能: 连接已打开的 CATIA，未启动时自动启动；输入: 无；输出: CATIA 应用对象。"""
    try:
        catia = win32com.client.GetActiveObject("CATIA.Application")
    except Exception:
        catia = win32com.client.Dispatch("CATIA.Application")
    catia.Visible = True
    return catia


def get_catia_window_handle(catia: Any) -> int | None:
    """功能: 获取 CATIA 主窗口句柄；输入: CATIA 应用对象；输出: 窗口句柄或 None。"""
    for attribute_name in ("HWND", "Hwnd", "hWnd"):
        try:
            hwnd = int(getattr(catia, attribute_name))
            if hwnd > 0:
                return hwnd
        except Exception:
            pass
    return None


def is_owned_by_window(hwnd: int, owner_hwnd: int | None) -> bool:
    """功能: 判断窗口是否属于指定父/所有者窗口；输入: 窗口句柄和所有者句柄；输出: 是否属于。"""
    if not owner_hwnd:
        return False
    current = hwnd
    for _index in range(8):
        if current == owner_hwnd:
            return True
        try:
            parent = win32gui.GetParent(current)
            owner = win32gui.GetWindow(current, win32con.GW_OWNER)
        except Exception:
            break
        current = parent or owner
        if not current:
            break
    return False


def click_dialog_confirm_button(dialog_hwnd: int) -> bool:
    """功能: 点击对话框中的“确定/OK”按钮；输入: 对话框句柄；输出: 是否点击成功。"""
    confirm_texts = {"确定", "OK", "&OK"}
    button_handles: list[int] = []

    def collect_button(child_hwnd: int, _param: Any) -> bool:
        try:
            if win32gui.GetClassName(child_hwnd) == "Button":
                text = win32gui.GetWindowText(child_hwnd).strip()
                if text in confirm_texts:
                    button_handles.append(child_hwnd)
        except Exception:
            pass
        return True

    try:
        win32gui.EnumChildWindows(dialog_hwnd, collect_button, None)
    except Exception:
        return False

    if not button_handles:
        return False
    try:
        win32gui.PostMessage(button_handles[0], win32con.BM_CLICK, 0, 0)
        return True
    except Exception:
        return False


def click_catia_confirm_dialog_once(catia_hwnd: int | None = None) -> bool:
    """功能: 查找并确认 CATIA 弹窗；输入: 可选 CATIA 主窗口句柄；输出: 是否处理了弹窗。"""
    handled = False

    def inspect_window(hwnd: int, _param: Any) -> bool:
        nonlocal handled
        if handled:
            return False
        try:
            if not win32gui.IsWindowVisible(hwnd):
                return True
            if win32gui.GetClassName(hwnd) != "#32770":
                return True
            title = win32gui.GetWindowText(hwnd)
            if catia_hwnd and not is_owned_by_window(hwnd, catia_hwnd) and "CATIA" not in title.upper():
                return True
            handled = click_dialog_confirm_button(hwnd)
        except Exception:
            pass
        return not handled

    try:
        win32gui.EnumWindows(inspect_window, None)
    except Exception:
        return False
    return handled


def start_catia_confirm_dialog_watcher(
    catia: Any,
    timeout_seconds: float = 20.0,
    interval_seconds: float = 0.2,
) -> threading.Thread:
    """功能: 后台等待并点击 CATIA 确认弹窗；输入: CATIA、超时时间和轮询间隔；输出: 后台线程。"""
    catia_hwnd = get_catia_window_handle(catia)

    def worker() -> None:
        deadline = time.monotonic() + timeout_seconds
        while time.monotonic() < deadline:
            if click_catia_confirm_dialog_once(catia_hwnd):
                return
            time.sleep(interval_seconds)

    watcher = threading.Thread(target=worker, daemon=True)
    watcher.start()
    return watcher


def close_loaded_target_document(catia: Any, file_path: Path) -> None:
    """功能: 关闭已加载的同名或同路径目标文档；输入: CATIA 应用和目标路径；输出: 无。"""
    target_name = file_path.name.casefold()
    target_path = normalized_path(file_path)

    for document in list(iter_collection(catia.Documents)):
        document_name = str(document.Name).casefold()
        try:
            document_path = normalized_path(document.FullName)
        except Exception:
            document_path = ""

        if document_name == target_name or document_path == target_path:
            document.Close()


def open_target_document(catia: Any, file_path: str | Path) -> tuple[Any, Any]:
    """功能: 从磁盘重新打开指定 CATPart；输入: CATIA 应用和文件路径；输出: 文档对象和 Part 对象。"""
    target_path = Path(file_path).expanduser().resolve()
    if not target_path.is_file():
        raise FileNotFoundError(f"读取文件不存在: {target_path}")
    if target_path.suffix.casefold() != ".catpart":
        raise ValueError(f"读取文件不是 CATPart: {target_path}")

    close_loaded_target_document(catia, target_path)
    document = catia.Documents.Open(str(target_path))
    try:
        part = document.Part
    except Exception as exc:
        document.Close()
        raise TypeError(f"无法从文件中获取 Part 对象: {target_path}") from exc
    return document, part


def find_hybrid_body(part: Any, hybrid_body_name: str) -> Any:
    """功能: 按名称查找几何图形集；输入: Part 和几何图形集名称；输出: HybridBody。"""
    try:
        return get_collection_item(part.HybridBodies, hybrid_body_name)
    except LookupError as exc:
        raise LookupError(f"未找到几何图形集: {hybrid_body_name}") from exc


def find_feature(hybrid_body: Any, feature_name: str) -> Any:
    """功能: 在几何图形集中递归查找特征；输入: 几何图形集和特征名称；输出: CATIA 特征。"""
    collections: list[Any] = []
    for attribute_name in ("HybridShapes", "HybridBodies", "HybridSketches"):
        try:
            collections.append(getattr(hybrid_body, attribute_name))
        except Exception:
            pass

    for collection in collections:
        try:
            return collection.Item(feature_name)
        except Exception:
            pass

    for collection in collections:
        for item in iter_collection(collection):
            if str(item.Name).strip() == feature_name:
                return item
            try:
                return find_feature(item, feature_name)
            except LookupError:
                pass
    raise LookupError(f"在几何图形集“{hybrid_body.Name}”中未找到特征: {feature_name}")


def print_input_check(label: str, name: str, found: bool, detail: str | None = None) -> None:
    """功能: 输出前端可识别的输入检查日志；输入: 标签、名称、是否找到和详情；输出: 无。"""
    status = "success" if found else "error"
    result_text = "已找到" if found else "未找到"
    suffix = f"（{detail}）" if detail else ""
    print(f"[输入检查][{status}] {label}“{name}”: {result_text}{suffix}")


def checked_hybrid_body(part: Any, label: str, name: str) -> Any:
    """功能: 查找几何图形集并输出检查结果；输入: Part、显示标签和名称；输出: HybridBody。"""
    try:
        hybrid_body = find_hybrid_body(part, name)
        print_input_check(label, name, True)
        return hybrid_body
    except Exception as exc:
        print_input_check(label, name, False, str(exc))
        raise


def checked_feature(hybrid_body: Any, label: str, name: str) -> Any:
    """功能: 查找特征并输出检查结果；输入: 几何图形集、显示标签和名称；输出: CATIA 特征。"""
    try:
        feature = find_feature(hybrid_body, name)
        print_input_check(label, name, True)
        return feature
    except Exception as exc:
        print_input_check(label, name, False, str(exc))
        raise


def create_hybrid_body(part: Any, hybrid_body_name: str) -> Any:
    """功能: 创建几何图形集并设为当前工作对象；输入: Part 和几何图形集名称；输出: HybridBody。"""
    hybrid_body = part.HybridBodies.Add()
    hybrid_body.Name = hybrid_body_name
    part.InWorkObject = hybrid_body
    part.Update()
    return hybrid_body


def create_reference(part: Any, feature: Any) -> Any:
    """功能: 创建 CATIA 引用对象；输入: Part 和目标特征；输出: Reference。"""
    return part.CreateReferenceFromObject(feature)


def hide_object(document: Any, feature: Any) -> None:
    """功能: 通过 Selection 将对象设置为隐藏；输入: 文档对象和待隐藏特征；输出: 无。"""
    selection = document.Selection
    try:
        selection.Clear()
        selection.Add(feature)
        selection.VisProperties.SetShow(1)
    finally:
        selection.Clear()


def delete_object(document: Any, feature: Any) -> None:
    """功能: 通过 Selection 删除 CATIA 对象；输入: 文档对象和待删除特征；输出: 无。"""
    selection = document.Selection
    try:
        selection.Clear()
        selection.Add(feature)
        selection.Delete()
    finally:
        selection.Clear()


def hide_hybrid_body_features_except(
    document: Any,
    hybrid_body: Any,
    visible_feature_names: set[str],
) -> list[str]:
    """功能: 隐藏几何图形集中未指定保留的特征；输入: 文档、几何图形集和保留名称集合；输出: 已隐藏名称列表。"""
    hidden_names: list[str] = []
    collections: list[Any] = []
    for collection_name in ("HybridShapes", "HybridSketches"):
        try:
            collections.append(getattr(hybrid_body, collection_name))
        except Exception:
            pass

    for collection in collections:
        for feature in iter_collection(collection):
            feature_name = str(feature.Name)
            if feature_name in visible_feature_names:
                continue
            hide_object(document, feature)
            hidden_names.append(feature_name)
    return hidden_names


def set_object_color(
    document: Any,
    feature: Any,
    red: int,
    green: int,
    blue: int,
) -> None:
    """功能: 设置特征显示颜色；输入: 文档、特征和 RGB 值；输出: 无。"""
    set_object_visual_style(document, feature, red, green, blue)


def set_object_visual_style(
    document: Any,
    feature: Any,
    red: int,
    green: int,
    blue: int,
    width: int | None = None,
) -> None:
    """功能: 设置特征显示颜色和可选线宽；输入: 文档、特征、RGB 值和线宽；输出: 无。"""
    selection = document.Selection
    try:
        selection.Clear()
        selection.Add(feature)
        try:
            selection.VisProperties.SetShow(0)
        except Exception:
            pass
        selection.VisProperties.SetRealColor(red, green, blue, 1)
        if width is not None:
            try:
                selection.VisProperties.SetRealWidth(width, 1)
            except Exception:
                pass
    finally:
        selection.Clear()


def append_hybrid_shape(part: Any, hybrid_body: Any, hybrid_shape: Any, name: str) -> Any:
    """功能: 将混合形状加入几何图形集并命名更新；输入: Part、几何图形集、形状和名称；输出: 形状对象。"""
    hybrid_shape.Name = name
    hybrid_body.AppendHybridShape(hybrid_shape)
    part.InWorkObject = hybrid_shape
    part.Update()
    return hybrid_shape


def create_or_update_angle_parameter(part: Any, name: str, value_text: str = "0deg") -> Any:
    """功能: 创建或更新角度参数；输入: Part、参数名和角度文本；输出: 参数对象。"""
    parameters = part.Parameters
    try:
        parameter = parameters.Item(name)
    except Exception:
        parameter = parameters.CreateDimension(name, "ANGLE", 0.0)

    try:
        parameter.ValuateFromString(value_text)
    except Exception:
        parameter.Value = 0.0
    return parameter


def get_parameter_value_text(parameter: Any) -> str:
    """功能: 安全读取 CATIA 参数值文本；输入: 参数对象；输出: 字符串值。"""
    try:
        return str(parameter.ValueAsString())
    except Exception:
        try:
            return str(parameter.Value)
        except Exception:
            return ""


def get_rotation_angle_parameter(part: Any, rotation_feature: Any) -> Any:
    """功能: 读取旋转特征内部可被公式驱动的角度参数；输入: Part 和旋转特征；输出: 角度参数。"""
    try:
        return rotation_feature.Angle
    except Exception:
        pass

    try:
        parameters = part.Parameters.SubList(rotation_feature, True)
        for parameter in iter_collection(parameters):
            parameter_name = str(parameter.Name)
            if "Angle" in parameter_name or "角度" in parameter_name:
                return parameter
    except Exception:
        pass

    raise RuntimeError(f"无法读取旋转特征“{rotation_feature.Name}”的角度参数。")


def bind_parameter_with_formula(
    part: Any,
    target_parameter: Any,
    source_parameter: Any,
    formula_name: str,
) -> tuple[str | None, str | None, str | None]:
    """功能: 用公式将目标参数绑定到源参数；输入: Part、目标参数、源参数和公式名；输出: 公式名、警告和公式表达式。"""
    relation_names: list[str] = []
    try:
        relation_names.append(str(part.Parameters.GetNameToUseInRelation(source_parameter)))
    except Exception:
        pass
    relation_names.extend(
        [
            f"`{source_parameter.Name}`",
            str(source_parameter.Name),
        ]
    )
    formula_bodies = []
    for relation_name in relation_names:
        if relation_name and relation_name not in formula_bodies:
            formula_bodies.append(relation_name)

    last_error = ""
    for formula_body in formula_bodies:
        try:
            existing_relation = target_parameter.OptionalRelation
            if existing_relation:
                existing_relation.Modify(formula_body)
                try:
                    existing_relation.Activated = True
                except Exception:
                    pass
                part.Update()
                return str(existing_relation.Name), None, formula_body
        except Exception as exc:
            last_error = str(exc)

        try:
            formula = part.Relations.CreateFormula(
                formula_name,
                "",
                target_parameter,
                formula_body,
            )
            try:
                formula.Activated = True
            except Exception:
                pass
            part.Update()
            return str(formula.Name), None, formula_body
        except Exception as exc:
            last_error = str(exc)
    return None, f"{formula_name}创建失败: {last_error}", None


def create_rotation_with_angle_parameter(
    document: Any,
    part: Any,
    hybrid_body: Any,
    element: Any,
    axis: Any,
    angle_parameter_name: str,
    name: str,
    hide_initial_element: bool = True,
) -> dict[str, Any]:
    """功能: 创建旋转特征并绑定角度参数；输入: 文档、Part、几何图形集、源元素、轴、角度参数名、特征名和隐藏选项；输出: 旋转结果字典。"""
    angle_parameter = part.Parameters.Item(angle_parameter_name)
    rotation = part.HybridShapeFactory.AddNewRotate(
        create_reference(part, element),
        create_reference(part, axis),
        0.0,
    )
    append_hybrid_shape(part, hybrid_body, rotation, name)

    formula_name: str | None = None
    formula_warning: str | None = None
    formula_body: str | None = None
    rotation_angle_parameter_name: str | None = None
    try:
        rotation.RotationType = 0
    except Exception:
        pass
    try:
        rotation.AngleValue = 0.0
    except Exception:
        pass
    try:
        rotation_angle_parameter = get_rotation_angle_parameter(part, rotation)
        rotation_angle_parameter_name = str(rotation_angle_parameter.Name)
        formula_name, formula_warning, formula_body = bind_parameter_with_formula(
            part,
            rotation_angle_parameter,
            angle_parameter,
            f"{name}角度公式",
        )
    except Exception as exc:
        formula_warning = str(exc)

    if hide_initial_element:
        hide_object(document, element)

    return {
        "feature": rotation,
        "name": str(rotation.Name),
        "axis_name": str(axis.Name),
        "source_element_name": str(element.Name),
        "angle_parameter_name": str(angle_parameter.Name),
        "rotation_angle_parameter_name": rotation_angle_parameter_name,
        "formula_name": formula_name,
        "formula_body": formula_body,
        "formula_warning": formula_warning,
        "initial_element_hidden": hide_initial_element,
    }


def create_parametric_rearview_mirror(
    document: Any,
    part: Any,
    parametric_hybrid_body_name: str,
    left_mirror: Any,
    right_mirror: Any,
    up_direction: Vector,
) -> dict[str, Any]:
    """功能: 创建参数化后视镜几何集、角度参数和左右镜片旋转参考；输入: 文档、Part、几何集名、左右镜片和上方向；输出: 构造结果字典。"""
    parametric_hybrid_body = create_hybrid_body(part, parametric_hybrid_body_name)
    parameter_names = (
        "左镜片水平旋转角度",
        "左镜片竖直旋转角度",
        "右镜片水平旋转角度",
        "右镜片竖直旋转角度",
    )
    parameters = []

    for parameter_name in parameter_names:
        parameter = create_or_update_angle_parameter(part, parameter_name, "0deg")
        parameters.append(
            {
                "name": str(parameter.Name),
                "type": "ANGLE",
                "value": get_parameter_value_text(parameter),
            }
        )

    def create_one_side_mirror_rotation(side_name: str, mirror: Any) -> dict[str, Any]:
        """功能: 为单侧镜片创建旋转中心、旋转轴和旋转特征；输入: 侧别和镜片特征；输出: 构造结果字典。"""
        centroid_coordinates = get_surface_center(document, part, mirror)
        centroid = create_coordinate_point(
            part,
            parametric_hybrid_body,
            centroid_coordinates,
            f"{side_name}镜片重心点",
        )
        sphere_center = create_point_center(
            part,
            parametric_hybrid_body,
            mirror,
            f"{side_name}镜片旋转参考球心",
        )
        hide_object(document, sphere_center)
        sphere_center_coordinates = get_point(document, part, sphere_center)
        mirror_normal_direction = normalize(
            subtract(centroid_coordinates, sphere_center_coordinates),
            f"{side_name}镜片法线方向",
        )
        mirror_normal = create_line_by_center_direction(
            document,
            part,
            parametric_hybrid_body,
            centroid_coordinates,
            mirror_normal_direction,
            -100.0,
            100.0,
            f"{side_name}镜片法线",
        )
        normal_intersection = create_intersection(
            part,
            parametric_hybrid_body,
            mirror_normal,
            mirror,
            f"{side_name}镜片法线与镜片交点",
        )
        normal_intersection_coordinates = get_point(document, part, normal_intersection)
        rotation_center_direction = normalize(
            subtract(sphere_center_coordinates, normal_intersection_coordinates),
            f"{side_name}镜片旋转中心偏移方向",
        )
        rotation_center_coordinates = add(
            normal_intersection_coordinates,
            scale(rotation_center_direction, 9.5),
        )
        rotation_center = create_coordinate_point(
            part,
            parametric_hybrid_body,
            rotation_center_coordinates,
            f"{side_name}镜片旋转中心",
        )
        rotation_reference_plane, horizontal_axis_direction, vertical_axis_direction = (
            create_plane_by_point_normal(
                document,
                part,
                parametric_hybrid_body,
                rotation_center_coordinates,
                mirror_normal_direction,
                up_direction,
                f"{side_name}镜片旋转参考平面",
            )
        )
        rotation_axis_sketch = create_sketch_on_plane(
            part,
            parametric_hybrid_body,
            rotation_reference_plane,
            f"{side_name}镜片旋转轴草图",
        )
        sketch_axis_line_names: list[str] = []
        axis_definitions = (
            (f"{side_name}镜片水平旋转轴", horizontal_axis_direction),
            (f"{side_name}镜片竖直旋转轴", vertical_axis_direction),
        )
        try:
            factory2d = rotation_axis_sketch.OpenEdition()
            plane_origin, plane_x, plane_y = get_plane(document, part, rotation_reference_plane)
            for axis_name, axis_direction in axis_definitions:
                start_xy = project_point_to_plane_2d(
                    add(rotation_center_coordinates, scale(axis_direction, -100.0)),
                    plane_origin,
                    plane_x,
                    plane_y,
                )
                end_xy = project_point_to_plane_2d(
                    add(rotation_center_coordinates, scale(axis_direction, 100.0)),
                    plane_origin,
                    plane_x,
                    plane_y,
                )
                axis_line = create_sketch_line_by_xy(
                    factory2d,
                    start_xy,
                    end_xy,
                    f"草图内{axis_name}",
                )
                sketch_axis_line_names.append(str(axis_line.Name))
            rotation_axis_sketch.Evaluate()
        finally:
            rotation_axis_sketch.CloseEdition()
            part.Update()
        hide_object(document, rotation_axis_sketch)

        axis_3d_lines = [
            create_line_by_center_direction(
                document,
                part,
                parametric_hybrid_body,
                rotation_center_coordinates,
                axis_direction,
                -100.0,
                100.0,
                axis_name,
            )
            for axis_name, axis_direction in axis_definitions
        ]
        axis_3d_line_by_name = {str(axis_line.Name): axis_line for axis_line in axis_3d_lines}

        horizontal_rotation = create_rotation_with_angle_parameter(
            document=document,
            part=part,
            hybrid_body=parametric_hybrid_body,
            element=mirror,
            axis=axis_3d_line_by_name[f"{side_name}镜片水平旋转轴"],
            angle_parameter_name=f"{side_name}镜片水平旋转角度",
            name=f"{side_name}镜片仅水平旋转",
            hide_initial_element=True,
        )
        final_rotation = create_rotation_with_angle_parameter(
            document=document,
            part=part,
            hybrid_body=parametric_hybrid_body,
            element=horizontal_rotation["feature"],
            axis=axis_3d_line_by_name[f"{side_name}镜片竖直旋转轴"],
            angle_parameter_name=f"{side_name}镜片竖直旋转角度",
            name=f"{side_name}镜片旋转",
            hide_initial_element=True,
        )
        rotation_results = [horizontal_rotation, final_rotation]
        for rotation_result in rotation_results:
            rotation_result.pop("feature", None)

        return {
            "centroid_name": str(centroid.Name),
            "centroid_coordinates": round_vector(centroid_coordinates),
            "normal_name": str(mirror_normal.Name),
            "normal_intersection_name": str(normal_intersection.Name),
            "normal_intersection_coordinates": round_vector(normal_intersection_coordinates),
            "rotation_center_name": str(rotation_center.Name),
            "rotation_center_coordinates": round_vector(rotation_center_coordinates),
            "rotation_reference_plane_name": str(rotation_reference_plane.Name),
            "rotation_axis_sketch_name": str(rotation_axis_sketch.Name),
            "rotation_axis_sketch_hidden": True,
            "rotation_axis_sketch_line_names": sketch_axis_line_names,
            "rotation_axis_line_names": [str(axis_line.Name) for axis_line in axis_3d_lines],
            "rotation_feature_names": [
                horizontal_rotation["name"],
                final_rotation["name"],
            ],
            "rotation_features": rotation_results,
            "horizontal_axis_direction": round_vector(horizontal_axis_direction),
            "vertical_axis_direction": round_vector(vertical_axis_direction),
        }

    mirror_results = {
        "左": create_one_side_mirror_rotation("左", left_mirror),
        "右": create_one_side_mirror_rotation("右", right_mirror),
    }
    left_result = mirror_results["左"]
    right_result = mirror_results["右"]
    visible_rotation_names = {
        left_result["rotation_feature_names"][-1],
        right_result["rotation_feature_names"][-1],
    }
    hidden_process_feature_names = hide_hybrid_body_features_except(
        document,
        parametric_hybrid_body,
        visible_rotation_names,
    )

    part.InWorkObject = parametric_hybrid_body
    part.Update()
    return {
        "geo_set_name": str(parametric_hybrid_body.Name),
        "parameters": parameters,
        "mirrors": mirror_results,
        "left_mirror_centroid_name": left_result["centroid_name"],
        "left_mirror_centroid_coordinates": left_result["centroid_coordinates"],
        "left_mirror_normal_name": left_result["normal_name"],
        "left_mirror_normal_intersection_name": left_result["normal_intersection_name"],
        "left_mirror_normal_intersection_coordinates": left_result[
            "normal_intersection_coordinates"
        ],
        "left_mirror_rotation_center_name": left_result["rotation_center_name"],
        "left_mirror_rotation_center_coordinates": left_result[
            "rotation_center_coordinates"
        ],
        "right_mirror_centroid_name": right_result["centroid_name"],
        "right_mirror_centroid_coordinates": right_result["centroid_coordinates"],
        "right_mirror_normal_name": right_result["normal_name"],
        "right_mirror_normal_intersection_name": right_result["normal_intersection_name"],
        "right_mirror_normal_intersection_coordinates": right_result[
            "normal_intersection_coordinates"
        ],
        "right_mirror_rotation_center_name": right_result["rotation_center_name"],
        "right_mirror_rotation_center_coordinates": right_result[
            "rotation_center_coordinates"
        ],
        "rotation_reference_plane_name": left_result["rotation_reference_plane_name"],
        "rotation_axis_sketch_name": left_result["rotation_axis_sketch_name"],
        "rotation_axis_sketch_hidden": left_result["rotation_axis_sketch_hidden"],
        "rotation_axis_sketch_line_names": (
            left_result["rotation_axis_sketch_line_names"]
            + right_result["rotation_axis_sketch_line_names"]
        ),
        "rotation_axis_line_names": (
            left_result["rotation_axis_line_names"]
            + right_result["rotation_axis_line_names"]
        ),
        "rotation_feature_names": (
            left_result["rotation_feature_names"]
            + right_result["rotation_feature_names"]
        ),
        "rotation_features": (
            left_result["rotation_features"]
            + right_result["rotation_features"]
        ),
        "visible_feature_names": sorted(visible_rotation_names),
        "hidden_process_feature_names": hidden_process_feature_names,
        "horizontal_axis_direction": left_result["horizontal_axis_direction"],
        "vertical_axis_direction": left_result["vertical_axis_direction"],
    }


def create_projection(
    part: Any,
    hybrid_body: Any,
    source_feature: Any,
    support_feature: Any,
    name: str,
) -> Any:
    """功能: 创建点到平面或直线的法向投影；输入: Part、几何图形集、源特征、支撑特征和名称；输出: 投影特征。"""
    factory = part.HybridShapeFactory
    projection = factory.AddNewProject(
        create_reference(part, source_feature),
        create_reference(part, support_feature),
    )
    projection.SolutionType = 0
    projection.Normal = True
    projection.SmoothingType = 0
    projection.ExtrapolationMode = 0
    return append_hybrid_shape(part, hybrid_body, projection, name)


def create_intersection(
    part: Any,
    hybrid_body: Any,
    first_feature: Any,
    second_feature: Any,
    name: str,
) -> Any:
    """功能: 创建两个 3D 元素的相交结果；输入: Part、几何图形集、两个特征和名称；输出: 相交特征。"""
    intersection = part.HybridShapeFactory.AddNewIntersection(
        create_reference(part, first_feature),
        create_reference(part, second_feature),
    )
    return append_hybrid_shape(part, hybrid_body, intersection, name)


def create_boundary(
    part: Any,
    hybrid_body: Any,
    surface_feature: Any,
    name: str,
) -> Any:
    """功能: 提取曲面边界；输入: Part、几何图形集、曲面特征和名称；输出: 边界特征。"""
    factory = part.HybridShapeFactory
    surface_reference = create_reference(part, surface_feature)
    try:
        boundary = factory.AddNewBoundaryOfSurface(surface_reference)
    except Exception:
        boundary = factory.AddNewBoundary(surface_reference, 0)
    return append_hybrid_shape(part, hybrid_body, boundary, name)


def create_coordinate_point(
    part: Any,
    hybrid_body: Any,
    coordinates: Vector,
    name: str,
) -> Any:
    """功能: 按绝对坐标创建点；输入: Part、几何图形集、三维坐标和名称；输出: 点特征。"""
    point = part.HybridShapeFactory.AddNewPointCoord(*coordinates)
    return append_hybrid_shape(part, hybrid_body, point, name)


def create_point_on_curve_with_reference_distance(
    part: Any,
    hybrid_body: Any,
    curve_feature: Any,
    reference_point: Any,
    distance_value: float,
    orientation: bool,
    name: str,
) -> Any:
    """功能: 在曲线上按参考点和距离创建点；输入: Part、几何集、曲线、参考点、距离、方向和名称；输出: 点特征。"""
    factory = part.HybridShapeFactory
    curve_reference = create_reference(part, curve_feature)
    point_reference = create_reference(part, reference_point)
    attempts = (
        lambda: factory.AddNewPointOnCurveWithReferenceFromDistance(
            curve_reference,
            point_reference,
            float(distance_value),
            bool(orientation),
        ),
        lambda: factory.AddNewPointOnCurveWithReferenceFromPercent(
            curve_reference,
            point_reference,
            float(distance_value),
            bool(orientation),
        ),
    )
    last_error: Exception | None = None
    for attempt in attempts:
        try:
            return append_hybrid_shape(part, hybrid_body, attempt(), name)
        except Exception as exc:
            last_error = exc
    raise RuntimeError(f"无法在曲线上创建点“{name}”: {last_error}") from last_error


def create_extremum_point(
    part: Any,
    hybrid_body: Any,
    feature: Any,
    direction: Vector,
    name: str,
) -> Any:
    """功能: 创建特征沿指定方向的极值点；输入: Part、几何图形集、特征、方向和名称；输出: 极值点特征。"""
    factory = part.HybridShapeFactory
    extremum_direction = factory.AddNewDirectionByCoord(*direction)
    extremum = factory.AddNewExtremum(
        create_reference(part, feature),
        extremum_direction,
        1,
    )
    return append_hybrid_shape(part, hybrid_body, extremum, name)


def create_point_center(part: Any, hybrid_body: Any, surface_feature: Any, name: str) -> Any:
    """功能: 创建曲面中心点；输入: Part、几何图形集、曲面特征和名称；输出: 中心点特征。"""
    point = part.HybridShapeFactory.AddNewPointCenter(create_reference(part, surface_feature))
    return append_hybrid_shape(part, hybrid_body, point, name)


def create_sphere_surface(
    document: Any,
    part: Any,
    hybrid_body: Any,
    center_point: Any,
    radius: float,
    name: str,
) -> Any:
    """功能: 按球心和半径创建球面，轴线使用 CATIA 默认值；输入: 文档、Part、几何图形集、球心点、半径和名称；输出: 球面特征。"""
    factory = part.HybridShapeFactory
    center_reference = create_reference(part, center_point)
    vba_nothing = """
Function N()
    Set N = Nothing
End Function
"""
    default_axis = document.Application.SystemService.Evaluate(
        vba_nothing,
        0,
        "N",
        [],
    )
    try:
        sphere = factory.AddNewSphere(
            center_reference,
            default_axis,
            float(radius),
            -90.0,
            90.0,
            0.0,
            360.0,
        )
        return append_hybrid_shape(part, hybrid_body, sphere, name)
    except Exception as exc:
        raise RuntimeError(f"无法使用默认轴线创建球面“{name}”: {exc}") from exc


def create_plane_3_points(
    part: Any,
    hybrid_body: Any,
    first_point: Any,
    second_point: Any,
    third_point: Any,
    name: str,
) -> Any:
    """功能: 通过三个点创建平面；输入: Part、几何图形集、三个点和名称；输出: 平面特征。"""
    plane = part.HybridShapeFactory.AddNewPlane3Points(
        create_reference(part, first_point),
        create_reference(part, second_point),
        create_reference(part, third_point),
    )
    return append_hybrid_shape(part, hybrid_body, plane, name)


def create_point_to_point_line(
    part: Any,
    hybrid_body: Any,
    start_point: Any,
    end_point: Any,
    name: str,
) -> Any:
    """功能: 连接两个点创建直线；输入: Part、几何图形集、起点、终点和名称；输出: 直线特征。"""
    line = part.HybridShapeFactory.AddNewLinePtPt(
        create_reference(part, start_point),
        create_reference(part, end_point),
    )
    return append_hybrid_shape(part, hybrid_body, line, name)


def create_line_by_center_direction(
    document: Any,
    part: Any,
    hybrid_body: Any,
    center: Vector,
    direction: Vector,
    start_offset: float,
    end_offset: float,
    name: str,
) -> Any:
    """功能: 按中心、方向和两端偏移创建直线并隐藏辅助端点；输入: 文档、Part、几何图形集、中心、方向、两端偏移和名称；输出: 直线特征。"""
    unit_direction = normalize(direction, f"{name}方向")
    start_point = create_coordinate_point(
        part,
        hybrid_body,
        add(center, scale(unit_direction, start_offset)),
        f"{name}起点",
    )
    end_point = create_coordinate_point(
        part,
        hybrid_body,
        add(center, scale(unit_direction, end_offset)),
        f"{name}终点",
    )
    hide_object(document, start_point)
    hide_object(document, end_point)
    return create_point_to_point_line(part, hybrid_body, start_point, end_point, name)


def create_plane_by_point_normal(
    document: Any,
    part: Any,
    hybrid_body: Any,
    origin: Vector,
    normal_direction: Vector,
    preferred_vertical_direction: Vector,
    name: str,
) -> tuple[Any, Vector, Vector]:
    """功能: 通过点和法线创建平面并计算平面内方向；输入: 文档、Part、几何图形集、原点、法线、优先竖直方向和名称；输出: 平面、水平向量、竖直向量。"""
    normal = normalize(normal_direction, f"{name}法线方向")
    vertical = project_to_plane(preferred_vertical_direction, normal)
    if math.sqrt(dot(vertical, vertical)) <= 1e-6:
        vertical = project_to_plane((0.0, 1.0, 0.0), normal)
    vertical = normalize(vertical, f"{name}竖直方向")
    horizontal = normalize(cross(vertical, normal), f"{name}水平方向")

    horizontal_point = create_coordinate_point(
        part,
        hybrid_body,
        add(origin, scale(horizontal, 100.0)),
        f"{name}水平参考点",
    )
    vertical_point = create_coordinate_point(
        part,
        hybrid_body,
        add(origin, scale(vertical, 100.0)),
        f"{name}竖直参考点",
    )
    origin_point = create_coordinate_point(
        part,
        hybrid_body,
        origin,
        f"{name}中心参考点",
    )
    for helper_point in (horizontal_point, vertical_point, origin_point):
        hide_object(document, helper_point)

    plane = create_plane_3_points(
        part,
        hybrid_body,
        origin_point,
        horizontal_point,
        vertical_point,
        name,
    )
    return plane, horizontal, vertical


def get_measurable(document: Any, part: Any, feature: Any) -> Any:
    """功能: 获取 CATIA SPA 可测量对象；输入: 文档、Part、目标特征；输出: Measurable。"""
    spa_workbench = document.GetWorkbench("SPAWorkbench")
    return spa_workbench.GetMeasurable(create_reference(part, feature))


def evaluate_measurable_array(
    document: Any,
    measurable: Any,
    method_name: str,
    item_count: int,
    label: str = "",
) -> tuple[float, ...]:
    """
    功能: 通过 CATIA VBA 执行带输出数组的测量方法；输入: 文档、可测对象、方法名、数组长度和说明；输出: 浮点数组。

    GetPoint、GetCOG 和 GetPlane 的 COM 参数是 ByRef 数组。win32com 无法直接
    获取返回值，因此使用 CATIA SystemService.Evaluate 代为调用。
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
        raise RuntimeError(f"CATIA 测量方法 {method_name} 执行失败{target}。") from exc


def evaluate_object_array(
    document: Any,
    target: Any,
    method_name: str,
    item_count: int,
    label: str,
) -> tuple[float, ...]:
    """功能: 通过 CATIA VBA 读取普通 COM 对象的 ByRef 数组返回值；输入: 文档、目标对象、方法名、数组长度和说明；输出: 浮点数组。"""
    function_name = f"read_object_{method_name.casefold()}"
    vba_code = f"""
Public Function {function_name}(target)
    Dim values({item_count - 1})
    target.{method_name} values
    {function_name} = values
End Function
"""
    try:
        values = document.Application.SystemService.Evaluate(
            vba_code,
            0,
            function_name,
            [target],
        )
        return tuple(float(value) for value in values)
    except Exception as exc:
        raise RuntimeError(f"无法读取{label}。") from exc


def as_vector(values: Any, label: str) -> Vector:
    """功能: 将数值序列转换为三维向量；输入: 坐标序列和标签；输出: Vector。"""
    try:
        items = tuple(float(value) for value in values)
    except Exception as exc:
        raise TypeError(f"无法读取{label}坐标。") from exc
    if len(items) < 3:
        raise ValueError(f"{label}坐标数量不足: {items}")
    return items[0], items[1], items[2]


def get_point(document: Any, part: Any, feature: Any) -> Vector:
    """功能: 读取点或可测对象中心坐标；输入: 文档、Part、特征；输出: 三维坐标。"""
    measurable = get_measurable(document, part, feature)
    label = str(feature.Name)
    try:
        values = evaluate_measurable_array(document, measurable, "GetPoint", 3, label)
    except RuntimeError as get_point_error:
        try:
            values = evaluate_measurable_array(document, measurable, "GetCOG", 3, label)
        except RuntimeError:
            raise get_point_error
    return as_vector(values, f"点“{label}”")


def get_surface_center(document: Any, part: Any, feature: Any) -> Vector:
    """功能: 读取曲面重心/中心坐标；输入: 文档、Part、曲面特征；输出: 三维坐标。"""
    measurable = get_measurable(document, part, feature)
    values = evaluate_measurable_array(document, measurable, "GetCOG", 3, str(feature.Name))
    return as_vector(values, f"曲面“{feature.Name}”中心")


def get_plane(document: Any, part: Any, feature: Any) -> tuple[Vector, Vector, Vector]:
    """功能: 读取平面的原点和两个平面内方向；输入: 文档、Part 和平面特征；输出: 原点、X方向、Y方向。"""
    measurable = get_measurable(document, part, feature)
    try:
        values = evaluate_measurable_array(document, measurable, "GetPlane", 9, str(feature.Name))
    except Exception as exc:
        raise TypeError(f"特征“{feature.Name}”不是可测量平面。") from exc
    if len(values) < 9:
        raise ValueError(f"平面“{feature.Name}”返回的数据不足: {values}")
    return values[0:3], values[3:6], values[6:9]


def get_minimum_distance(document: Any, part: Any, first_feature: Any, second_feature: Any) -> float:
    """功能: 使用 CATIA 原生测量获取两个特征最小距离；输入: 文档、Part 和两个特征；输出: 距离值。"""
    measurable = get_measurable(document, part, first_feature)
    try:
        return float(measurable.GetMinimumDistance(create_reference(part, second_feature)))
    except Exception as exc:
        raise RuntimeError(
            f"无法测量“{first_feature.Name}”到“{second_feature.Name}”的最小距离。"
        ) from exc


def get_minimum_distance_points(
    document: Any,
    part: Any,
    first_feature: Any,
    second_feature: Any,
) -> tuple[Vector, Vector, tuple[float, ...]]:
    """功能: 读取最小距离对应的两端测量点坐标；输入: 文档、Part 和两个特征；输出: 两端坐标和原始坐标元组。"""
    vba_code = """
Public Function read_minimum_distance_points(measurable, measured_item)
    Dim values(8)
    measurable.GetMinimumDistancePoints measured_item, values
    read_minimum_distance_points = values
End Function
"""

    def read_coordinates(source_feature: Any, measured_feature: Any) -> tuple[float, ...]:
        """功能: 调用 CATIA 获取两个特征最小距离端点；输入: 源特征和被测特征；输出: 原始坐标元组。"""
        measurable = get_measurable(document, part, source_feature)
        measured_reference = create_reference(part, measured_feature)
        values = document.Application.SystemService.Evaluate(
            vba_code,
            0,
            "read_minimum_distance_points",
            [measurable, measured_reference],
        )
        coordinates: list[float] = []
        for value in values:
            if value is None:
                continue
            coordinates.append(float(value))
        if len(coordinates) < 6:
            raise ValueError(f"有效坐标数量不足，原始返回值: {values}")
        return tuple(coordinates)

    diagnostics: list[str] = []
    try:
        coordinates = read_coordinates(first_feature, second_feature)
        return coordinates[0:3], coordinates[3:6], coordinates
    except Exception as exc:
        diagnostics.append(f"正向失败: {exc}")

    try:
        coordinates = read_coordinates(second_feature, first_feature)
        # 反向读取时，前 3 个坐标属于 second_feature，后 3 个坐标属于 first_feature。
        return coordinates[3:6], coordinates[0:3], coordinates
    except Exception as exc:
        diagnostics.append(f"反向失败: {exc}")

    detail = "；".join(diagnostics)
    raise RuntimeError(
        f"无法读取“{first_feature.Name}”到“{second_feature.Name}”的最小距离测量点。{detail}"
    )


def try_call_com_methods(target: Any, method_names: tuple[str, ...], *args: Any) -> None:
    """功能: 依次尝试多个 COM 方法名；输入: 目标对象、方法名和参数；输出: 无，失败则抛异常。"""
    last_error: Exception | None = None
    for method_name in method_names:
        try:
            getattr(target, method_name)(*args)
            return
        except Exception as exc:
            last_error = exc
    if last_error is not None:
        raise last_error


def configure_active_view(
    document: Any,
    target_point: Vector,
    sight_direction: Vector,
    up_direction: Vector,
    view_distance: float,
) -> Any:
    """功能: 设置 CATIA 当前 3D 视图；输入: 文档、目标点、视角方向、上方向和视距；输出: ActiveViewer。"""
    viewer = document.Application.ActiveWindow.ActiveViewer
    viewpoint = viewer.Viewpoint3D
    target_point = tuple(float(value) for value in target_point)
    sight_direction = normalize(sight_direction, "截图视角方向")
    up_direction = normalize(project_to_plane(up_direction, sight_direction), "截图上方向")

    try_call_com_methods(viewpoint, ("PutOrigin", "SetOrigin"), target_point)
    try_call_com_methods(viewpoint, ("PutSightDirection", "SetSightDirection"), sight_direction)
    try_call_com_methods(viewpoint, ("PutUpDirection", "SetUpDirection"), up_direction)
    try:
        viewpoint.FocusDistance = float(view_distance)
    except Exception:
        pass
    try:
        viewer.Update()
    except Exception:
        pass
    return viewer


def capture_viewer_to_file(viewer: Any, image_path: Path, document: Any | None = None) -> Path:
    """功能: 保存 CATIA 当前视图截图；输入: ActiveViewer、目标路径和可选文档；输出: 实际保存路径。"""
    target_path = image_path.with_suffix(".png")
    target_path.parent.mkdir(parents=True, exist_ok=True)

    if document is not None:
        try:
            document.Selection.Clear()
        except Exception:
            pass

    try:
        viewer.Update()
    except Exception:
        pass
    time.sleep(1)

    try:
        viewer.CaptureToFile(4, str(target_path))
    except Exception as exc:
        raise RuntimeError(f"CATIA 视图截图保存失败。format=4: {exc}") from exc

    if not target_path.exists() or target_path.stat().st_size <= 0:
        raise RuntimeError(f"CATIA 视图截图保存失败，未生成有效 PNG 文件: {target_path}")
    return target_path


def capture_mirror_screenshot(
    document: Any,
    part: Any,
    hybrid_body: Any,
    mirror_feature: Any,
    mirror_side_name: str,
    up_direction: Vector,
    output_dir: Path,
    view_distance: float = SCREENSHOT_VIEW_DISTANCE,
) -> dict[str, Any]:
    """功能: 按固定视距截取单侧后视镜图片；输入: 文档、Part、几何图形集、镜片特征、侧别、上方向、输出目录和视距；输出: 截图结果字典。"""
    centroid = get_surface_center(document, part, mirror_feature)
    sphere_center_feature = create_point_center(
        part,
        hybrid_body,
        mirror_feature,
        f"{mirror_side_name}后视镜截图临时球心",
    )
    hide_object(document, sphere_center_feature)
    sphere_center = get_point(document, part, sphere_center_feature)
    sight_direction = normalize(
        subtract(sphere_center, centroid),
        f"{mirror_side_name}后视镜截图视角方向",
    )
    screenshot_path = build_screenshot_save_path(mirror_side_name, output_dir)
    result: dict[str, Any] = {
        "mirror_side_name": mirror_side_name,
        "target_point": round_vector(centroid),
        "sphere_center": round_vector(sphere_center),
        "view_direction": round_vector(sight_direction),
        "up_direction": round_vector(normalize(up_direction, "截图上方向")),
        "view_distance": float(view_distance),
        "requested_path": str(screenshot_path),
        "spec_tree_hidden_for_capture": False,
        "spec_tree_restored_after_capture": False,
    }
    try:
        viewer = configure_active_view(
            document,
            centroid,
            sight_direction,
            up_direction,
            view_distance,
        )
        saved_path = capture_viewer_to_file(viewer, screenshot_path, document)
        result.update({"success": True, "saved_path": str(saved_path)})
    except Exception as exc:
        result.update({"success": False, "error": str(exc)})
        print(f"{mirror_side_name}后视镜截图失败: {exc}", file=sys.stderr)
    return result


def capture_regulation_vision_screenshot(
    document: Any,
    left_regulation_line: dict[str, Any],
    right_regulation_line: dict[str, Any],
    ground_up_direction: Vector,
    vehicle_right_direction: Vector,
    output_dir: Path,
    view_distance: float = REGULATION_VISION_SCREENSHOT_VIEW_DISTANCE,
) -> dict[str, Any]:
    """功能: 按法规视野方向截取法规线总体图片；输入: 文档、左右法规线结果、地面法线、车辆右方向、输出目录和视距；输出: 截图结果字典。"""
    left_projection = tuple(
        float(value)
        for value in left_regulation_line["projection_point_coordinates"]["地面点到左车宽线投影点"]
    )
    right_point4 = tuple(
        float(value)
        for value in right_regulation_line["point_coordinates"]["右法规点4"]
    )
    target_point = average(left_projection, right_point4)
    sight_direction = normalize(
        scale(ground_up_direction, -1.0),
        "法规视野截图视角方向",
    )
    screenshot_path = build_named_screenshot_save_path("法规视野截图", output_dir)
    result: dict[str, Any] = {
        "name": "法规视野截图",
        "target_point": round_vector(target_point),
        "left_projection_point": round_vector(left_projection),
        "right_regulation_point4": round_vector(right_point4),
        "view_direction": round_vector(sight_direction),
        "up_direction": round_vector(normalize(vehicle_right_direction, "法规视野截图上方向")),
        "view_distance": float(view_distance),
        "requested_path": str(screenshot_path),
        "spec_tree_hidden_for_capture": False,
        "spec_tree_restored_after_capture": False,
    }
    try:
        viewer = configure_active_view(
            document,
            target_point,
            sight_direction,
            vehicle_right_direction,
            view_distance,
        )
        saved_path = capture_viewer_to_file(viewer, screenshot_path, document)
        result.update({"success": True, "saved_path": str(saved_path)})
    except Exception as exc:
        result.update({"success": False, "error": str(exc)})
        print(f"法规视野截图失败: {exc}", file=sys.stderr)
    return result


def choose_boundary_measure_point(
    document: Any,
    part: Any,
    gap_check_hybrid_body: Any,
    boundary: Any,
    first_coordinates: Vector,
    second_coordinates: Vector,
    point_name: str,
) -> tuple[Vector, Vector]:
    """功能: 判断最小距离测量点中哪一端属于边界；输入: 文档、Part、几何图形集、边界、两个候选坐标和点名；输出: 边界侧坐标和点侧坐标。"""
    first_temp = create_coordinate_point(
        part,
        gap_check_hybrid_body,
        first_coordinates,
        f"{point_name}测量端点1",
    )
    second_temp = create_coordinate_point(
        part,
        gap_check_hybrid_body,
        second_coordinates,
        f"{point_name}测量端点2",
    )
    hide_object(document, first_temp)
    hide_object(document, second_temp)
    first_to_boundary = get_minimum_distance(document, part, first_temp, boundary)
    second_to_boundary = get_minimum_distance(document, part, second_temp, boundary)
    if first_to_boundary <= second_to_boundary:
        return first_coordinates, second_coordinates
    return second_coordinates, first_coordinates


def add(left: Vector, right: Vector) -> Vector:
    """功能: 三维向量相加；输入: 两个 Vector；输出: 相加后的 Vector。"""
    return tuple(left[index] + right[index] for index in range(3))  # type: ignore[return-value]


def subtract(left: Vector, right: Vector) -> Vector:
    """功能: 三维向量相减；输入: 被减向量和减向量；输出: 差向量。"""
    return tuple(left[index] - right[index] for index in range(3))  # type: ignore[return-value]


def scale(vector: Vector, factor: float) -> Vector:
    """功能: 向量数乘；输入: Vector 和比例因子；输出: 缩放后的 Vector。"""
    return tuple(value * factor for value in vector)  # type: ignore[return-value]


def dot(left: Vector, right: Vector) -> float:
    """功能: 计算点积；输入: 两个 Vector；输出: 标量点积。"""
    return sum(left[index] * right[index] for index in range(3))


def cross(left: Vector, right: Vector) -> Vector:
    """功能: 计算叉积；输入: 两个 Vector；输出: 法向 Vector。"""
    return (
        left[1] * right[2] - left[2] * right[1],
        left[2] * right[0] - left[0] * right[2],
        left[0] * right[1] - left[1] * right[0],
    )


def normalize(vector: Vector, label: str) -> Vector:
    """功能: 单位化向量；输入: Vector 和错误提示标签；输出: 单位 Vector。"""
    length = math.sqrt(dot(vector, vector))
    if length <= 1e-9:
        raise ValueError(f"无法确定{label}：向量长度接近 0。")
    return scale(vector, 1.0 / length)


def project_to_plane(vector: Vector, plane_normal: Vector) -> Vector:
    """功能: 将向量投影到给定法向的平面内；输入: 向量和平面法向；输出: 投影向量。"""
    return subtract(vector, scale(plane_normal, dot(vector, plane_normal)))


def orient_vector(vector: Vector, preferred_direction: Vector) -> Vector:
    """功能: 按参考方向调整向量正负；输入: 向量和期望方向；输出: 同向或反向后的向量。"""
    return vector if dot(vector, preferred_direction) >= 0 else scale(vector, -1.0)


def snap_direction_to_global_axis(
    vector: Vector,
    label: str,
    angle_tolerance_degrees: float = AXIS_SNAP_ANGLE_DEGREES,
) -> Vector:
    """功能: 将接近全局 XYZ 轴的方向吸附到对应正负轴；输入: 方向、标签和角度容差；输出: 吸附后方向。"""
    unit_vector = normalize(vector, label)
    axes: tuple[tuple[str, Vector], ...] = (
        ("+X", (1.0, 0.0, 0.0)),
        ("-X", (-1.0, 0.0, 0.0)),
        ("+Y", (0.0, 1.0, 0.0)),
        ("-Y", (0.0, -1.0, 0.0)),
        ("+Z", (0.0, 0.0, 1.0)),
        ("-Z", (0.0, 0.0, -1.0)),
    )
    best_axis_name, best_axis = max(
        axes,
        key=lambda axis_item: dot(unit_vector, axis_item[1]),
    )
    best_cosine = max(-1.0, min(1.0, dot(unit_vector, best_axis)))
    angle_degrees = math.degrees(math.acos(best_cosine))
    if angle_degrees <= angle_tolerance_degrees:
        print(
            f"{label}接近全局{best_axis_name}轴，已吸附为{best_axis_name}方向"
            f"（夹角{angle_degrees:.2f}°）",
            flush=True,
        )
        return best_axis
    return unit_vector


def average(left: Vector, right: Vector) -> Vector:
    """功能: 计算两个点/向量的中点；输入: 两个 Vector；输出: 平均 Vector。"""
    return scale(add(left, right), 0.5)


def calculate_vehicle_directions(
    document: Any,
    part: Any,
    features: dict[str, Any],
) -> dict[str, Vector]:
    """
    功能: 根据地面、眼点、镜片和车宽线计算车辆方向；输入: 文档、Part 和输入特征字典；输出: 方向向量字典。

    地面法线由空载地面的两个平面内方向叉乘得到，并朝向右眼点一侧；
    若接近全局 XYZ 正负轴，则吸附为对应轴向。左右镜片中心的连线用于
    确定粗略车辆右方向，镜片中心均值指向眼点均值的方向用于确定粗略
    后方向；后续法规线构造会再通过左车宽线延长线修正车辆后方向。
    """
    right_eye = get_point(document, part, features["right_eye"])
    left_eye = get_point(document, part, features["left_eye"])
    left_mirror_center = get_surface_center(document, part, features["left_mirror"])
    right_mirror_center = get_surface_center(document, part, features["right_mirror"])
    ground_origin, ground_x, ground_y = get_plane(document, part, features["ground"])

    up_direction = normalize(cross(ground_x, ground_y), "地面法线方向")
    up_direction = orient_vector(up_direction, subtract(right_eye, ground_origin))
    up_direction = snap_direction_to_global_axis(up_direction, "地面法线方向")

    mirror_left_to_right = subtract(right_mirror_center, left_mirror_center)
    right_direction = normalize(project_to_plane(mirror_left_to_right, up_direction), "车辆右方向")

    mirror_center = average(left_mirror_center, right_mirror_center)
    eye_center = average(left_eye, right_eye)
    mirror_to_eye = project_to_plane(subtract(eye_center, mirror_center), up_direction)

    # 车辆后方向需要与车辆右方向正交；用镜片到眼点方向消除正负歧义。
    rear_direction = normalize(cross(right_direction, up_direction), "车辆后方向")
    rear_direction = orient_vector(rear_direction, mirror_to_eye)
    right_direction = normalize(cross(up_direction, rear_direction), "车辆右方向")
    left_direction = scale(right_direction, -1.0)
    front_direction = scale(rear_direction, -1.0)

    return {
        "地面法线/上方向": up_direction,
        "车辆后方向": rear_direction,
        "车辆粗略后方向": rear_direction,
        "车辆前方向": front_direction,
        "车辆左方向": left_direction,
        "车辆右方向": right_direction,
    }


def format_vector(vector: Vector) -> str:
    """功能: 格式化三维向量；输入: Vector；输出: 保留两位小数的字符串。"""
    return f"({vector[0]:.2f}, {vector[1]:.2f}, {vector[2]:.2f})"


def round_vector(vector: Vector) -> list[float]:
    """功能: 向量数值四舍五入；输入: Vector；输出: 保留两位小数的列表。"""
    return [round(value, 2) for value in vector]


def vector_length(vector: Vector) -> float:
    """功能: 计算向量长度；输入: Vector；输出: 欧氏长度。"""
    return math.sqrt(dot(vector, vector))


def distance(left: Vector, right: Vector) -> float:
    """功能: 计算两点距离；输入: 两个 Vector；输出: 欧氏距离。"""
    return vector_length(subtract(left, right))


def polygon_area_2d(points: list[tuple[float, float]]) -> float:
    """功能: 计算二维多边形面积；输入: 按顺序排列的二维点；输出: 面积。"""
    return abs(
        sum(
            points[index][0] * points[(index + 1) % len(points)][1]
            - points[index][1] * points[(index + 1) % len(points)][0]
            for index in range(len(points))
        )
    ) / 2.0


def order_points_around_center(
    points: list[tuple[float, float]],
) -> list[int]:
    """功能: 按点绕中心的极角排序；输入: 二维点列表；输出: 排序后的原索引列表。"""
    center_x = sum(point[0] for point in points) / len(points)
    center_y = sum(point[1] for point in points) / len(points)
    return sorted(
        range(len(points)),
        key=lambda index: math.atan2(
            points[index][1] - center_y,
            points[index][0] - center_x,
        ),
    )


def is_convex_quadrilateral(points: list[tuple[float, float]]) -> bool:
    """功能: 判断四个二维点是否构成凸四边形；输入: 四个二维点；输出: 是否凸。"""
    cross_values: list[float] = []
    for index in range(4):
        first = points[index]
        second = points[(index + 1) % 4]
        third = points[(index + 2) % 4]
        cross_value = (
            (second[0] - first[0]) * (third[1] - second[1])
            - (second[1] - first[1]) * (third[0] - second[0])
        )
        if abs(cross_value) > 1e-9:
            cross_values.append(cross_value)
    return bool(cross_values) and (
        all(value > 0.0 for value in cross_values)
        or all(value < 0.0 for value in cross_values)
    )


def find_maximum_area_quadrilateral(
    sphere_center: Vector,
    point_records: list[dict[str, Any]],
) -> tuple[list[dict[str, Any]], list[tuple[float, float]], float]:
    """功能: 在镜片局部切平面中选取面积最大的凸四边形；输入: 球心坐标和反射点记录；输出: 选中点记录、局部二维坐标和面积。"""
    if len(point_records) != 12:
        raise ValueError(f"最佳视野计算需要 12 个点，实际为 {len(point_records)}。")

    coordinates = [
        tuple(record["_coordinates_full"]) for record in point_records
    ]
    average_point = scale(
        tuple(sum(point[index] for point in coordinates) for index in range(3)),
        1.0 / len(coordinates),
    )
    normal = normalize(subtract(average_point, sphere_center), "镜片局部法线")
    tangent_candidates = [
        project_to_plane(subtract(point, average_point), normal)
        for point in coordinates
    ]
    tangent_x = max(tangent_candidates, key=vector_length)
    axis_x = normalize(tangent_x, "镜片局部横向")
    axis_y = normalize(cross(normal, axis_x), "镜片局部纵向")
    local_points = [
        (
            dot(subtract(point, average_point), axis_x),
            dot(subtract(point, average_point), axis_y),
        )
        for point in coordinates
    ]

    best_indices: list[int] | None = None
    best_local_points: list[tuple[float, float]] = []
    best_area = -1.0
    for combination_indices in combinations(range(len(point_records)), 4):
        selected = [local_points[index] for index in combination_indices]
        order = order_points_around_center(selected)
        ordered_points = [selected[index] for index in order]
        if not is_convex_quadrilateral(ordered_points):
            continue
        area = polygon_area_2d(ordered_points)
        if area > best_area:
            best_area = area
            best_indices = [combination_indices[index] for index in order]
            best_local_points = ordered_points

    if best_indices is None:
        raise RuntimeError("无法从镜片的 12 个反射点中确定最佳视野四边形。")
    return [point_records[index] for index in best_indices], best_local_points, best_area


def order_selected_quadrilateral(
    sphere_center: Vector,
    point_records: list[dict[str, Any]],
) -> tuple[list[dict[str, Any]], list[tuple[float, float]], float, bool]:
    """功能: 将指定4个反射点按镜片局部平面极角排序并计算面积；输入: 球心和4个点记录；输出: 排序点、局部坐标、面积和是否凸。"""
    if len(point_records) != 4:
        raise ValueError(f"指定最佳视野点需要 4 个点，实际为 {len(point_records)}。")
    coordinates = [tuple(record["_coordinates_full"]) for record in point_records]
    average_point = scale(
        tuple(sum(point[index] for point in coordinates) for index in range(3)),
        1.0 / len(coordinates),
    )
    normal = normalize(subtract(average_point, sphere_center), "镜片局部法线")
    tangent_candidates = [
        project_to_plane(subtract(point, average_point), normal)
        for point in coordinates
    ]
    tangent_x = max(tangent_candidates, key=vector_length)
    axis_x = normalize(tangent_x, "镜片局部横向")
    axis_y = normalize(cross(normal, axis_x), "镜片局部纵向")
    local_points = [
        (
            dot(subtract(point, average_point), axis_x),
            dot(subtract(point, average_point), axis_y),
        )
        for point in coordinates
    ]
    order = order_points_around_center(local_points)
    ordered_records = [point_records[index] for index in order]
    ordered_local_points = [local_points[index] for index in order]
    return (
        ordered_records,
        ordered_local_points,
        polygon_area_2d(ordered_local_points),
        is_convex_quadrilateral(ordered_local_points),
    )


def create_best_view_frame(
    document: Any,
    part: Any,
    hybrid_body: Any,
    mirror_side_name: str,
    sphere_center: Vector,
    point_records: list[dict[str, Any]],
    final_view_point_names: list[str] | None = None,
) -> dict[str, Any]:
    """功能: 创建一侧镜片最佳视野红色线框；输入: 文档、Part、几何图形集、侧别、球心、点记录和可选最终点名；输出: 线框结果字典。"""
    selection_source = "maximum_area_quadrilateral"
    is_convex = True
    if final_view_point_names:
        records_by_name = {
            str(record["reflection_point_name"]): record
            for record in point_records
        }
        missing_names = [
            point_name
            for point_name in final_view_point_names
            if point_name not in records_by_name
        ]
        if missing_names:
            raise RuntimeError(
                f"{mirror_side_name}镜片最佳视野线框缺少正式法规反射取点: {missing_names}"
            )
        selected_records, local_coordinates, area, is_convex = order_selected_quadrilateral(
            sphere_center,
            [records_by_name[point_name] for point_name in final_view_point_names],
        )
        selection_source = "temporary_final_view_points"
    else:
        selected_records, local_coordinates, area = find_maximum_area_quadrilateral(
            sphere_center,
            point_records,
        )
    frame_lines: list[Any] = []
    for index in range(4):
        line = create_point_to_point_line(
            part,
            hybrid_body,
            selected_records[index]["_feature"],
            selected_records[(index + 1) % 4]["_feature"],
            f"{mirror_side_name}镜片最佳视野线框{index + 1}",
        )
        set_object_color(document, line, 255, 0, 0)
        frame_lines.append(line)

    return {
        "mirror_side_name": mirror_side_name,
        "boundary_point_names": [
            record["reflection_point_name"] for record in selected_records
        ],
        "boundary_point_coordinates": [
            round_vector(record["_coordinates_full"]) for record in selected_records
        ],
        "boundary_point_coordinates_full": [
            tuple(record["_coordinates_full"]) for record in selected_records
        ],
        "boundary_local_coordinates": [
            [round(value, 2) for value in coordinates]
            for coordinates in local_coordinates
        ],
        "quadrilateral_area": round(area, 2),
        "selection_source": selection_source,
        "is_convex_quadrilateral": is_convex,
        "frame_line_names": [str(line.Name) for line in frame_lines],
        "frame_color": [255, 0, 0],
    }


def collect_best_view_distance_annotation_data(
    document: Any,
    part: Any,
    mirror_data: dict[str, Any],
    best_view_frames: dict[str, Any],
    selected_mirror_sides: set[str],
) -> dict[str, Any]:
    """功能: 采集最佳视野边界点到镜片边界的标注数据；输入: 文档、Part、镜片、最佳视野结果和有效左/右镜片；输出: 标注数据。"""
    measurement_body = create_hybrid_body(part, "距离标注数据测量")
    annotation_groups: dict[str, Any] = {}
    try:
        for mirror_side_name in ("左", "右"):
            if mirror_side_name not in selected_mirror_sides:
                annotation_groups[mirror_side_name] = {
                    "skipped": True,
                    "reason": "该侧未创建法规反射取点",
                    "items": [],
                }
                continue

            frame_result = best_view_frames.get(mirror_side_name, {})
            vertex_coordinates_list = frame_result.get("boundary_point_coordinates_full", [])
            if not vertex_coordinates_list:
                annotation_groups[mirror_side_name] = {
                    "skipped": True,
                    "reason": "未找到最佳视野边界点",
                    "items": [],
                }
                continue

            mirror_feature = mirror_data[mirror_side_name]
            boundary = create_boundary(
                part,
                measurement_body,
                mirror_feature,
                f"{mirror_side_name}镜片标注测量边界",
            )
            items: list[dict[str, Any]] = []
            for index, vertex_coordinates in enumerate(vertex_coordinates_list, start=1):
                vertex_point = create_coordinate_point(
                    part,
                    measurement_body,
                    tuple(float(value) for value in vertex_coordinates),
                    f"{mirror_side_name}镜片最佳视野标注点{index}",
                )
                first_coordinates, second_coordinates, _raw_coordinates = (
                    get_minimum_distance_points(
                        document,
                        part,
                        vertex_point,
                        boundary,
                    )
                )
                boundary_coordinates, measured_vertex_coordinates = choose_boundary_measure_point(
                    document,
                    part,
                    measurement_body,
                    boundary,
                    first_coordinates,
                    second_coordinates,
                    f"{mirror_side_name}镜片最佳视野标注点{index}",
                )
                distance_to_boundary = distance(
                    measured_vertex_coordinates,
                    boundary_coordinates,
                )
                items.append(
                    {
                        "name": f"{mirror_side_name}镜片最佳视野边界点{index}到镜片边界距离",
                        "mirror_side_name": mirror_side_name,
                        "index": index,
                        "best_view_point": round_vector(measured_vertex_coordinates),
                        "nearest_boundary_point": round_vector(boundary_coordinates),
                        "distance": round(distance_to_boundary, 3),
                        "label": f"{distance_to_boundary:.3f} mm",
                    }
                )

            annotation_groups[mirror_side_name] = {
                "skipped": False,
                "items": items,
            }
    finally:
        try:
            delete_object(document, measurement_body)
        except Exception:
            try:
                hide_object(document, measurement_body)
            except Exception:
                pass

    return {
        "groups": annotation_groups,
        "item_count": sum(
            len(group.get("items", []))
            for group in annotation_groups.values()
        ),
    }


def add_component_from_file_to_product(document: Any, root_product: Any, file_path: Path) -> Any:
    """功能: 将已有 CATPart 文件装配到 CATProduct 根产品；输入: ProductDocument、根 Product 和文件路径；输出: 新增子 Product。"""
    vba_code = """
Public Function add_component(rootProduct, componentPath)
    Dim fileList(0)
    fileList(0) = componentPath
    rootProduct.Products.AddComponentsFromFiles fileList, "All"
    add_component = True
End Function
"""
    try:
        before_count = int(root_product.Products.Count)
    except Exception:
        before_count = 0
    try:
        document.Application.SystemService.Evaluate(
            vba_code,
            0,
            "add_component",
            [root_product, str(file_path)],
        )
    except Exception as exc:
        raise RuntimeError(f"无法将结果 CATPart 装配到 CATProduct: {file_path}") from exc
    try:
        after_count = int(root_product.Products.Count)
        if after_count > before_count:
            return root_product.Products.Item(after_count)
    except Exception:
        pass
    raise RuntimeError(f"无法获取装配后的 CATPart 子组件: {file_path}")


def get_part_from_product_component(component_product: Any) -> Any:
    """功能: 尽力从 Product 子组件获取 Part；输入: 子 Product；输出: Part。"""
    attempts = (
        lambda product: product.ReferenceProduct.Parent.Part,
        lambda product: product.ReferenceProduct.Parent.GetItem("Part"),
        lambda product: product.Parent.Part,
    )
    for attempt in attempts:
        try:
            part = attempt(component_product)
            if part is not None:
                return part
        except Exception:
            pass
    raise RuntimeError("无法从标注承载 Product 获取 Part。")


def get_part_document_from_product_component(component_product: Any, part: Any) -> Any:
    """功能: 尽力从 Product 子组件获取 PartDocument；输入: 子 Product 和 Part；输出: PartDocument。"""
    attempts = (
        lambda: component_product.ReferenceProduct.Parent,
        lambda: part.Parent,
        lambda: component_product.Parent,
    )
    for attempt in attempts:
        try:
            document = attempt()
            if document is not None and hasattr(document, "SaveAs"):
                return document
        except Exception:
            pass
    raise RuntimeError("无法从标注承载 Product 获取 PartDocument。")


def get_part_from_product_document(product_document: Any, part_number: str) -> Any:
    """功能: 在 ProductDocument 中按 PartNumber 查找子组件 Part；输入: ProductDocument 和 PartNumber；输出: Part。"""
    root_product = product_document.Product
    for product in iter_collection(root_product.Products):
        try:
            if str(product.PartNumber) == part_number:
                return get_part_from_product_component(product)
        except Exception:
            pass
    raise RuntimeError(f"无法在 CATProduct 中找到 PartNumber 为 {part_number} 的组件。")


def get_marker3ds(root_product: Any) -> Any | None:
    """功能: 获取 Product 的 3D 标记集合；输入: 根 Product；输出: Marker3Ds 或 None。"""
    try:
        return root_product.GetTechnologicalObject("Marker3Ds")
    except Exception:
        return None


def activate_feature_for_display_refresh(document: Any, feature: Any) -> None:
    """功能: 通过选择对象触发 CATIA 显示属性刷新；输入: 文档和特征；输出: 无。"""
    try:
        try:
            document.Activate()
        except Exception:
            pass
        selection = document.Selection
        selection.Clear()
        selection.Add(feature)
        try:
            selection.VisProperties.SetShow(0)
        except Exception:
            pass
        try:
            viewer = document.Application.ActiveWindow.ActiveViewer
            viewer.Update()
            try:
                viewer.Reframe()
            except Exception:
                pass
            viewer.Update()
        except Exception:
            pass
        time.sleep(0.1)
        selection.Clear()
    except Exception:
        pass


def create_annotation_text_marker(
    product_document: Any,
    root_product: Any,
    annotation_part: Any,
    support_feature: Any,
    text_position: Vector,
    anchor_position: Vector,
    label: str,
    name: str,
) -> Any | None:
    """功能: 创建 3D 文本标记并触发显示刷新；输入: ProductDocument、根 Product、支撑几何、文本位置、锚点、文本和名称；输出: 文本标记或 None。"""
    marker3ds = get_marker3ds(root_product)
    if marker3ds is None:
        return None
    try:
        try:
            product_document.Activate()
        except Exception:
            pass
        activate_feature_for_display_refresh(product_document, support_feature)
        try:
            annotation_part.Update()
        except Exception:
            pass
        try:
            root_product.Update()
        except Exception:
            pass
        marker = marker3ds.Add3DText(
            tuple(float(value) for value in text_position),
            label,
            tuple(float(value) for value in anchor_position),
            support_feature,
        )
        try:
            marker.Name = name
        except Exception:
            pass
        try:
            marker.TextSize = ANNOTATION_TEXT_SIZE
        except Exception:
            pass
        try:
            root_product.Update()
        except Exception:
            pass
        activate_feature_for_display_refresh(product_document, marker)
        try:
            product_document.Save()
        except Exception:
            pass
        return marker
    except Exception:
        return None


def create_distance_annotation_line_geometry(
    document: Any,
    part: Any,
    hybrid_body: Any,
    item: dict[str, Any],
) -> dict[str, Any]:
    """功能: 在标注 CATPart 中创建黑色标注点和距离线；输入: 文档、Part、几何集和标注项；输出: 线标注基础数据。"""
    point1 = tuple(float(value) for value in item["best_view_point"])
    point2 = tuple(float(value) for value in item["nearest_boundary_point"])
    label = str(item["label"])
    name = str(item["name"])
    line_name = f"{name}_Line"

    # 标注几何放在独立 Part 中，避免修改已另存的计算结果 CATPart。
    point_feature1 = create_coordinate_point(part, hybrid_body, point1, f"{name}_P1")
    point_feature2 = create_coordinate_point(part, hybrid_body, point2, f"{name}_P2")
    line = create_point_to_point_line(part, hybrid_body, point_feature1, point_feature2, line_name)
    set_object_visual_style(document, point_feature1, *ANNOTATION_COLOR)
    set_object_visual_style(document, point_feature2, *ANNOTATION_COLOR)
    set_object_visual_style(document, line, *ANNOTATION_COLOR, width=2)

    return {
        "name": name,
        "distance": item["distance"],
        "label": label,
        "line_name": str(line.Name),
        "best_view_point": point1,
        "nearest_boundary_point": point2,
        "text_created": False,
    }


def create_distance_annotation_text_marker(
    product_document: Any,
    root_product: Any,
    annotation_part: Any,
    line_feature: Any,
    item: dict[str, Any],
    vehicle_right_direction: Vector,
    text_offset_distance: float = 5.0,
) -> dict[str, Any]:
    """功能: 在 CATProduct 中为距离线创建 3D 文本标记；输入: ProductDocument、根 Product、标注 Part、支撑线、标注数据和车辆右方向；输出: 文本创建结果。"""
    point1 = tuple(float(value) for value in item["best_view_point"])
    point2 = tuple(float(value) for value in item["nearest_boundary_point"])
    label = str(item["label"])
    name = str(item["name"])
    midpoint = average(point1, point2)
    text_offset = scale(
        normalize(vehicle_right_direction, "标注文本车辆右方向"),
        text_offset_distance,
    )
    text_position = add(midpoint, text_offset)
    marker = create_annotation_text_marker(
        product_document,
        root_product,
        annotation_part,
        line_feature,
        text_position,
        midpoint,
        label,
        f"{name}_Text",
    )
    if marker is not None:
        try:
            set_object_visual_style(product_document, marker, *ANNOTATION_COLOR, width=2)
            activate_feature_for_display_refresh(product_document, marker)
        except Exception:
            pass

    return {
        "name": name,
        "distance": item["distance"],
        "label": label,
        "line_name": str(line_feature.Name),
        "text_created": marker is not None,
    }


def refresh_annotation_product_view(
    product_document: Any,
    root_product: Any,
    annotation_part: Any,
) -> None:
    """功能: 刷新标注装配和当前视图；输入: ProductDocument、根 Product 和标注 Part；输出: 无。"""
    try:
        annotation_part.Update()
    except Exception:
        pass
    try:
        root_product.Update()
    except Exception:
        pass
    try:
        product_document.Activate()
    except Exception:
        pass
    try:
        viewer = product_document.Application.ActiveWindow.ActiveViewer
        try:
            viewer.Update()
        except Exception:
            pass
        try:
            viewer.Reframe()
        except Exception:
            pass
        try:
            viewer.Update()
        except Exception:
            pass
    except Exception:
        pass


def create_annotation_product_file(
    catia: Any,
    saved_catpart_path: Path,
    annotation_data: dict[str, Any],
    output_dir: Path,
    vehicle_right_direction: Vector,
) -> dict[str, Any]:
    """功能: 新建 CATProduct、装配结果 CATPart 并创建距离标注；输入: CATIA、结果 CATPart、标注数据、输出目录和车辆右方向；输出: 标注文件结果。"""
    annotation_path = build_annotation_product_save_path(output_dir)
    annotation_part_path = build_annotation_part_save_path(output_dir)

    annotation_part_document = catia.Documents.Add("Part")
    annotation_part = annotation_part_document.Part
    try:
        annotation_part.PartNumber = ANNOTATION_PART_NUMBER
    except Exception:
        pass
    try:
        annotation_part_document.Product.PartNumber = ANNOTATION_PART_NUMBER
        annotation_part_document.Product.Name = "距离标注"
    except Exception:
        pass
    confirm_watcher = start_catia_confirm_dialog_watcher(catia)
    annotation_part_document.SaveAs(str(annotation_part_path))
    confirm_watcher.join(timeout=1.0)
    print(f"标注 CATPart 已创建: {annotation_part_path}", flush=True)

    annotation_body = create_hybrid_body(annotation_part, "距离标注")
    created_items: list[dict[str, Any]] = []
    source_items: list[dict[str, Any]] = []
    for group in annotation_data.get("groups", {}).values():
        if group.get("skipped"):
            continue
        for item in group.get("items", []):
            try:
                created_items.append(
                    create_distance_annotation_line_geometry(
                        annotation_part_document,
                        annotation_part,
                        annotation_body,
                        item,
                    )
                )
                source_items.append(item)
            except Exception as exc:
                created_items.append(
                    {
                        "name": item.get("name"),
                        "distance": item.get("distance"),
                        "label": item.get("label"),
                        "error": str(exc),
                        "text_created": False,
                    }
                )

    try:
        annotation_part.Update()
    except Exception:
        pass
    annotation_part_document.Save()
    print(f"标注 CATPart 已保存: {annotation_part_path}", flush=True)

    product_document = catia.Documents.Add("Product")
    root_product = product_document.Product
    try:
        root_product.PartNumber = "外后视镜视野校核标注"
        root_product.Name = "外后视镜视野校核标注"
    except Exception:
        pass

    add_component_from_file_to_product(product_document, root_product, saved_catpart_path)
    annotation_component = add_component_from_file_to_product(
        product_document,
        root_product,
        annotation_part_path,
    )
    annotation_part = get_part_from_product_component(annotation_component)
    annotation_part_document = get_part_document_from_product_component(
        annotation_component,
        annotation_part,
    )
    annotation_body = find_hybrid_body(annotation_part, "距离标注")
    try:
        annotation_part.PartNumber = ANNOTATION_PART_NUMBER
    except Exception:
        pass

    text_results: list[dict[str, Any]] = []
    for line_item, source_item in zip(
        [item for item in created_items if item.get("line_name")],
        source_items,
    ):
        try:
            line_feature = find_feature(annotation_body, line_item["line_name"])
            text_result = create_distance_annotation_text_marker(
                product_document,
                root_product,
                annotation_part,
                line_feature,
                source_item,
                vehicle_right_direction,
            )
            line_item.update(text_result)
            text_results.append(text_result)
        except Exception as exc:
            line_item.update(
                {
                    "error": str(exc),
                    "text_created": False,
                }
            )

    try:
        annotation_part.Update()
    except Exception:
        pass
    try:
        annotation_part_document.Save()
    except Exception:
        pass

    refresh_annotation_product_view(product_document, root_product, annotation_part)
    product_document.SaveAs(str(annotation_path))
    print(f"标注 CATProduct 已保存: {annotation_path}", flush=True)

    try:
        product_document.Close()
    except Exception:
        pass
    product_document = catia.Documents.Open(str(annotation_path))
    root_product = product_document.Product
    try:
        annotation_part = get_part_from_product_document(
            product_document,
            ANNOTATION_PART_NUMBER,
        )
        refresh_annotation_product_view(product_document, root_product, annotation_part)
    except Exception:
        try:
            product_document.Activate()
            product_document.Application.ActiveWindow.ActiveViewer.Update()
        except Exception:
            pass
    print(f"标注 CATProduct 已重新打开: {annotation_path}", flush=True)

    return {
        "annotation_product_path": str(annotation_path),
        "annotation_part_path": str(annotation_part_path),
        "_product_document": product_document,
        "annotation_count": len(created_items),
        "items": created_items,
        "text_items": text_results,
    }


def create_reflection_triangle_frames(
    document: Any,
    part: Any,
    hybrid_body: Any,
    point_records: list[dict[str, Any]],
) -> dict[str, Any]:
    """功能: 按眼点和法规侧别创建 1-2-3、4-5-6 两组三角线框；输入: 文档、Part、几何图形集和点记录；输出: 三角线框结果字典。"""
    triangle_edges = (
        (1, 1, 2),
        (2, 2, 3),
        (3, 3, 1),
        (4, 4, 5),
        (5, 5, 6),
        (6, 6, 4),
    )
    frame_results: dict[str, Any] = {}

    for eye_side_name in ("右", "左"):
        for mirror_side_name in ("左", "右"):
            group_records = {
                int(record["point_index"]): record
                for record in point_records
                if record.get("eye_side_name") == eye_side_name
                and record.get("mirror_side_name") == mirror_side_name
            }
            group_name = f"{eye_side_name}眼{mirror_side_name}法规反射取点"
            if not group_records:
                frame_results[group_name] = {
                    "skipped": True,
                    "reason": "临时法规反射取点测量未通过",
                    "line_names": [],
                }
                continue
            if len(group_records) != 6:
                raise RuntimeError(
                    f"{group_name}三角线框需要 6 个反射点，实际为 {len(group_records)}。"
                )

            line_names: list[str] = []
            for line_index, start_index, end_index in triangle_edges:
                line = create_point_to_point_line(
                    part,
                    hybrid_body,
                    group_records[start_index]["_feature"],
                    group_records[end_index]["_feature"],
                    f"{group_name}线框{line_index}",
                )
                if eye_side_name == "右":
                    set_object_color(document, line, 255, 165, 0)
                else:
                    set_object_color(document, line, 255, 255, 0)
                line_names.append(str(line.Name))

            frame_results[group_name] = {
                "skipped": False,
                "line_names": line_names,
                "line_color": [255, 165, 0] if eye_side_name == "右" else [255, 255, 0],
                "triangle_point_groups": [
                    [
                        group_records[1]["reflection_point_name"],
                        group_records[2]["reflection_point_name"],
                        group_records[3]["reflection_point_name"],
                    ],
                    [
                        group_records[4]["reflection_point_name"],
                        group_records[5]["reflection_point_name"],
                        group_records[6]["reflection_point_name"],
                    ],
                ],
            }

    return frame_results


def create_gap_check(
    document: Any,
    part: Any,
    gap_check_hybrid_body_name: str,
    left_mirror: Any,
    right_mirror: Any,
    reflection_point_hybrid_body: Any,
    reflection_points: list[dict[str, Any]],
    clearance_threshold: float = 3.0,
    selected_mirror_sides: set[str] | None = None,
) -> dict[str, Any]:
    """功能: 提取镜片边界并测量每侧反射点到边界的间隙；输入: 文档、Part、几何集名、左右镜片、反射点集合、阈值和镜片侧；输出: 间隙校验结果字典。"""
    gap_check_hybrid_body = create_hybrid_body(part, gap_check_hybrid_body_name)
    mirror_inputs = {
        "左": (left_mirror, "左镜片边界"),
        "右": (right_mirror, "右镜片边界"),
    }
    if selected_mirror_sides is None:
        selected_mirror_sides = set(mirror_inputs)
    mirror_results: dict[str, Any] = {}

    for mirror_side_name, (mirror, boundary_name) in mirror_inputs.items():
        if mirror_side_name not in selected_mirror_sides:
            mirror_results[mirror_side_name] = {
                "boundary_name": None,
                "boundary_hidden": False,
                "threshold": clearance_threshold,
                "points": [],
                "all_points_pass": False,
                "skipped": True,
            }
            continue
        boundary = create_boundary(
            part,
            gap_check_hybrid_body,
            mirror,
            boundary_name,
        )
        point_results: list[dict[str, Any]] = []
        side_points = [
            point
            for point in reflection_points
            if point.get("mirror_side_name") == mirror_side_name
        ]
        if len(side_points) != 12:
            raise RuntimeError(
                f"{mirror_side_name}镜片间隙校验需要 12 个反射点，实际为 {len(side_points)}。"
            )

        for point_result in side_points:
            point_feature = find_feature(
                reflection_point_hybrid_body,
                point_result["reflection_point_name"],
            )
            clearance = get_minimum_distance(document, part, point_feature, boundary)
            point_side_coordinates, boundary_side_coordinates, raw_measurement_coordinates = (
                get_minimum_distance_points(
                    document,
                    part,
                    point_feature,
                    boundary,
                )
            )
            boundary_side_coordinates, point_side_coordinates = choose_boundary_measure_point(
                document,
                part,
                gap_check_hybrid_body,
                boundary,
                point_side_coordinates,
                boundary_side_coordinates,
                point_result["reflection_point_name"],
            )
            point_results.append(
                {
                    "point_name": point_result["reflection_point_name"],
                    "distance_to_boundary": round(clearance, 2),
                    "point_side_measure_coordinates": round_vector(point_side_coordinates),
                    "boundary_measure_point_coordinates": round_vector(boundary_side_coordinates),
                    "minimum_distance_points_raw": [
                        round(value, 2) for value in raw_measurement_coordinates
                    ],
                    "is_clearance_greater_equal_3mm": clearance >= clearance_threshold,
                }
            )

        hide_object(document, boundary)
        mirror_results[mirror_side_name] = {
            "boundary_name": str(boundary.Name),
            "boundary_hidden": True,
            "threshold": clearance_threshold,
            "points": point_results,
            "all_points_pass": all(
                point["is_clearance_greater_equal_3mm"]
                for point in point_results
            ),
        }

    return {
        "geo_set_name": str(gap_check_hybrid_body.Name),
        "clearance_threshold": clearance_threshold,
        "mirrors": mirror_results,
    }


def compare_gap_check_with_temporary_points(
    gap_check_result: dict[str, Any],
    temporary_reflection_points: dict[str, Any],
) -> dict[str, Any]:
    """功能: 对比 CATIA 过程测量结果与临时算法测量结果；输入: 间隙校验结果和临时反射点结果；输出: 对比结果字典。"""
    comparison_results: dict[str, Any] = {}
    for mirror_side_name in ("左", "右"):
        temporary_points = {
            point["reflection_point_name"]: point
            for point in temporary_reflection_points.get("mirrors", {})
            .get(mirror_side_name, {})
            .get("reflection_points", [])
        }
        gap_points = (
            gap_check_result.get("mirrors", {})
            .get(mirror_side_name, {})
            .get("points", [])
        )
        point_results = []
        for gap_point in gap_points:
            point_name = gap_point.get("point_name")
            temporary_point = temporary_points.get(point_name)
            if not temporary_point:
                point_results.append(
                    {
                        "point_name": point_name,
                        "matched": False,
                    }
                )
                continue
            temporary_distance = temporary_point.get("distance_to_boundary")
            catia_distance = gap_point.get("distance_to_boundary")
            point_results.append(
                {
                    "point_name": point_name,
                    "matched": True,
                    "temporary_distance_to_boundary": temporary_distance,
                    "catia_distance_to_boundary": catia_distance,
                    "distance_difference": (
                        round(catia_distance - temporary_distance, 2)
                        if catia_distance is not None and temporary_distance is not None
                        else None
                    ),
                }
            )
        comparison_results[mirror_side_name] = {
            "points": point_results,
            "max_abs_distance_difference": (
                max(
                    abs(point["distance_difference"])
                    for point in point_results
                    if point.get("distance_difference") is not None
                )
                if any(point.get("distance_difference") is not None for point in point_results)
                else None
            ),
        }
    return {
        "mirrors": comparison_results,
    }


def solve_spherical_reflection_point(
    sphere_center: Vector,
    sphere_radius: float,
    regulation_point: Vector,
    eye_point: Vector,
) -> Vector:
    """
    功能: 求解球面上的法规反射点；输入: 球心、半径、法规点和眼点；输出: 反射点坐标。

    反射点位于球心、法规点、眼点确定的截面圆上。球面法线与入射/出射单位
    方向的角平分线共线，因此在一维角度参数上求解该条件。
    """
    if sphere_radius <= 1e-9:
        raise ValueError("镜片球面半径过小，无法计算反射取点。")

    center_to_regulation = subtract(regulation_point, sphere_center)
    center_to_eye = subtract(eye_point, sphere_center)
    plane_normal = cross(center_to_regulation, center_to_eye)
    if vector_length(plane_normal) <= 1e-9:
        raise ValueError("球心、法规点、眼点近似共线，无法确定反射截面。")

    axis_u = normalize(project_to_plane(center_to_regulation, normalize(plane_normal, "反射截面法线")), "反射截面横向")
    axis_w = normalize(cross(normalize(plane_normal, "反射截面法线"), axis_u), "反射截面纵向")

    def point_at(theta: float) -> Vector:
        """功能: 计算截面圆上指定角度点；输入: 参数角 theta；输出: 球面候选点坐标。"""
        return add(
            sphere_center,
            add(
                scale(axis_u, sphere_radius * math.cos(theta)),
                scale(axis_w, sphere_radius * math.sin(theta)),
            ),
        )

    def residual(theta: float) -> float:
        """功能: 计算反射条件残差；输入: 参数角 theta；输出: 法线与角平分线共线误差。"""
        point = point_at(theta)
        normal = normalize(subtract(point, sphere_center), "反射点法线")
        to_regulation = normalize(subtract(regulation_point, point), "反射点到法规点方向")
        to_eye = normalize(subtract(eye_point, point), "反射点到眼点方向")
        bisector = add(to_regulation, to_eye)
        return dot(cross(normal, bisector), normalize(plane_normal, "反射截面法线"))

    samples = 720
    brackets: list[tuple[float, float]] = []
    previous_theta = 0.0
    previous_value = residual(previous_theta)
    best_theta = previous_theta
    best_value = abs(previous_value)
    for index in range(1, samples + 1):
        theta = 2.0 * math.pi * index / samples
        value = residual(theta)
        if abs(value) < best_value:
            best_theta = theta
            best_value = abs(value)
        if previous_value == 0.0 or value == 0.0 or previous_value * value < 0.0:
            brackets.append((previous_theta, theta))
        previous_theta = theta
        previous_value = value

    candidates: list[Vector] = []
    for low, high in brackets:
        low_value = residual(low)
        for _ in range(60):
            mid = (low + high) / 2.0
            mid_value = residual(mid)
            if low_value * mid_value <= 0.0:
                high = mid
            else:
                low = mid
                low_value = mid_value
        candidates.append(point_at((low + high) / 2.0))

    if not candidates:
        candidates.append(point_at(best_theta))

    eye_to_center = subtract(sphere_center, eye_point)

    def is_between_eye_and_center_side(point: Vector) -> bool:
        """功能: 判断候选点是否位于眼点到球心的可见侧；输入: 候选点；输出: 是否可选。"""
        eye_to_point = subtract(point, eye_point)
        projection = dot(eye_to_point, eye_to_center)
        return 0.0 <= projection <= dot(eye_to_center, eye_to_center)

    def reflection_error(point: Vector) -> float:
        """功能: 计算候选反射点误差；输入: 候选点；输出: 反射角平分条件误差。"""
        to_regulation = normalize(subtract(regulation_point, point), "反射点到法规点方向")
        to_eye = normalize(subtract(eye_point, point), "反射点到眼点方向")
        normal = normalize(subtract(point, sphere_center), "反射点法线")
        return vector_length(cross(normal, add(to_regulation, to_eye)))

    visible_candidates = [
        candidate
        for candidate in candidates
        if is_between_eye_and_center_side(candidate)
    ]
    selectable_candidates = visible_candidates or candidates
    return min(
        selectable_candidates,
        key=lambda point: (
            vector_length(subtract(point, eye_point)),
            reflection_error(point),
        ),
    )


def create_temporary_reflection_points(
    document: Any,
    part: Any,
    regulation_line_hybrid_body: Any,
    parametric_rearview_hybrid_body: Any,
    left_eye: Any,
    right_eye: Any,
    clearance_threshold: float = 3.0,
    angle_limit: float = 2.0,
    angle_step: float = 0.2,
    surface_tolerance: float = 0.5,
) -> dict[str, Any]:
    """功能: 搜索镜片角度并创建满足边界间隙和镜面约束的临时反射点；输入: 文档、Part、法规线集、参数化镜片集、左右眼点、间隙阈值、角度范围/步长和面容差；输出: 临时反射取点结果字典。"""
    mirror_inputs = {
        "左": {
            "rotated_mirror_name": "左镜片旋转",
            "geo_set_name": "左镜片临时法规反射取点",
            "boundary_name": "左镜片临时边界",
            "horizontal_parameter_name": "左镜片水平旋转角度",
            "vertical_parameter_name": "左镜片竖直旋转角度",
        },
        "右": {
            "rotated_mirror_name": "右镜片旋转",
            "geo_set_name": "右镜片临时法规反射取点",
            "boundary_name": "右镜片临时边界",
            "horizontal_parameter_name": "右镜片水平旋转角度",
            "vertical_parameter_name": "右镜片竖直旋转角度",
        },
    }
    eye_inputs = {
        "右": right_eye,
        "左": left_eye,
    }
    eye_coordinates = {
        eye_side_name: get_point(document, part, eye_feature)
        for eye_side_name, eye_feature in eye_inputs.items()
    }
    mirror_results: dict[str, Any] = {}

    def set_angle_parameter(parameter_name: str, degree_value: float) -> None:
        """功能: 设置镜片旋转角度参数；输入: 参数名和角度值；输出: 无。"""
        parameter = part.Parameters.Item(parameter_name)
        parameter.ValuateFromString(f"{degree_value:.1f}deg")

    def generate_angle_pairs() -> list[tuple[float, float]]:
        """功能: 生成按偏转幅度排序的水平/竖直角度组合；输入: 搜索范围参数；输出: 角度组合列表。"""
        step_count = int(round(angle_limit / angle_step))
        values = [round(index * angle_step, 1) for index in range(-step_count, step_count + 1)]
        pairs = [(horizontal, vertical) for horizontal in values for vertical in values]
        return sorted(
            pairs,
            key=lambda pair: (
                max(abs(pair[0]), abs(pair[1])),
                abs(pair[0]) + abs(pair[1]),
                abs(pair[0]),
                abs(pair[1]),
                pair[0],
                pair[1],
            ),
        )

    angle_pairs = generate_angle_pairs()

    def final_view_group_definitions(
        mirror_side_name: str,
    ) -> tuple[tuple[str, tuple[tuple[str, int], ...]], ...]:
        """功能: 按镜片侧生成最终视野点分组；输入: 镜片侧；输出: 分组名和(眼点侧, 点序号)。"""
        fourth_group_eye_side = "左" if mirror_side_name == "左" else "右"
        return (
            ("最终视野点1", (("右", 1), ("左", 1))),
            ("最终视野点2", (("右", 2), ("左", 2))),
            ("最终视野点3", (("右", 5), ("左", 5))),
            (
                "最终视野点4",
                ((fourth_group_eye_side, 3), (fourth_group_eye_side, 6)),
            ),
        )

    def select_final_view_points(
        mirror_side_name: str,
        point_results: list[dict[str, Any]],
    ) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
        """功能: 按需求分组，从每组在镜面上的点中选出距边界最远的最终视野点；输入: 12个临时点结果；输出: 分组结果和最终点列表。"""
        group_results: list[dict[str, Any]] = []
        selected_points: list[dict[str, Any]] = []
        for group_name, point_keys in final_view_group_definitions(mirror_side_name):
            group_candidates = [
                point
                for point in point_results
                if (point.get("eye_side_name"), point.get("point_index")) in point_keys
            ]
            surface_candidates = [
                point
                for point in group_candidates
                if point.get("is_point_on_mirror_surface")
            ]
            selected_point = (
                max(surface_candidates, key=lambda point: point["distance_to_boundary"])
                if surface_candidates
                else None
            )
            selected_distance = (
                selected_point["distance_to_boundary"] if selected_point else None
            )
            group_pass = (
                selected_distance is not None
                and selected_distance >= clearance_threshold
            )
            group_result = {
                "group_name": group_name,
                "candidate_point_names": [
                    str(point["reflection_point_name"]) for point in group_candidates
                ],
                "surface_candidate_point_names": [
                    str(point["reflection_point_name"]) for point in surface_candidates
                ],
                "valid_candidate_point_names": [
                    str(point["reflection_point_name"])
                    for point in surface_candidates
                    if point.get("distance_to_boundary", 0.0) >= clearance_threshold
                ],
                "selected_point_name": (
                    str(selected_point["reflection_point_name"]) if selected_point else None
                ),
                "selected_point_coordinates": (
                    selected_point["reflection_point_coordinates"] if selected_point else None
                ),
                "selected_distance_to_boundary": selected_distance,
                "pass": group_pass,
            }
            group_results.append(group_result)
            if selected_point is not None:
                selected_points.append(
                    {
                        "group_name": group_name,
                        "reflection_point_name": str(selected_point["reflection_point_name"]),
                        "reflection_point_coordinates": selected_point[
                            "reflection_point_coordinates"
                        ],
                        "eye_side_name": selected_point["eye_side_name"],
                        "mirror_side_name": selected_point["mirror_side_name"],
                        "point_index": selected_point["point_index"],
                        "distance_to_boundary": selected_point["distance_to_boundary"],
                        "distance_to_mirror_surface": selected_point[
                            "distance_to_mirror_surface"
                        ],
                    }
                )
        return group_results, selected_points

    def calculate_temporary_local_points(
        point_records: list[dict[str, Any]],
        sphere_center: Vector,
    ) -> dict[str, tuple[float, float]]:
        """功能: 将临时反射点转换到镜片局部二维平面；输入: 点结果和球心；输出: 点名到二维坐标的映射。"""
        coordinates = [
            tuple(point["reflection_point_coordinates"]) for point in point_records
        ]
        average_point = scale(
            tuple(sum(point[index] for point in coordinates) for index in range(3)),
            1.0 / len(coordinates),
        )
        normal = normalize(subtract(average_point, sphere_center), "临时镜片局部法线")
        tangent_candidates = [
            project_to_plane(subtract(point, average_point), normal)
            for point in coordinates
        ]
        tangent_x = max(tangent_candidates, key=vector_length)
        axis_x = normalize(tangent_x, "临时镜片局部横向")
        axis_y = normalize(cross(normal, axis_x), "临时镜片局部纵向")
        return {
            str(record["reflection_point_name"]): (
                dot(subtract(tuple(record["reflection_point_coordinates"]), average_point), axis_x),
                dot(subtract(tuple(record["reflection_point_coordinates"]), average_point), axis_y),
            )
            for record in point_records
        }

    def select_final_view_points_by_minimum_area(
        mirror_side_name: str,
        point_results: list[dict[str, Any]],
        sphere_center: Vector,
        allow_all_points_when_no_surface: bool = False,
    ) -> tuple[list[dict[str, Any]], list[dict[str, Any]], dict[str, Any]]:
        """功能: 按四组候选点遍历四点组合，并选取面积最小凸四边形；输入: 12个临时点、球心和回退策略；输出: 分组结果、最终点和选择摘要。"""
        group_results: list[dict[str, Any]] = []
        candidate_groups: list[list[dict[str, Any]]] = []
        for group_name, point_keys in final_view_group_definitions(mirror_side_name):
            group_candidates = [
                point
                for point in point_results
                if (point.get("eye_side_name"), point.get("point_index")) in point_keys
            ]
            surface_candidates = [
                point
                for point in group_candidates
                if point.get("is_point_on_mirror_surface")
            ]
            valid_candidates = [
                point
                for point in surface_candidates
                if point.get("distance_to_boundary", 0.0) >= clearance_threshold
            ]
            if valid_candidates:
                temporary_candidates = valid_candidates
                candidate_source = "surface_clearance_pass"
            elif surface_candidates:
                temporary_candidates = surface_candidates
                candidate_source = "surface_points"
            elif allow_all_points_when_no_surface:
                temporary_candidates = group_candidates
                candidate_source = "all_points_fallback"
            else:
                temporary_candidates = []
                candidate_source = "no_surface_points"
            candidate_groups.append(temporary_candidates)
            group_results.append(
                {
                    "group_name": group_name,
                    "candidate_point_names": [
                        str(point["reflection_point_name"]) for point in group_candidates
                    ],
                    "surface_candidate_point_names": [
                        str(point["reflection_point_name"]) for point in surface_candidates
                    ],
                    "valid_candidate_point_names": [
                        str(point["reflection_point_name"]) for point in valid_candidates
                    ],
                    "temporary_candidate_point_names": [
                        str(point["reflection_point_name"]) for point in temporary_candidates
                    ],
                    "candidate_source": candidate_source,
                    "has_surface_points": bool(surface_candidates),
                    "has_clearance_pass_points": bool(valid_candidates),
                    "pass": False,
                }
            )

        if any(not candidates for candidates in candidate_groups):
            return (
                group_results,
                [],
                {
                    "selection_success": False,
                    "failure_reason": "存在分组无法生成临时视野候选点。",
                    "combination_count": 0,
                    "convex_combination_count": 0,
                },
            )

        local_points = calculate_temporary_local_points(point_results, sphere_center)
        best_records: list[dict[str, Any]] | None = None
        best_local_points: list[tuple[float, float]] = []
        best_area: float | None = None
        best_min_distance = -1.0
        combination_count = 0
        convex_combination_count = 0
        for combination_records in product(*candidate_groups):
            combination_count += 1
            selected_local_points = [
                local_points[str(point["reflection_point_name"])]
                for point in combination_records
            ]
            order = order_points_around_center(selected_local_points)
            ordered_local_points = [selected_local_points[index] for index in order]
            if not is_convex_quadrilateral(ordered_local_points):
                continue
            convex_combination_count += 1
            ordered_records = [combination_records[index] for index in order]
            area = polygon_area_2d(ordered_local_points)
            min_distance = min(
                float(point["distance_to_boundary"]) for point in ordered_records
            )
            if (
                best_area is None
                or area < best_area
                or (abs(area - best_area) <= 1e-9 and min_distance > best_min_distance)
            ):
                best_area = area
                best_min_distance = min_distance
                best_records = list(ordered_records)
                best_local_points = ordered_local_points

        if best_records is None or best_area is None:
            return (
                group_results,
                [],
                {
                    "selection_success": False,
                    "failure_reason": "无法从临时视野候选点形成凸四边形。",
                    "combination_count": combination_count,
                    "convex_combination_count": convex_combination_count,
                },
            )

        selected_points: list[dict[str, Any]] = []
        for selected_point in best_records:
            selected_group_name = next(
                (
                    group_name
                    for group_name, point_keys in final_view_group_definitions(
                        mirror_side_name
                    )
                    if (
                        selected_point.get("eye_side_name"),
                        selected_point.get("point_index"),
                    )
                    in point_keys
                ),
                "",
            )
            selected_points.append(
                {
                    "group_name": selected_group_name,
                    "reflection_point_name": str(selected_point["reflection_point_name"]),
                    "reflection_point_coordinates": selected_point[
                        "reflection_point_coordinates"
                    ],
                    "eye_side_name": selected_point["eye_side_name"],
                    "mirror_side_name": selected_point["mirror_side_name"],
                    "point_index": selected_point["point_index"],
                    "distance_to_boundary": selected_point["distance_to_boundary"],
                    "distance_to_mirror_surface": selected_point[
                        "distance_to_mirror_surface"
                    ],
                }
            )
            for group_result in group_results:
                if group_result["group_name"] == selected_group_name:
                    group_result.update(
                        {
                            "selected_point_name": str(
                                selected_point["reflection_point_name"]
                            ),
                            "selected_point_coordinates": selected_point[
                                "reflection_point_coordinates"
                            ],
                            "selected_distance_to_boundary": selected_point[
                                "distance_to_boundary"
                            ],
                            "pass": (
                                selected_point["distance_to_boundary"]
                                >= clearance_threshold
                            ),
                        }
                    )
                    break

        return (
            group_results,
            selected_points,
            {
                "selection_success": True,
                "selection_method": "minimum_area_convex_quadrilateral",
                "quadrilateral_area": round(best_area, 2),
                "min_distance_to_boundary": round(best_min_distance, 2),
                "combination_count": combination_count,
                "convex_combination_count": convex_combination_count,
                "local_coordinates": [
                    [round(value, 2) for value in point] for point in best_local_points
                ],
            },
        )

    def print_temporary_candidate_result(candidate_result: dict[str, Any]) -> None:
        """功能: 输出当前临时反射取点候选角度简要进度；输入: 候选结果；输出: 无。"""
        print(
            f"[临时法规反射取点] {candidate_result.get('rotated_mirror_name')} "
            f"搜索序号={candidate_result.get('searched_angle_count')}，"
            f"水平角={candidate_result.get('horizontal_angle')} deg，"
            f"竖直角={candidate_result.get('vertical_angle')} deg，"
            f"最终视野点组通过={candidate_result.get('all_final_view_groups_pass')}",
            flush=True,
        )

    def print_temporary_final_result(mirror_result: dict[str, Any]) -> None:
        """功能: 在单侧镜片搜索结束后输出最终角度和最终视野点；输入: 单侧临时测量结果；输出: 无。"""
        status = "通过" if mirror_result.get("calibration_success") else "失败"
        fallback_text = (
            "，已回退到0度姿态创建临时反射取点"
            if mirror_result.get("fallback_to_initial_angle")
            else ""
        )
        print(
            "\n"
            f"[临时法规反射取点结果] {mirror_result.get('rotated_mirror_name')}\n"
            f"校核结果={status}{fallback_text}，"
            f"水平角={mirror_result.get('horizontal_angle')} deg，"
            f"竖直角={mirror_result.get('vertical_angle')} deg，"
            f"最终视野点数量={mirror_result.get('final_view_point_count', 0)}",
            flush=True,
        )
        best_failed_result = mirror_result.get("best_failed_result")
        if best_failed_result:
            print(
                "最佳失败结果: "
                f"水平角={best_failed_result.get('horizontal_angle')} deg，"
                f"竖直角={best_failed_result.get('vertical_angle')} deg，"
                f"最终视野点数量={best_failed_result.get('final_view_point_count', 0)}，"
                f"最小最终视野点边界距离="
                f"{best_failed_result.get('min_final_view_point_distance_to_boundary')} mm",
                flush=True,
            )
        for final_point in mirror_result.get("final_view_points", []):
            print(
                f"{final_point.get('group_name')}: "
                f"{final_point.get('reflection_point_name')}，"
                f"坐标={final_point.get('reflection_point_coordinates')}，"
                f"到边界距离={final_point.get('distance_to_boundary')} mm",
                flush=True,
            )
        final_view_selection = mirror_result.get("final_view_selection", {})
        if final_view_selection.get("selection_success"):
            print(
                f"最终视野四边形面积={final_view_selection.get('quadrilateral_area')}，"
                f"凸四边形组合数={final_view_selection.get('convex_combination_count')}",
                flush=True,
            )
        elif final_view_selection:
            print(
                f"最终视野点选择失败: {final_view_selection.get('failure_reason')}",
                flush=True,
            )

    for mirror_side_name, mirror_input in mirror_inputs.items():
        rotated_mirror = find_feature(
            parametric_rearview_hybrid_body,
            mirror_input["rotated_mirror_name"],
        )
        best_failed_result: dict[str, Any] | None = None
        accepted_result: dict[str, Any] | None = None
        found_clearance_pass_off_surface = False

        def create_candidate_result(
            horizontal_angle: float,
            vertical_angle: float,
            angle_search_index: int,
            allow_all_points_when_no_surface: bool = False,
        ) -> tuple[dict[str, Any], Any]:
            """功能: 创建并测量指定角度的临时反射点；输入: 水平/竖直角和序号；输出: 结果字典和临时几何集。"""
            set_angle_parameter(mirror_input["horizontal_parameter_name"], horizontal_angle)
            set_angle_parameter(mirror_input["vertical_parameter_name"], vertical_angle)
            part.Update()

            temporary_hybrid_body = None
            try:
                temporary_hybrid_body = create_hybrid_body(part, mirror_input["geo_set_name"])
                sphere_center_feature = create_point_center(
                    part,
                    temporary_hybrid_body,
                    rotated_mirror,
                    f"{mirror_side_name}镜片临时球心",
                )
                sphere_center = get_point(document, part, sphere_center_feature)
                sphere_radius = get_minimum_distance(
                    document,
                    part,
                    sphere_center_feature,
                    rotated_mirror,
                )
                boundary = create_boundary(
                    part,
                    temporary_hybrid_body,
                    rotated_mirror,
                    mirror_input["boundary_name"],
                )

                point_results: list[dict[str, Any]] = []
                for eye_side_name, eye_point in eye_coordinates.items():
                    for point_index in range(1, 7):
                        regulation_point_name = f"{mirror_side_name}法规点{point_index}"
                        regulation_point_feature = find_feature(
                            regulation_line_hybrid_body,
                            regulation_point_name,
                        )
                        regulation_point = get_point(document, part, regulation_point_feature)
                        reflection_coordinates = solve_spherical_reflection_point(
                            sphere_center,
                            sphere_radius,
                            regulation_point,
                            eye_point,
                        )
                        reflection_point_name = (
                            f"{eye_side_name}眼{mirror_side_name}法规反射取点{point_index}"
                        )
                        reflection_point = create_coordinate_point(
                            part,
                            temporary_hybrid_body,
                            reflection_coordinates,
                            reflection_point_name,
                        )
                        distance_to_boundary = get_minimum_distance(
                            document,
                            part,
                            reflection_point,
                            boundary,
                        )
                        distance_to_mirror_surface = get_minimum_distance(
                            document,
                            part,
                            reflection_point,
                            rotated_mirror,
                        )
                        is_point_on_mirror_surface = distance_to_mirror_surface <= surface_tolerance
                        point_results.append(
                            {
                                "reflection_point_name": str(reflection_point.Name),
                                "reflection_point_coordinates": round_vector(reflection_coordinates),
                                "eye_side_name": eye_side_name,
                                "mirror_side_name": mirror_side_name,
                                "point_index": point_index,
                                "regulation_point_name": regulation_point_name,
                                "distance_to_boundary": round(distance_to_boundary, 2),
                                "distance_to_mirror_surface": round(distance_to_mirror_surface, 3),
                                "is_point_on_mirror_surface": is_point_on_mirror_surface,
                                "is_clearance_greater_equal_3mm": (
                                    distance_to_boundary >= clearance_threshold
                                ),
                            }
                        )

                min_distance = min(
                    point_result["distance_to_boundary"]
                    for point_result in point_results
                )
                max_surface_distance = max(
                    point_result["distance_to_mirror_surface"]
                    for point_result in point_results
                )
                all_points_on_surface = all(
                    point_result["is_point_on_mirror_surface"]
                    for point_result in point_results
                )
                all_clearance_pass = min_distance >= clearance_threshold
                final_view_groups, final_view_points, final_view_selection = (
                    select_final_view_points_by_minimum_area(
                        mirror_side_name,
                        point_results,
                        sphere_center,
                        allow_all_points_when_no_surface,
                    )
                )
                all_final_view_groups_pass = all(
                    group_result["pass"] for group_result in final_view_groups
                ) and len(final_view_points) == 4
                all_points_pass = all_final_view_groups_pass
                selected_distances = [
                    point["distance_to_boundary"] for point in final_view_points
                ]
                return (
                    {
                        "geo_set_name": str(temporary_hybrid_body.Name),
                        "rotated_mirror_name": str(rotated_mirror.Name),
                        "boundary_name": str(boundary.Name),
                        "sphere_center_name": str(sphere_center_feature.Name),
                        "sphere_center_coordinates": round_vector(sphere_center),
                        "sphere_radius": round(sphere_radius, 2),
                        "horizontal_angle": horizontal_angle,
                        "vertical_angle": vertical_angle,
                        "searched_angle_count": angle_search_index,
                        "min_distance_to_boundary": round(min_distance, 2),
                        "max_distance_to_mirror_surface": round(max_surface_distance, 3),
                        "surface_tolerance": surface_tolerance,
                        "all_clearance_pass": all_clearance_pass,
                        "all_points_on_surface": all_points_on_surface,
                        "all_final_view_groups_pass": all_final_view_groups_pass,
                        "all_points_pass": all_points_pass,
                        "calibration_success": all_points_pass,
                        "reflection_point_count": len(point_results),
                        "reflection_points": point_results,
                        "final_view_group_count": len(final_view_groups),
                        "final_view_point_count": len(final_view_points),
                        "min_final_view_point_distance_to_boundary": (
                            round(min(selected_distances), 2) if selected_distances else None
                        ),
                        "final_view_groups": final_view_groups,
                        "final_view_points": final_view_points,
                        "final_view_selection": final_view_selection,
                    },
                    temporary_hybrid_body,
                )
            except Exception:
                if temporary_hybrid_body is not None:
                    delete_object(document, temporary_hybrid_body)
                raise

        for angle_search_index, (horizontal_angle, vertical_angle) in enumerate(
            angle_pairs,
            start=1,
        ):
            temporary_hybrid_body = None
            try:
                candidate_result, temporary_hybrid_body = create_candidate_result(
                    horizontal_angle,
                    vertical_angle,
                    angle_search_index,
                )
                print_temporary_candidate_result(candidate_result)

                if (
                    best_failed_result is None
                    or (
                        candidate_result["final_view_point_count"],
                        candidate_result.get("min_final_view_point_distance_to_boundary") or -1.0,
                        candidate_result["min_distance_to_boundary"],
                    )
                    > (
                        best_failed_result.get("final_view_point_count", 0),
                        best_failed_result.get("min_final_view_point_distance_to_boundary") or -1.0,
                        best_failed_result.get("min_distance_to_boundary", -1.0),
                    )
                ):
                    best_failed_result = {
                        key: value
                        for key, value in candidate_result.items()
                        if key not in {"reflection_points"}
                    }

                if (
                    candidate_result["all_clearance_pass"]
                    and not candidate_result["all_points_on_surface"]
                ):
                    found_clearance_pass_off_surface = True

                if candidate_result["all_points_pass"]:
                    accepted_result = candidate_result
                    break

                delete_object(document, temporary_hybrid_body)
            except Exception:
                if temporary_hybrid_body is not None:
                    delete_object(document, temporary_hybrid_body)
                raise

        if accepted_result is None:
            print(
                f"[临时法规反射取点] {rotated_mirror.Name} 搜索失败，"
                "将回退到0度姿态创建临时反射取点。",
                flush=True,
            )
            try:
                fallback_result, _fallback_body = create_candidate_result(
                    0.0,
                    0.0,
                    1,
                    allow_all_points_when_no_surface=True,
                )
                fallback_result.update(
                    {
                        "all_points_pass": False,
                        "calibration_success": False,
                        "fallback_to_initial_angle": True,
                        "failure_reason": (
                            "未找到同时满足边界距离和点在镜面上的偏转角度，"
                            "已回退到初始0度姿态创建反射取点；镜片过小时允许反射点位于镜片外。"
                        ),
                        "found_clearance_pass_off_surface": found_clearance_pass_off_surface,
                        "best_failed_result": best_failed_result,
                    }
                )
                mirror_results[mirror_side_name] = fallback_result
                print_temporary_final_result(fallback_result)
            except Exception as exc:
                set_angle_parameter(mirror_input["horizontal_parameter_name"], 0.0)
                set_angle_parameter(mirror_input["vertical_parameter_name"], 0.0)
                part.Update()
                failed_result = {
                    "geo_set_name": None,
                    "rotated_mirror_name": str(rotated_mirror.Name),
                    "horizontal_angle": 0.0,
                    "vertical_angle": 0.0,
                    "all_points_pass": False,
                    "calibration_success": False,
                    "reflection_point_count": 0,
                    "reflection_points": [],
                    "fallback_to_initial_angle": True,
                    "failure_reason": f"已回退到0度，但0度临时反射取点创建失败: {exc}",
                    "best_failed_result": best_failed_result,
                }
                mirror_results[mirror_side_name] = failed_result
                print_temporary_final_result(failed_result)
        else:
            mirror_results[mirror_side_name] = accepted_result
            print_temporary_final_result(accepted_result)

    return {
        "total_reflection_point_count": sum(
            mirror_result["reflection_point_count"]
            for mirror_result in mirror_results.values()
        ),
        "clearance_threshold": clearance_threshold,
        "surface_tolerance": surface_tolerance,
        "angle_limit": angle_limit,
        "angle_step": angle_step,
        "mirrors": mirror_results,
    }


def create_regulation_projection_references(
    document: Any,
    part: Any,
    regulation_line_hybrid_body: Any,
    eye_to_ground_projection: Any,
    left_vehicle_width_line: Any,
    right_vehicle_width_line: Any,
    up_direction: Vector,
    rough_rear_direction: Vector,
) -> dict[str, Any]:
    """功能: 创建车宽线投影和左右车宽线延长线，并据此修正车辆方向；输入: 文档、Part、法规线集、眼点投影、左右车宽线、上方向和粗略后方向；输出: 投影、延长线和方向。"""
    left_projection = create_projection(
        part,
        regulation_line_hybrid_body,
        eye_to_ground_projection,
        left_vehicle_width_line,
        "地面点到左车宽线投影点",
    )
    right_projection = create_projection(
        part,
        regulation_line_hybrid_body,
        eye_to_ground_projection,
        right_vehicle_width_line,
        "地面点到右车宽线投影点",
    )
    left_projection_coordinates = get_point(document, part, left_projection)
    right_projection_coordinates = get_point(document, part, right_projection)
    right_direction = normalize(
        project_to_plane(
            subtract(right_projection_coordinates, left_projection_coordinates),
            up_direction,
        ),
        "车辆右方向",
    )
    left_direction = scale(right_direction, -1.0)
    rough_rear_direction = normalize(
        project_to_plane(rough_rear_direction, up_direction),
        "车辆粗略后方向",
    )

    def create_extension(
        side_name: str,
        vehicle_width_line: Any,
        projection: Any,
        projection_coordinates: Vector,
        reference_rear_direction: Vector,
    ) -> dict[str, Any]:
        selected_endpoint = None
        selected_coordinates: Vector | None = None
        for orientation in (False, True):
            candidate_name = f"{side_name}延长线端点"
            try:
                candidate = create_point_on_curve_with_reference_distance(
                    part,
                    regulation_line_hybrid_body,
                    vehicle_width_line,
                    projection,
                    VEHICLE_WIDTH_LINE_EXTENSION_DISTANCE,
                    orientation,
                    candidate_name,
                )
                candidate_coordinates = get_point(document, part, candidate)
                candidate_direction = normalize(
                    project_to_plane(
                        subtract(candidate_coordinates, projection_coordinates),
                        up_direction,
                    ),
                    f"{side_name}车宽线延长方向",
                )
                if dot(candidate_direction, reference_rear_direction) >= 0:
                    selected_endpoint = candidate
                    selected_coordinates = candidate_coordinates
                    break
                delete_object(document, candidate)
            except Exception:
                pass

        if selected_endpoint is None or selected_coordinates is None:
            selected_coordinates = add(
                projection_coordinates,
                scale(reference_rear_direction, VEHICLE_WIDTH_LINE_EXTENSION_DISTANCE),
            )
            selected_endpoint = create_coordinate_point(
                part,
                regulation_line_hybrid_body,
                selected_coordinates,
                f"{side_name}延长线端点",
            )

        try:
            selected_endpoint.Name = f"{side_name}延长线端点"
        except Exception:
            pass
        extension_line = create_point_to_point_line(
            part,
            regulation_line_hybrid_body,
            projection,
            selected_endpoint,
            f"{side_name}车宽线延长线",
        )
        extension_direction = normalize(
            project_to_plane(
                subtract(selected_coordinates, projection_coordinates),
                up_direction,
            ),
            f"{side_name}车宽线延长方向",
        )
        return {
            "endpoint": selected_endpoint,
            "endpoint_coordinates": selected_coordinates,
            "line": extension_line,
            "direction": extension_direction,
            "summary": {
                "endpoint_name": str(selected_endpoint.Name),
                "endpoint_coordinates": round_vector(selected_coordinates),
                "line_name": str(extension_line.Name),
                "line_length": round(distance(projection_coordinates, selected_coordinates), 2),
            },
        }

    left_extension = create_extension(
        "左",
        left_vehicle_width_line,
        left_projection,
        left_projection_coordinates,
        rough_rear_direction,
    )
    rear_direction = normalize(
        project_to_plane(left_extension["direction"], up_direction),
        "车辆后方向",
    )
    front_direction = scale(rear_direction, -1.0)
    right_extension = create_extension(
        "右",
        right_vehicle_width_line,
        right_projection,
        right_projection_coordinates,
        rear_direction,
    )

    return {
        "left_projection": left_projection,
        "right_projection": right_projection,
        "left_extension_endpoint": left_extension["endpoint"],
        "left_extension_line": left_extension["line"],
        "right_extension_endpoint": right_extension["endpoint"],
        "right_extension_line": right_extension["line"],
        "directions": {
            "车辆后方向": rear_direction,
            "车辆前方向": front_direction,
            "车辆左方向": left_direction,
            "车辆右方向": right_direction,
        },
        "left_extension": left_extension["summary"],
        "right_extension": right_extension["summary"],
    }


def create_regulation_line(
    document: Any,
    part: Any,
    regulation_line_hybrid_body: Any,
    mirror: Any,
    eye: Any | None,
    ground: Any,
    vehicle_width_line: Any,
    up_direction: Vector,
    rear_direction: Vector,
    side_direction: Vector,
    side_name: str,
    eye_to_ground_projection: Any | None = None,
    ground_to_width_projection: Any | None = None,
    rear_reference_line: Any | None = None,
) -> dict[str, Any]:
    """功能: 创建一侧法规点和法规线；输入: 文档、Part、法规线集、镜片、眼点/投影、地面、车宽线、车辆方向和侧别；输出: 法规线结果字典。"""
    if side_name not in ("左", "右"):
        raise ValueError(f"不支持的法规线侧别: {side_name}")

    mirror_extremum_name = "镜片极值点" if side_name == "左" else "右镜片极值点"
    eye_to_ground_projection_name = "眼点到地面投影点"
    mirror_extremum = create_extremum_point(
        part,
        regulation_line_hybrid_body,
        mirror,
        up_direction,
        mirror_extremum_name,
    )
    mirror_extremum_to_ground_distance = get_minimum_distance(
        document,
        part,
        mirror_extremum,
        ground,
    )
    try:
        mirror_extremum_coordinates: Vector | None = get_point(document, part, mirror_extremum)
    except RuntimeError:
        mirror_extremum_coordinates = None
    hide_object(document, mirror_extremum)

    if eye_to_ground_projection is None:
        if eye is None:
            raise ValueError("缺少眼点，无法创建眼点到地面投影点。")
        eye_to_ground_projection = create_projection(
            part,
            regulation_line_hybrid_body,
            eye,
            ground,
            eye_to_ground_projection_name,
        )
    if ground_to_width_projection is None:
        ground_to_width_projection = create_projection(
            part,
            regulation_line_hybrid_body,
            eye_to_ground_projection,
            vehicle_width_line,
            f"地面点到{side_name}车宽线投影点",
        )
    base_point = get_point(document, part, ground_to_width_projection)

    points: dict[str, Any] = {}
    point_coordinates: dict[str, Vector] = {}
    try:
        if rear_reference_line is None:
            raise RuntimeError("缺少车宽线延长线")
        point1 = create_point_on_curve_with_reference_distance(
            part,
            regulation_line_hybrid_body,
            rear_reference_line,
            ground_to_width_projection,
            4000.0,
            False,
            f"{side_name}法规点1",
        )
        point4 = create_point_on_curve_with_reference_distance(
            part,
            regulation_line_hybrid_body,
            rear_reference_line,
            ground_to_width_projection,
            20000.0,
            False,
            f"{side_name}法规点4",
        )
        point1_coordinates = get_point(document, part, point1)
        point4_coordinates = get_point(document, part, point4)
        if dot(subtract(point1_coordinates, base_point), rear_direction) < 0:
            delete_object(document, point1)
            delete_object(document, point4)
            point1 = create_point_on_curve_with_reference_distance(
                part,
                regulation_line_hybrid_body,
                rear_reference_line,
                ground_to_width_projection,
                4000.0,
                True,
                f"{side_name}法规点1",
            )
            point4 = create_point_on_curve_with_reference_distance(
                part,
                regulation_line_hybrid_body,
                rear_reference_line,
                ground_to_width_projection,
                20000.0,
                True,
                f"{side_name}法规点4",
            )
            point1_coordinates = get_point(document, part, point1)
            point4_coordinates = get_point(document, part, point4)
        points[f"{side_name}法规点1"] = point1
        points[f"{side_name}法规点4"] = point4
        point_coordinates[f"{side_name}法规点1"] = point1_coordinates
        point_coordinates[f"{side_name}法规点4"] = point4_coordinates
    except Exception:
        point_coordinates[f"{side_name}法规点1"] = add(base_point, scale(rear_direction, 4000.0))
        point_coordinates[f"{side_name}法规点4"] = add(base_point, scale(rear_direction, 20000.0))

    point_coordinates[f"{side_name}法规点2"] = add(
        point_coordinates[f"{side_name}法规点1"],
        scale(side_direction, 1000.0),
    )
    point_coordinates[f"{side_name}法规点3"] = add(
        point_coordinates[f"{side_name}法规点1"],
        scale(up_direction, mirror_extremum_to_ground_distance),
    )
    # 点 5 与点 4 对应，形成远端法规边界。
    point_coordinates[f"{side_name}法规点5"] = add(
        point_coordinates[f"{side_name}法规点4"],
        scale(side_direction, 4000.0),
    )
    point_coordinates[f"{side_name}法规点6"] = add(
        point_coordinates[f"{side_name}法规点4"],
        scale(up_direction, mirror_extremum_to_ground_distance),
    )

    for name, coordinates in point_coordinates.items():
        if name not in points:
            points[name] = create_coordinate_point(
                part,
                regulation_line_hybrid_body,
                coordinates,
                name,
            )
    line_point_names = (
        (f"{side_name}法规线1", f"{side_name}法规点1", f"{side_name}法规点2"),
        (f"{side_name}法规线2", f"{side_name}法规点1", f"{side_name}法规点3"),
        (f"{side_name}法规线3", f"{side_name}法规点4", f"{side_name}法规点5"),
        (f"{side_name}法规线4", f"{side_name}法规点4", f"{side_name}法规点6"),
        (f"{side_name}法规线5", f"{side_name}法规点1", f"{side_name}法规点4"),
        (f"{side_name}法规线6", f"{side_name}法规点2", f"{side_name}法规点5"),
    )
    lines = {
        line_name: create_point_to_point_line(
            part,
            regulation_line_hybrid_body,
            points[start_point_name],
            points[end_point_name],
            line_name,
        )
        for line_name, start_point_name, end_point_name in line_point_names
    }

    return {
        "mirror_extremum_point_name": str(mirror_extremum.Name),
        "mirror_extremum_point_hidden": True,
        "mirror_extremum_point_coordinates": (
            round_vector(mirror_extremum_coordinates)
            if mirror_extremum_coordinates is not None
            else None
        ),
        "mirror_extremum_to_ground_distance": round(mirror_extremum_to_ground_distance, 2),
        "projection_point_names": [
            str(eye_to_ground_projection.Name),
            str(ground_to_width_projection.Name),
        ],
        "projection_point_coordinates": {
            str(eye_to_ground_projection.Name): round_vector(
                get_point(document, part, eye_to_ground_projection)
            ),
            str(ground_to_width_projection.Name): round_vector(base_point),
        },
        "point_coordinates": {
            name: round_vector(coordinates)
            for name, coordinates in point_coordinates.items()
        },
        "line_point_names": {
            line_name: [start_point_name, end_point_name]
            for line_name, start_point_name, end_point_name in line_point_names
        },
        "line_names": list(lines),
    }


def distance_between_points(first: Vector, second: Vector) -> float:
    """功能: 计算两点间距；输入: 两个三维坐标；输出: 距离值。"""
    return vector_length(subtract(first, second))


def is_close_distance(value: float, target: float, tolerance: float = 1.0) -> bool:
    """功能: 判断距离是否在容差内接近目标值；输入: 实测值、目标值、容差；输出: 是否接近。"""
    return abs(value - target) <= tolerance


def is_parallel_to_direction(
    start_point: Vector,
    end_point: Vector,
    direction: Vector,
    tolerance: float = 1e-3,
) -> bool:
    """功能: 判断线段方向是否与参考方向平行或反向；输入: 线段端点、参考方向、容差；输出: 是否平行。"""
    line_direction = normalize(subtract(end_point, start_point), "法规检查线方向")
    reference_direction = normalize(direction, "法规检查参考方向")
    return vector_length(cross(line_direction, reference_direction)) <= tolerance


def check_regulation_vision_requirements(
    document: Any,
    part: Any,
    regulation_line_hybrid_body: Any,
    left_regulation_line: dict[str, Any],
    right_regulation_line: dict[str, Any],
    rear_direction: Vector,
    up_direction: Vector,
) -> dict[str, Any]:
    """功能: 检查已创建法规线是否满足法规视野构造要求；输入: 文档、Part、法规线集、左右法规线结果、车辆后方向和上方向；输出: 检查结果字典。"""
    line_results = {
        "左": left_regulation_line,
        "右": right_regulation_line,
    }

    def point(side: str, name: str) -> Vector:
        """功能: 读取指定侧法规点坐标；输入: 侧别和点名；输出: 三维坐标。"""
        coordinates = line_results[side]["point_coordinates"][name]
        return tuple(float(value) for value in coordinates)  # type: ignore[return-value]

    def line_length(side: str, line_name: str) -> float:
        """功能: 计算指定法规线长度；输入: 侧别和线名；输出: 长度。"""
        start_name, end_name = line_results[side]["line_point_names"][line_name]
        return distance_between_points(point(side, start_name), point(side, end_name))

    def line_parallel_up(side: str, line_name: str) -> bool:
        """功能: 判断法规线是否平行地面法线；输入: 侧别和线名；输出: 是否平行。"""
        start_name, end_name = line_results[side]["line_point_names"][line_name]
        return is_parallel_to_direction(point(side, start_name), point(side, end_name), up_direction)

    def line_parallel_rear(side: str, line_name: str) -> bool:
        """功能: 判断法规线是否平行车辆后方向；输入: 侧别和线名；输出: 是否平行。"""
        start_name, end_name = line_results[side]["line_point_names"][line_name]
        return is_parallel_to_direction(point(side, start_name), point(side, end_name), rear_direction)

    def line_parallel_vehicle_width(side: str, line_name: str) -> bool:
        """功能: 判断法规线是否平行对应车宽线延长线；输入: 侧别和线名；输出: 是否平行。"""
        start_name, end_name = line_results[side]["line_point_names"][line_name]
        projection_name = f"地面点到{side}车宽线投影点"
        extension_key = "left_extension" if side == "左" else "right_extension"
        extension = line_results[side].get(extension_key, {})
        projection_coordinates = tuple(
            float(value)
            for value in line_results[side]["projection_point_coordinates"][projection_name]
        )
        endpoint_coordinates = tuple(
            float(value)
            for value in extension.get("endpoint_coordinates", [])
        )
        return is_parallel_to_direction(
            point(side, start_name),
            point(side, end_name),
            subtract(endpoint_coordinates, projection_coordinates),
        )

    def projection_distance_to_line(projection_name: str, line_name: str) -> float:
        """功能: 测量投影点到法规线的最小距离；输入: 投影点名和法规线名；输出: 距离。"""
        return get_minimum_distance(
            document,
            part,
            find_feature(regulation_line_hybrid_body, projection_name),
            find_feature(regulation_line_hybrid_body, line_name),
        )

    left_line3_length = line_length("左", "左法规线3")
    right_line3_length = line_length("右", "右法规线3")
    left_line1_length = line_length("左", "左法规线1")
    right_line1_length = line_length("右", "右法规线1")
    left_line5_parallel_rear = line_parallel_rear("左", "左法规线5")
    right_line5_parallel_rear = line_parallel_rear("右", "右法规线5")
    left_line5_parallel_vehicle_width = line_parallel_vehicle_width("左", "左法规线5")
    right_line5_parallel_vehicle_width = line_parallel_vehicle_width("右", "右法规线5")
    left_projection_name = "地面点到左车宽线投影点"
    right_projection_name = "地面点到右车宽线投影点"
    left_projection_to_left_line3 = projection_distance_to_line(left_projection_name, "左法规线3")
    right_projection_to_right_line3 = projection_distance_to_line(right_projection_name, "右法规线3")
    left_projection_to_left_line1 = projection_distance_to_line(left_projection_name, "左法规线1")
    right_projection_to_right_line1 = projection_distance_to_line(right_projection_name, "右法规线1")

    checks = [
        {
            "item_no": "1.1",
            "category": "检查项 1：4米宽视野区域",
            "content": "驾驶员应能看到一个宽度 >= 4 m 的水平道路区域。",
            "basis": "左法规线3和右法规线3长度是否为4000mm。",
            "measure": (
                f"左法规线3={left_line3_length:.2f}mm；"
                f"右法规线3={right_line3_length:.2f}mm"
            ),
            "pass": is_close_distance(left_line3_length, 4000.0)
            and is_close_distance(right_line3_length, 4000.0),
        },
        {
            "item_no": "1.2",
            "category": "检查项 1：4米宽视野区域",
            "content": "该区域的内侧边界由平行于车辆垂直纵向中间平面、且通过驾驶员侧车身最外点的平面所界定。",
            "basis": "左法规线5与左车宽线平行；左法规线2和左法规线4与地面法线方向平行。",
            "measure": (
                f"左法规线5平行左车宽线={left_line5_parallel_vehicle_width}；"
                f"左法规线2平行={line_parallel_up('左', '左法规线2')}；"
                f"左法规线4平行={line_parallel_up('左', '左法规线4')}"
            ),
            "pass": left_line5_parallel_vehicle_width
            and line_parallel_up("左", "左法规线2")
            and line_parallel_up("左", "左法规线4"),
        },
        {
            "item_no": "1.3",
            "category": "检查项 1：4米宽视野区域",
            "content": "该区域的纵向范围：从驾驶员眼点后方20m处开始，一直延伸至地平线。",
            "basis": "地面点到左车宽线投影点到左法规线3、地面点到右车宽线投影点到右法规线3距离是否均为20000mm。",
            "measure": (
                f"左投影点到左法规线3={left_projection_to_left_line3:.2f}mm；"
                f"右投影点到右法规线3={right_projection_to_right_line3:.2f}mm"
            ),
            "pass": is_close_distance(left_projection_to_left_line3, 20000.0)
            and is_close_distance(right_projection_to_right_line3, 20000.0),
        },
        {
            "item_no": "2.1",
            "category": "检查项 2：1米宽视野区域",
            "content": "驾驶员应能看到一个宽度 > 1 m 的路面区域。",
            "basis": "左法规线1和右法规线1长度是否为1000mm。",
            "measure": (
                f"左法规线1={left_line1_length:.2f}mm；"
                f"右法规线1={right_line1_length:.2f}mm"
            ),
            "pass": is_close_distance(left_line1_length, 1000.0)
            and is_close_distance(right_line1_length, 1000.0),
        },
        {
            "item_no": "2.2",
            "category": "检查项 2：1米宽视野区域",
            "content": "该区域的内侧边界由平行于车辆垂直纵向中间平面、且通过车辆最外点的平面所界定。",
            "basis": "左法规线5与左车宽线平行，右法规线5与右车宽线平行；左/右法规线2和4与地面法线方向平行。",
            "measure": (
                f"左法规线5平行左车宽线={left_line5_parallel_vehicle_width}；"
                f"右法规线5平行右车宽线={right_line5_parallel_vehicle_width}；"
                f"左2/4平行={line_parallel_up('左', '左法规线2') and line_parallel_up('左', '左法规线4')}；"
                f"右2/4平行={line_parallel_up('右', '右法规线2') and line_parallel_up('右', '右法规线4')}"
            ),
            "pass": left_line5_parallel_vehicle_width
            and right_line5_parallel_vehicle_width
            and line_parallel_up("左", "左法规线2")
            and line_parallel_up("左", "左法规线4")
            and line_parallel_up("右", "右法规线2")
            and line_parallel_up("右", "右法规线4"),
        },
        {
            "item_no": "2.3",
            "category": "检查项 2：1米宽视野区域",
            "content": "该区域的纵向范围：从通过驾驶员两眼点的垂面后方4m处开始延伸。",
            "basis": "地面点到左车宽线投影点到左法规线1、地面点到右车宽线投影点到右法规线1距离是否均为4000mm。",
            "measure": (
                f"左投影点到左法规线1={left_projection_to_left_line1:.2f}mm；"
                f"右投影点到右法规线1={right_projection_to_right_line1:.2f}mm"
            ),
            "pass": is_close_distance(left_projection_to_left_line1, 4000.0)
            and is_close_distance(right_projection_to_right_line1, 4000.0),
        },
    ]

    return {
        "all_pass": all(check["pass"] for check in checks),
        "checks": checks,
    }


def create_sketch_on_plane(part: Any, hybrid_body: Any, plane_feature: Any, name: str) -> Any:
    """功能: 在指定平面上创建草图；输入: Part、几何图形集、平面特征和草图名；输出: 草图对象。"""
    sketch = hybrid_body.HybridSketches.Add(create_reference(part, plane_feature))
    sketch.Name = name
    part.InWorkObject = sketch
    part.Update()
    return sketch


def set_object_name(obj: Any, name: str) -> Any:
    """功能: 尽力为 CATIA 对象命名；输入: 对象和名称；输出: 原对象。"""
    for attr_name in ("Name", "name"):
        try:
            setattr(obj, attr_name, name)
            return obj
        except Exception:
            pass
    return obj


def get_first_item(collection: Any) -> Any:
    """功能: 从 CATIA 返回集合中读取第一个可用元素；输入: 集合；输出: 第一个元素。"""
    try:
        return collection.Item(1)
    except Exception:
        pass
    for item_name in ("标记.1", "Mark.1", "交点.1", "Intersection.1"):
        try:
            return collection.Item(item_name)
        except Exception:
            pass
    raise LookupError("无法读取草图生成的第一个几何元素。")


def create_sketch_projection(part: Any, factory2d: Any, source_feature: Any, name: str) -> Any:
    """功能: 将 3D 特征投影到当前草图并命名；输入: Part、2D工厂、源特征、名称；输出: 草图元素。"""
    projected_elements = factory2d.CreateProjections(create_reference(part, source_feature))
    return set_object_name(get_first_item(projected_elements), name)


def create_sketch_intersections(
    part: Any,
    sketch: Any,
    factory2d: Any,
    source_feature: Any,
) -> tuple[list[Any], list[Any]]:
    """功能: 创建草图相交并收集候选元素；输入: Part、草图、2D工厂和源特征；输出: 相交候选和新增草图元素。"""
    geometric_elements = sketch.GeometricElements
    element_count_before = int(geometric_elements.Count)
    intersected_elements = factory2d.CreateIntersections(create_reference(part, source_feature))
    returned_candidates = list(iter_collection(intersected_elements))
    sketch.Evaluate()
    element_count_after = int(geometric_elements.Count)
    newly_created_elements = [
        geometric_elements.Item(index)
        for index in range(element_count_before + 1, element_count_after + 1)
    ]
    intersection_candidates = newly_created_elements or returned_candidates
    if not intersection_candidates:
        raise LookupError(f"草图与“{source_feature.Name}”相交后没有生成几何元素。")
    return intersection_candidates, newly_created_elements


def get_intersection_arc_endpoints(
    document: Any,
    newly_created_elements: list[Any],
    endpoint_name_prefix: str,
) -> tuple[list[Any], list[tuple[float, float]], tuple[float, float]]:
    """功能: 从相交新增草图元素中寻找两个圆弧范围端点；输入: 文档、新增元素和端点名前缀；输出: 端点对象、端点坐标和中点坐标。"""
    point_candidates: list[tuple[Any, tuple[float, float]]] = []
    for index, candidate in enumerate(newly_created_elements, start=1):
        try:
            coordinates = evaluate_object_array(
                document,
                candidate,
                "GetCoordinates",
                2,
                f"相交候选{index}坐标",
            )
        except RuntimeError:
            continue
        point_candidates.append((candidate, (coordinates[0], coordinates[1])))

    if len(point_candidates) < 2:
        raise RuntimeError(
            f"本次草图相交新增元素中只找到 {len(point_candidates)} 个点，"
            "无法确定相交圆弧的两个范围端点。"
        )

    endpoint_pair = max(
        (
            (first, second)
            for first_index, first in enumerate(point_candidates)
            for second in point_candidates[first_index + 1 :]
        ),
        key=lambda pair: math.hypot(
            pair[0][1][0] - pair[1][1][0],
            pair[0][1][1] - pair[1][1][1],
        ),
    )
    endpoint_objects = [endpoint_pair[0][0], endpoint_pair[1][0]]
    endpoint_coordinates = [endpoint_pair[0][1], endpoint_pair[1][1]]
    for index, endpoint in enumerate(endpoint_objects, start=1):
        set_object_name(endpoint, f"{endpoint_name_prefix}{index}")

    midpoint = (
        (endpoint_coordinates[0][0] + endpoint_coordinates[1][0]) / 2.0,
        (endpoint_coordinates[0][1] + endpoint_coordinates[1][1]) / 2.0,
    )
    return endpoint_objects, endpoint_coordinates, midpoint


def constrain_point_on_intersection(
    document: Any,
    part: Any,
    constraints: Any,
    sketch_point: Any,
    intersection_candidates: list[Any],
    intersection_name: str,
    constraint_name: str,
) -> tuple[Any, str, list[str]]:
    """功能: 从相交候选中找到可用于点在线约束的曲线并创建约束；输入: 文档、Part、约束集合、草图点、候选、相交名和约束名；输出: 曲线对象、约束名和警告列表。"""
    warnings: list[str] = []
    for index, candidate in enumerate(intersection_candidates, start=1):
        try:
            evaluate_object_array(
                document,
                candidate,
                "GetCoordinates",
                2,
                f"相交候选{index}坐标",
            )
            continue
        except RuntimeError:
            pass

        created_name, warning = try_add_on_constraint(
            part,
            constraints,
            sketch_point,
            candidate,
            constraint_name,
        )
        if created_name:
            set_object_name(candidate, intersection_name)
            return candidate, created_name, warnings
        warnings.append(f"相交候选{index}: {warning}")

    detail = "；".join(warnings)
    raise RuntimeError(f"{constraint_name}创建失败，草图点无法约束在镜片相交线上。{detail}")


def create_sketch_line_by_xy(
    factory2d: Any,
    start_xy: tuple[float, float],
    end_xy: tuple[float, float],
    name: str,
) -> Any:
    """功能: 在草图中按二维坐标创建直线；输入: 2D工厂、起终点坐标和名称；输出: 草图直线。"""
    line = factory2d.CreateLine(start_xy[0], start_xy[1], end_xy[0], end_xy[1])
    return set_object_name(line, name)


def project_point_to_plane_2d(point: Vector, plane_origin: Vector, plane_x: Vector, plane_y: Vector) -> tuple[float, float]:
    """功能: 将 3D 点按平面坐标轴转换为草图 2D 坐标；输入: 三维点、平面原点、X方向和Y方向；输出: 二维坐标。"""
    offset = subtract(point, plane_origin)
    x_axis = normalize(plane_x, "草图横向")
    y_axis = normalize(plane_y, "草图纵向")
    return dot(offset, x_axis), dot(offset, y_axis)


def sketch_point_to_part_coordinates(
    document: Any,
    sketch: Any,
    sketch_point: Any,
) -> tuple[tuple[float, float], Vector]:
    """功能: 读取约束求解后的草图点并转换为零件坐标；输入: 文档、草图和草图点；输出: 草图二维坐标和零件三维坐标。"""
    sketch.Evaluate()
    point_xy = evaluate_object_array(
        document,
        sketch_point,
        "GetCoordinates",
        2,
        "草图内反射取点1的草图坐标",
    )
    axis_data = evaluate_object_array(
        document,
        sketch,
        "GetAbsoluteAxisData",
        9,
        "草图绝对坐标轴",
    )
    origin = as_vector(axis_data[0:3], "草图原点")
    x_axis = as_vector(axis_data[3:6], "草图横向")
    y_axis = as_vector(axis_data[6:9], "草图纵向")
    part_coordinates = add(
        origin,
        add(scale(x_axis, point_xy[0]), scale(y_axis, point_xy[1])),
    )
    return (point_xy[0], point_xy[1]), part_coordinates


def get_catia_constant(name: str) -> Any:
    """功能: 获取 CATIA COM 常量并提供兜底值；输入: 常量名；输出: 常量值。"""
    fallback_constants = {
        "catCstModeDrivingDimension": 0,
        "catCstTypeOn": 2,
        "catCstTypeSymmetry": 15,
    }
    try:
        return getattr(win32com.client.constants, name)
    except Exception:
        if name in fallback_constants:
            return fallback_constants[name]
        raise RuntimeError(f"当前环境无法读取 CATIA 常量: {name}")


def try_add_on_constraint(
    part: Any,
    constraints: Any,
    first_element: Any,
    second_element: Any,
    name: str,
) -> tuple[str | None, str | None]:
    """功能: 尝试创建两元素相合/点在线约束；输入: Part、约束集合、两个元素和名称；输出: 约束名和警告。"""
    try:
        constraint = constraints.AddBiEltCst(
            get_catia_constant("catCstTypeOn"),
            create_reference(part, first_element),
            create_reference(part, second_element),
        )
        set_object_name(constraint, name)
        try:
            constraint.Mode = get_catia_constant("catCstModeDrivingDimension")
        except Exception:
            pass
        return str(getattr(constraint, "Name", name)), None
    except Exception as exc:
        return None, f"{name} 创建失败: {exc}"


def try_add_symmetry_constraint(
    part: Any,
    constraints: Any,
    first_line: Any,
    second_line: Any,
    symmetry_axis: Any,
    name: str,
) -> tuple[str | None, str | None]:
    """功能: 尝试创建两条线关于轴线的对称约束；输入: Part、约束集合、两线、轴线和名称；输出: 约束名和警告。"""
    try:
        constraint = constraints.AddTriEltCst(
            get_catia_constant("catCstTypeSymmetry"),
            create_reference(part, first_line),
            create_reference(part, second_line),
            create_reference(part, symmetry_axis),
        )
        set_object_name(constraint, name)
        try:
            constraint.Mode = get_catia_constant("catCstModeDrivingDimension")
        except Exception:
            pass
        return str(getattr(constraint, "Name", name)), None
    except Exception as exc:
        return None, f"{name} 创建失败: {exc}"


def populate_reflection_sketch(
    document: Any,
    part: Any,
    sketch: Any,
    intersection_feature: Any,
    mirror_center: Any,
    initial_point_feature: Any,
    regulation_point: Any,
    eye: Any,
    mirror_side_name: str,
    regulation_point_name: str,
    eye_side_name: str,
    point_index: int,
    center_initial_xy: tuple[float, float],
    initial_point_initial_xy: tuple[float, float],
    regulation_initial_xy: tuple[float, float],
    eye_initial_xy: tuple[float, float],
) -> dict[str, Any]:
    """功能: 在反射取点草图中创建相交、投影、连线和约束；输入: 文档、Part、草图、相交特征、球心、初始点、法规点、眼点、侧别和初始坐标；输出: 草图过程结果字典。"""
    part.InWorkObject = sketch
    factory2d = sketch.OpenEdition()
    constraint_warnings: list[str] = []
    constraint_names: list[str] = []
    try:
        mirror_intersection_candidates, mirror_intersection_new_elements = (
            create_sketch_intersections(
                part,
                sketch,
                factory2d,
                intersection_feature,
            )
        )

        center_projection = create_sketch_projection(
            part,
            factory2d,
            mirror_center,
            f"{mirror_side_name}镜片球心投影",
        )
        initial_point_projection = create_sketch_projection(
            part,
            factory2d,
            initial_point_feature,
            f"{mirror_side_name}镜片法线与镜片交点投影",
        )
        reflection_point_initial_xy = initial_point_initial_xy
        regulation_projection = create_sketch_projection(
            part,
            factory2d,
            regulation_point,
            f"{regulation_point_name}投影",
        )
        eye_projection = create_sketch_projection(
            part,
            factory2d,
            eye,
            f"{eye_side_name}眼点投影",
        )

        reflection_sketch_point = set_object_name(
            factory2d.CreatePoint(*reflection_point_initial_xy),
            f"草图内反射取点{point_index}",
        )

        center_to_reflection_line = create_sketch_line_by_xy(
            factory2d,
            center_initial_xy,
            reflection_point_initial_xy,
            f"球心到反射取点{point_index}",
        )
        regulation_to_reflection_line = create_sketch_line_by_xy(
            factory2d,
            regulation_initial_xy,
            reflection_point_initial_xy,
            f"法规点到反射取点{point_index}",
        )
        eye_to_reflection_line = create_sketch_line_by_xy(
            factory2d,
            eye_initial_xy,
            reflection_point_initial_xy,
            f"{eye_side_name}眼点到反射取点{point_index}",
        )

        constraints = sketch.Constraints
        mirror_intersection, point_on_intersection_constraint_name, intersection_warnings = (
            constrain_point_on_intersection(
                document,
                part,
                constraints,
                reflection_sketch_point,
                mirror_intersection_candidates,
                f"{mirror_side_name}镜面球面相交线",
                f"反射取点{point_index}与镜面球面相交线相合",
            )
        )
        constraint_names.append(point_on_intersection_constraint_name)
        constraint_warnings.extend(intersection_warnings)

        endpoint_constraints = (
            (f"球心到反射取点{point_index}起点相合", center_to_reflection_line.StartPoint, center_projection),
            (f"球心到反射取点{point_index}终点相合", center_to_reflection_line.EndPoint, reflection_sketch_point),
            (f"法规点到反射取点{point_index}起点相合", regulation_to_reflection_line.StartPoint, regulation_projection),
            (f"法规点到反射取点{point_index}终点相合", regulation_to_reflection_line.EndPoint, reflection_sketch_point),
            (f"{eye_side_name}眼点到反射取点{point_index}起点相合", eye_to_reflection_line.StartPoint, eye_projection),
            (f"{eye_side_name}眼点到反射取点{point_index}终点相合", eye_to_reflection_line.EndPoint, reflection_sketch_point),
        )
        for constraint_name, first_element, second_element in endpoint_constraints:
            created_name, warning = try_add_on_constraint(
                part,
                constraints,
                first_element,
                second_element,
                constraint_name,
            )
            if created_name:
                constraint_names.append(created_name)
            if warning:
                constraint_warnings.append(warning)

        created_name, warning = try_add_symmetry_constraint(
            part,
            constraints,
            regulation_to_reflection_line,
            eye_to_reflection_line,
            center_to_reflection_line,
            f"反射对称约束{point_index}",
        )
        symmetry_constraint_name = created_name
        symmetry_constraint_warning = warning
        if created_name:
            constraint_names.append(created_name)
        if warning:
            constraint_warnings.append(warning)

        if not symmetry_constraint_name:
            raise RuntimeError(f"对称约束设置失败: {symmetry_constraint_warning}")

        reflection_point_sketch_coordinates, reflection_point_part_coordinates = (
            sketch_point_to_part_coordinates(
                document,
                sketch,
                reflection_sketch_point,
            )
        )

        return {
            "creation_method": "sketch_constraint_on_mirror_sphere",
            "intersection_names": [f"{mirror_side_name}镜面球面相交线"],
            "intersection_created_element_count": len(mirror_intersection_new_elements),
            "projection_names": [
                f"{mirror_side_name}镜片球心投影",
                f"{mirror_side_name}镜片法线与镜片交点投影",
                f"{regulation_point_name}投影",
                f"{eye_side_name}眼点投影",
            ],
            "sketch_point_names": [f"草图内反射取点{point_index}"],
            "process_line_names": [
                f"球心到反射取点{point_index}",
                f"法规点到反射取点{point_index}",
                f"{eye_side_name}眼点到反射取点{point_index}",
            ],
            "constraint_names": constraint_names,
            "constraint_warnings": constraint_warnings,
            "point_on_intersection_constraint_name": point_on_intersection_constraint_name,
            "intersection_endpoint_names": [],
            "intersection_endpoint_coordinates": [],
            "reflection_point_initial_coordinates": reflection_point_initial_xy,
            "reflection_point_sketch_coordinates": reflection_point_sketch_coordinates,
            "reflection_point_part_coordinates": reflection_point_part_coordinates,
        }
    finally:
        sketch.CloseEdition()
        part.Update()


def create_regulation_reflection_points(
    document: Any,
    part: Any,
    reflection_point_hybrid_body_name: str,
    parametric_rearview_hybrid_body: Any,
    regulation_line_hybrid_body: Any,
    left_mirror: Any,
    right_mirror: Any,
    left_eye: Any,
    right_eye: Any,
    selected_mirror_sides: set[str] | None = None,
    temporary_reflection_points: dict[str, Any] | None = None,
) -> dict[str, Any]:
    """功能: 创建通过临时测量镜片侧的法规反射取点；输入: 文档、Part、几何集名、参数化镜片集、法规线集、左右镜片、左右眼点和可选镜片侧；输出: 法规反射取点结果字典。"""
    reflection_hybrid_body = create_hybrid_body(part, reflection_point_hybrid_body_name)
    if selected_mirror_sides is None:
        selected_mirror_sides = {"左", "右"}
    mirror_centers: dict[str, Any] = {}
    if "左" in selected_mirror_sides:
        mirror_centers["左"] = create_point_center(
            part,
            reflection_hybrid_body,
            left_mirror,
            "左镜片球心",
        )
    if "右" in selected_mirror_sides:
        mirror_centers["右"] = create_point_center(
            part,
            reflection_hybrid_body,
            right_mirror,
            "右镜片球心",
        )
    mirror_sphere_data: dict[str, dict[str, Any]] = {}
    for mirror_side_name, mirror_feature in (("左", left_mirror), ("右", right_mirror)):
        mirror_center = mirror_centers.get(mirror_side_name)
        if mirror_side_name not in selected_mirror_sides or mirror_center is None:
            continue
        normal_intersection = find_feature(
            parametric_rearview_hybrid_body,
            f"{mirror_side_name}镜片法线与镜片交点",
        )
        sphere_radius = get_minimum_distance(document, part, mirror_center, mirror_feature)
        sphere_name = f"{mirror_side_name}镜面球面"
        sphere_feature = create_sphere_surface(
            document,
            part,
            reflection_hybrid_body,
            mirror_center,
            sphere_radius,
            sphere_name,
        )
        hide_object(document, sphere_feature)
        mirror_sphere_data[mirror_side_name] = {
            "sphere_radius": sphere_radius,
            "sphere_name": str(sphere_feature.Name),
            "sphere_created": True,
            "sphere_hidden": True,
            "sphere_warning": None,
            "_sphere_feature": sphere_feature,
            "_normal_intersection": normal_intersection,
        }
    mirror_data = {
        "左": (left_mirror, mirror_centers.get("左")),
        "右": (right_mirror, mirror_centers.get("右")),
    }
    eye_data = {
        "右": right_eye,
        "左": left_eye,
    }
    results: list[dict[str, Any]] = []

    for eye_side_name, eye in eye_data.items():
        for regulation_side_name, (mirror, mirror_center) in mirror_data.items():
            if regulation_side_name not in selected_mirror_sides or mirror_center is None:
                continue
            sphere_center = get_point(document, part, mirror_center)
            sphere_radius = float(mirror_sphere_data[regulation_side_name]["sphere_radius"])
            eye_point = get_point(document, part, eye)

            for point_index in range(1, 7):
                regulation_point_name = f"{regulation_side_name}法规点{point_index}"
                regulation_point_feature = find_feature(
                    regulation_line_hybrid_body,
                    regulation_point_name,
                )
                group_name = f"{eye_side_name}眼{regulation_side_name}法规点{point_index}"
                reflection_point_name = (
                    f"{eye_side_name}眼{regulation_side_name}法规反射取点{point_index}"
                )
                regulation_point = get_point(document, part, regulation_point_feature)
                reference_plane = create_plane_3_points(
                    part,
                    reflection_hybrid_body,
                    mirror_center,
                    regulation_point_feature,
                    eye,
                    f"{group_name}参考平面",
                )
                reference_sketch = create_sketch_on_plane(
                    part,
                    reflection_hybrid_body,
                    reference_plane,
                    f"{reflection_point_name}草图",
                )
                plane_origin, plane_x, plane_y = get_plane(document, part, reference_plane)
                sketch_process = populate_reflection_sketch(
                    document,
                    part,
                    reference_sketch,
                    mirror_sphere_data[regulation_side_name]["_sphere_feature"],
                    mirror_center,
                    mirror_sphere_data[regulation_side_name]["_normal_intersection"],
                    regulation_point_feature,
                    eye,
                    regulation_side_name,
                    regulation_point_name,
                    eye_side_name,
                    point_index,
                    project_point_to_plane_2d(sphere_center, plane_origin, plane_x, plane_y),
                    project_point_to_plane_2d(
                        get_point(
                            document,
                            part,
                            mirror_sphere_data[regulation_side_name]["_normal_intersection"],
                        ),
                        plane_origin,
                        plane_x,
                        plane_y,
                    ),
                    project_point_to_plane_2d(regulation_point, plane_origin, plane_x, plane_y),
                    project_point_to_plane_2d(eye_point, plane_origin, plane_x, plane_y),
                )
                reflection_point = create_coordinate_point(
                    part,
                    reflection_hybrid_body,
                    sketch_process["reflection_point_part_coordinates"],
                    reflection_point_name,
                )
                reflection_point_coordinates = get_point(document, part, reflection_point)
                hide_object(document, reference_plane)
                hide_object(document, reference_sketch)
                distance_to_mirror_surface = get_minimum_distance(
                    document,
                    part,
                    reflection_point,
                    mirror,
                )
                print(
                    f"[法规反射取点] {reflection_point_name}: "
                    f"坐标={round_vector(reflection_point_coordinates)}，"
                    f"参考平面={reference_plane.Name}，"
                    f"草图={reference_sketch.Name}，"
                    f"到裁剪镜片距离={round(distance_to_mirror_surface, 3)} mm",
                    flush=True,
                )
                sketch_process["mirror_sphere_name"] = mirror_sphere_data[regulation_side_name]["sphere_name"]
                sketch_process["mirror_sphere_created"] = mirror_sphere_data[regulation_side_name]["sphere_created"]
                sketch_process["mirror_sphere_warning"] = mirror_sphere_data[regulation_side_name]["sphere_warning"]
                sketch_process["reflection_point_sketch_coordinates"] = [
                    round(value, 2)
                    for value in sketch_process["reflection_point_sketch_coordinates"]
                ]
                sketch_process["reflection_point_part_coordinates"] = round_vector(
                    sketch_process["reflection_point_part_coordinates"]
                )
                sketch_process["reflection_point_initial_coordinates"] = [
                    round(value, 2)
                    for value in sketch_process["reflection_point_initial_coordinates"]
                ]
                sketch_process["distance_to_trimmed_mirror_surface"] = round(
                    distance_to_mirror_surface,
                    3,
                )
                sketch_process["point_allowed_outside_trimmed_mirror"] = True
                results.append(
                    {
                        "reflection_point_name": str(reflection_point.Name),
                        "reflection_point_coordinates": round_vector(
                            reflection_point_coordinates
                        ),
                        "reference_plane_name": str(reference_plane.Name),
                        "reference_sketch_name": str(reference_sketch.Name),
                        "reference_plane_hidden": True,
                        "reference_sketch_hidden": True,
                        "mirror_radius": round(sphere_radius, 2),
                        "distance_to_trimmed_mirror_surface": round(distance_to_mirror_surface, 3),
                        "point_allowed_outside_trimmed_mirror": True,
                        "sketch_process": sketch_process,
                        "eye_side_name": eye_side_name,
                        "mirror_side_name": regulation_side_name,
                        "point_index": point_index,
                        "_feature": reflection_point,
                        "_coordinates_full": reflection_point_coordinates,
                    }
                )

    expected_count = len(selected_mirror_sides) * 12
    if len(results) != expected_count:
        raise RuntimeError(
            f"法规反射取点创建数量错误，应为 {expected_count}，实际为 {len(results)}。"
        )

    triangle_frames = create_reflection_triangle_frames(
        document,
        part,
        reflection_hybrid_body,
        results,
    )

    best_view_frames = {}
    temporary_mirror_results = (
        temporary_reflection_points.get("mirrors", {})
        if temporary_reflection_points
        else {}
    )
    for mirror_side_name in ("左", "右"):
        if mirror_side_name not in selected_mirror_sides:
            best_view_frames[mirror_side_name] = {
                "skipped": True,
                "reason": "临时法规反射取点测量未通过",
            }
            continue
        temporary_mirror_result = temporary_mirror_results.get(mirror_side_name, {})
        final_view_point_names = [
            str(point["reflection_point_name"])
            for point in temporary_mirror_result.get("final_view_points", [])
        ]
        if temporary_reflection_points is not None and len(final_view_point_names) != 4:
            best_view_frames[mirror_side_name] = {
                "skipped": True,
                "reason": (
                    "临时法规反射取点未能从搜索结果中选出4个在镜面上的最终视野点，"
                    "不使用12点最大面积方式回退。"
                ),
                "temporary_final_view_point_count": len(final_view_point_names),
            }
            continue
        if len(final_view_point_names) != 4:
            final_view_point_names = None
        best_view_frames[mirror_side_name] = create_best_view_frame(
            document,
            part,
            reflection_hybrid_body,
            mirror_side_name,
            get_point(document, part, mirror_centers[mirror_side_name]),
            [record for record in results if record["mirror_side_name"] == mirror_side_name],
            final_view_point_names=final_view_point_names,
        )
    best_view_distance_annotation_data = collect_best_view_distance_annotation_data(
        document=document,
        part=part,
        mirror_data={
            "左": left_mirror,
            "右": right_mirror,
        },
        best_view_frames=best_view_frames,
        selected_mirror_sides=selected_mirror_sides,
    )
    hidden_reflection_point_names: list[str] = []
    for record in results:
        hide_object(document, record["_feature"])
        hidden_reflection_point_names.append(record["reflection_point_name"])

    for record in results:
        record.pop("_feature", None)
        record.pop("_coordinates_full", None)

    return {
        "geo_set_name": str(reflection_hybrid_body.Name),
        "selected_mirror_sides": sorted(selected_mirror_sides),
        "left_mirror_center_name": (
            str(mirror_centers["左"].Name) if "左" in mirror_centers else None
        ),
        "right_mirror_center_name": (
            str(mirror_centers["右"].Name) if "右" in mirror_centers else None
        ),
        "mirror_spheres": {
            side_name: {
                key: value
                for key, value in sphere_result.items()
                if not key.startswith("_")
            } | {
                "sphere_radius": round(float(sphere_result["sphere_radius"]), 2),
            }
            for side_name, sphere_result in mirror_sphere_data.items()
        },
        "reflection_point_count": len(results),
        "reflection_points_hidden": True,
        "hidden_reflection_point_names": hidden_reflection_point_names,
        "reflection_points": results,
        "triangle_frames": triangle_frames,
        "best_view_frames": best_view_frames,
        "best_view_distance_annotation_data": best_view_distance_annotation_data,
    }


def print_section(title: str) -> None:
    """功能: 输出命令行摘要分节标题；输入: 标题文本；输出: 无。"""
    print(f"\n=== {title} ===")


def format_angle_offset(horizontal_angle: Any, vertical_angle: Any) -> str:
    """功能: 将水平/竖直角度格式化为中文偏转说明；输入: 水平角和竖直角；输出: 偏转说明文本。"""
    if horizontal_angle is None or vertical_angle is None:
        return "未找到满足条件的偏转角度"

    descriptions: list[str] = []
    horizontal_value = float(horizontal_angle)
    vertical_value = float(vertical_angle)
    if abs(horizontal_value) > 1e-9:
        direction = "右" if horizontal_value > 0 else "左"
        descriptions.append(f"向{direction}偏转{abs(horizontal_value):.1f}度")
    if abs(vertical_value) > 1e-9:
        direction = "下" if vertical_value > 0 else "上"
        descriptions.append(f"向{direction}偏转{abs(vertical_value):.1f}度")
    return "，".join(descriptions) if descriptions else "不偏转"


def format_calibration_status(mirror_result: dict[str, Any]) -> str:
    """功能: 将单侧镜片校核结果格式化为报告文本；输入: 单侧镜片结果字典；输出: 成功/失败文本。"""
    if mirror_result.get("calibration_success") or mirror_result.get("all_points_pass"):
        return "成功"
    reason = mirror_result.get("failure_reason")
    return f"失败（{reason}）" if reason else "失败"


def print_result_summary(result: dict[str, Any]) -> None:
    """功能: 按程序执行顺序输出本地调试摘要；输入: 算法结果字典；输出: 无。"""
    if not result.get("success"):
        print(f"执行失败: {result.get('error', '未知错误')}", file=sys.stderr)
        return

    print_section("基础信息")
    print(f"读取文件: {result.get('document_path')}")
    print(f"输出文件夹: {result.get('output_dir')}")
    print(f"另存文件: {result.get('saved_as_path')}")
    print(f"标注CATPart文件: {result.get('annotation_part_path')}")
    print(f"标注CATProduct文件: {result.get('annotation_product_path')}")
    print(f"输入参数几何图形集: {result.get('input_parameter_geo_set_name')}")
    print(f"输入特征: {result.get('input_feature_names')}")

    print_section("车辆方向")
    for name, vector in result.get("directions", {}).items():
        print(f"{name}: {vector}")

    print_section("法规线")
    print(f"法规线几何集: {result.get('regulation_line_geo_set_name')}")
    for side_key, title in (
        ("left_regulation_line", "左法规线"),
        ("right_regulation_line", "右法规线"),
    ):
        side_result = result.get(side_key, {})
        print(f"\n{title}:")
        hidden_text = "（已隐藏）" if side_result.get("mirror_extremum_point_hidden") else ""
        print(f"镜片极值点: {side_result.get('mirror_extremum_point_name')}{hidden_text}")
        print(f"极值点到地面距离: {side_result.get('mirror_extremum_to_ground_distance')}")
        print(f"投影点: {side_result.get('projection_point_names')}")
        extension = side_result.get("left_extension") or side_result.get("right_extension")
        if extension:
            print(
                f"{title[:1]}车宽线延长: "
                f"端点={extension.get('endpoint_name')} "
                f"坐标={extension.get('endpoint_coordinates')} "
                f"线={extension.get('line_name')} "
                f"长度={extension.get('line_length')}"
            )
        print("法规点坐标:")
        for point_name, coordinates in side_result.get("point_coordinates", {}).items():
            print(f"  {point_name}: {coordinates}")
        print(f"法规线: {side_result.get('line_names')}")

    print_section("法规视野检查")
    regulation_vision_check = result.get("regulation_vision_check", {})
    print(f"全部合格: {regulation_vision_check.get('all_pass')}")
    for check in regulation_vision_check.get("checks", []):
        print(
            f"{check.get('item_no')} {check.get('content')} "
            f"{'合格' if check.get('pass') else '不合格'}"
        )
        print(f"  依据: {check.get('basis')}")
        print(f"  测量: {check.get('measure')}")

    print_section("参数化后视镜")
    parametric_result = result.get("parametric_rearview_mirror", {})
    print(f"参数化后视镜几何集: {parametric_result.get('geo_set_name')}")
    for parameter_result in parametric_result.get("parameters", []):
        print(
            f"{parameter_result.get('name')}: "
            f"{parameter_result.get('value')} ({parameter_result.get('type')})"
        )
    for mirror_side_name in ("左", "右"):
        mirror_result = parametric_result.get("mirrors", {}).get(mirror_side_name, {})
        print(f"\n{mirror_side_name}镜片:")
        print(
            f"重心点: {mirror_result.get('centroid_name')}, "
            f"坐标={mirror_result.get('centroid_coordinates')}"
        )
        print(f"法线: {mirror_result.get('normal_name')}")
        print(
            f"法线与镜片交点: {mirror_result.get('normal_intersection_name')}, "
            f"坐标={mirror_result.get('normal_intersection_coordinates')}"
        )
        print(
            f"旋转中心: {mirror_result.get('rotation_center_name')}, "
            f"坐标={mirror_result.get('rotation_center_coordinates')}"
        )
        print(f"旋转参考平面: {mirror_result.get('rotation_reference_plane_name')}")
        sketch_hidden_text = (
            "（已隐藏）" if mirror_result.get("rotation_axis_sketch_hidden") else ""
        )
        print(
            f"旋转轴草图: {mirror_result.get('rotation_axis_sketch_name')}"
            f"{sketch_hidden_text}"
        )
        print(f"草图过程线: {mirror_result.get('rotation_axis_sketch_line_names')}")
        print(f"3D旋转轴: {mirror_result.get('rotation_axis_line_names')}")
        print(f"旋转特征: {mirror_result.get('rotation_feature_names')}")
        for rotation_result in mirror_result.get("rotation_features", []):
            print(
                f"  {rotation_result.get('name')}: "
                f"源={rotation_result.get('source_element_name')}, "
                f"轴={rotation_result.get('axis_name')}, "
                f"参数={rotation_result.get('angle_parameter_name')}, "
                f"目标角度={rotation_result.get('rotation_angle_parameter_name')}, "
                f"公式={rotation_result.get('formula_name') or '未创建'}, "
                f"表达式={rotation_result.get('formula_body') or '无'}, "
                f"警告={rotation_result.get('formula_warning') or '无'}"
            )
        print(f"水平轴方向: {mirror_result.get('horizontal_axis_direction')}")
        print(f"竖直轴方向: {mirror_result.get('vertical_axis_direction')}")
    print(f"\n保留显示特征: {parametric_result.get('visible_feature_names')}")
    hidden_process_names = parametric_result.get("hidden_process_feature_names", [])
    print(f"隐藏过程特征数量: {len(hidden_process_names)}")
    print(f"隐藏过程特征: {hidden_process_names}")

    print_section("临时法规反射取点测量")
    temporary_result = result.get("temporary_reflection_points", {})
    print(f"临时反射取点总数: {temporary_result.get('total_reflection_point_count')}")
    print(
        f"间隙阈值: {temporary_result.get('clearance_threshold')} mm, "
        f"角度范围: ±{temporary_result.get('angle_limit')} deg, "
        f"步长: {temporary_result.get('angle_step')} deg"
    )

    for mirror_side_name in ("左", "右"):
        mirror_result = temporary_result.get("mirrors", {}).get(mirror_side_name, {})
        print(f"\n{mirror_side_name}镜片:")
        print(f"校核结果: {format_calibration_status(mirror_result)}")
        print(f"全部通过: {mirror_result.get('all_points_pass')}")
        print(f"所有点在镜面上: {mirror_result.get('all_points_on_surface')}")
        print(
            "偏转说明: "
            f"{format_angle_offset(mirror_result.get('horizontal_angle'), mirror_result.get('vertical_angle'))}"
        )
        print(
            f"搜索角度: 水平={mirror_result.get('horizontal_angle')} deg, "
            f"竖直={mirror_result.get('vertical_angle')} deg, "
            f"已测试组合数={mirror_result.get('searched_angle_count')}"
        )
        print(f"临时几何集: {mirror_result.get('geo_set_name')}")
        print(f"旋转镜片: {mirror_result.get('rotated_mirror_name')}")
        print(
            f"临时球心: {mirror_result.get('sphere_center_name')}, "
            f"坐标={mirror_result.get('sphere_center_coordinates')}"
        )
        print(f"球面半径: {mirror_result.get('sphere_radius')}")
        print(f"最小边界距离: {mirror_result.get('min_distance_to_boundary')} mm")
        print(
            f"最终视野点组通过: {mirror_result.get('all_final_view_groups_pass')}, "
            f"最终视野点数量: {mirror_result.get('final_view_point_count')}"
        )
        print(f"最大离镜面距离: {mirror_result.get('max_distance_to_mirror_surface')} mm")
        print(f"生成反射取点数量: {mirror_result.get('reflection_point_count')}")
        if not mirror_result.get("all_points_pass"):
            print(f"最佳失败结果: {mirror_result.get('best_failed_result')}")
        for final_point in mirror_result.get("final_view_points", []):
            print(
                f"  {final_point.get('group_name')} -> "
                f"{final_point.get('reflection_point_name')}: "
                f"{final_point.get('reflection_point_coordinates')}, "
                f"边界距离={final_point.get('distance_to_boundary')} mm"
            )
        for point_result in mirror_result.get("reflection_points", []):
            print(
                f"  {point_result.get('reflection_point_name')}: "
                f"{point_result.get('reflection_point_coordinates')}, "
                f"边界距离={point_result.get('distance_to_boundary')} mm, "
                f"离镜面距离={point_result.get('distance_to_mirror_surface')} mm, "
                f"在镜面上={point_result.get('is_point_on_mirror_surface')}, "
                f">=3mm={point_result.get('is_clearance_greater_equal_3mm')}"
            )

    print_section("法规反射取点")
    reflection_result = result.get("regulation_reflection_points", {})
    print(f"法规反射取点几何集: {result.get('regulation_reflection_point_geo_set_name')}")
    print(f"左镜片球心: {reflection_result.get('left_mirror_center_name')}")
    print(f"右镜片球心: {reflection_result.get('right_mirror_center_name')}")
    for mirror_side_name, sphere_result in reflection_result.get("mirror_spheres", {}).items():
        print(
            f"{mirror_side_name}镜面球面: {sphere_result.get('sphere_name')}, "
            f"半径={sphere_result.get('sphere_radius')} mm, "
            f"已创建={sphere_result.get('sphere_created')}, "
            f"警告={sphere_result.get('sphere_warning') or '无'}"
        )
    print(f"生成反射取点数量: {reflection_result.get('reflection_point_count')}")
    for point_result in reflection_result.get("reflection_points", []):
        sketch_process = point_result.get("sketch_process", {})
        print(f"\n{point_result.get('reflection_point_name')}:")
        print(f"创建方式: {sketch_process.get('creation_method')}")
        print(f"镜面球面: {sketch_process.get('mirror_sphere_name')}")
        print(f"允许位于裁剪镜片外: {point_result.get('point_allowed_outside_trimmed_mirror')}")
        print(f"到裁剪镜片距离: {point_result.get('distance_to_trimmed_mirror_surface')} mm")
        print(f"参考平面: {point_result.get('reference_plane_name')}")
        print(f"草图: {point_result.get('reference_sketch_name')}")
        print(f"相交/球面: {sketch_process.get('intersection_names')}")
        print(
            f"本次相交新增草图元素数量: "
            f"{sketch_process.get('intersection_created_element_count')}"
        )
        print(f"草图投影: {sketch_process.get('projection_names')}")
        print(f"草图点: {sketch_process.get('sketch_point_names')}")
        print(f"草图线: {sketch_process.get('process_line_names')}")
        print(f"相交圆弧端点: {sketch_process.get('intersection_endpoint_names')}")
        print(
            f"相交圆弧端点坐标: "
            f"{sketch_process.get('intersection_endpoint_coordinates')}"
        )
        print(
            f"草图内反射取点初始中点坐标: "
            f"{sketch_process.get('reflection_point_initial_coordinates')}"
        )
        print(
            "点在线约束: "
            f"{sketch_process.get('point_on_intersection_constraint_name')}"
        )
        print(f"约束: {sketch_process.get('constraint_names')}")
        print(f"约束警告: {sketch_process.get('constraint_warnings') or '无'}")
        print(
            f"更新后草图坐标: "
            f"{sketch_process.get('reflection_point_sketch_coordinates')}"
        )
        print(
            f"转换后零件坐标: "
            f"{sketch_process.get('reflection_point_part_coordinates')}"
        )
        print(f"生成点实际坐标: {point_result.get('reflection_point_coordinates')}")

    print("\n三角线框:")
    for group_name, frame_result in reflection_result.get("triangle_frames", {}).items():
        if frame_result.get("skipped"):
            print(f"  {group_name}: 已跳过（{frame_result.get('reason')}）")
            continue
        print(
            f"  {group_name}: {frame_result.get('line_names')}, "
            f"颜色={frame_result.get('line_color')}"
        )
    hidden_point_names = reflection_result.get("hidden_reflection_point_names", [])
    print(f"法规反射取点已隐藏: {reflection_result.get('reflection_points_hidden')}")
    print(f"隐藏反射取点数量: {len(hidden_point_names)}")
    annotation_data = reflection_result.get("best_view_distance_annotation_data", {})
    print(f"最佳视野距离标注数据数量: {annotation_data.get('item_count')}")

    print_section("最佳视野线框")
    for mirror_side_name in ("左", "右"):
        frame_result = reflection_result.get("best_view_frames", {}).get(
            mirror_side_name,
            {},
        )
        print(f"\n{mirror_side_name}镜片:")
        if frame_result.get("skipped"):
            print(f"已跳过: {frame_result.get('reason')}")
            continue
        print(f"最佳视野边界点: {frame_result.get('boundary_point_names')}")
        print(f"边界点坐标: {frame_result.get('boundary_point_coordinates')}")
        print(f"镜片局部二维坐标: {frame_result.get('boundary_local_coordinates')}")
        print(
            f"选点来源: {frame_result.get('selection_source')}, "
            f"是否凸四边形: {frame_result.get('is_convex_quadrilateral')}"
        )
        print(f"最佳视野四边形面积: {frame_result.get('quadrilateral_area')}")
        print(f"红色线框: {frame_result.get('frame_line_names')}")

    print_section("测量点间隙校验")
    gap_check_result = result.get("gap_check", {})
    print(f"测量点间隙校验几何集: {gap_check_result.get('geo_set_name')}")
    print(f"间隙阈值: {gap_check_result.get('clearance_threshold')} mm")
    for mirror_side_name in ("左", "右"):
        mirror_result = gap_check_result.get("mirrors", {}).get(mirror_side_name, {})
        print(f"\n{mirror_side_name}镜片:")
        if mirror_result.get("skipped"):
            print("已跳过: 临时法规反射取点测量未通过")
            continue
        hidden_text = "（已隐藏）" if mirror_result.get("boundary_hidden") else ""
        print(f"边界: {mirror_result.get('boundary_name')}{hidden_text}")
        print(f"全部通过: {mirror_result.get('all_points_pass')}")
        for point_result in mirror_result.get("points", []):
            print(
                f"  {point_result.get('point_name')}: "
                f"{point_result.get('distance_to_boundary')} mm, "
                f"边界测量点={point_result.get('boundary_measure_point_coordinates')}, "
                f">=3mm={point_result.get('is_clearance_greater_equal_3mm')}"
            )

    print_section("二次校核")
    comparison_result = result.get("measurement_comparison", {})
    for mirror_side_name in ("左", "右"):
        mirror_result = comparison_result.get("mirrors", {}).get(mirror_side_name, {})
        print(f"\n{mirror_side_name}镜片:")
        print(f"最大距离差值: {mirror_result.get('max_abs_distance_difference')} mm")
        for point_result in mirror_result.get("points", []):
            print(
                f"  {point_result.get('point_name')}: "
                f"临时={point_result.get('temporary_distance_to_boundary')} mm, "
                f"CATIA={point_result.get('catia_distance_to_boundary')} mm, "
                f"差值={point_result.get('distance_difference')} mm"
            )

    print_section("法规视野截图")
    regulation_screenshot = result.get("screenshots", {}).get("法规视野", {})
    print(f"视点: {regulation_screenshot.get('target_point')}")
    print(f"视角方向: {regulation_screenshot.get('view_direction')}")
    print(f"上方向: {regulation_screenshot.get('up_direction')}")
    print(f"视距: {regulation_screenshot.get('view_distance')} mm")
    print(f"左投影点: {regulation_screenshot.get('left_projection_point')}")
    print(f"右法规点4: {regulation_screenshot.get('right_regulation_point4')}")
    if regulation_screenshot.get("success"):
        print(f"截图文件: {regulation_screenshot.get('saved_path')}")
    else:
        print(f"截图失败: {regulation_screenshot.get('error')}")

    print_section("后视镜截图")
    for mirror_side_name in ("左", "右"):
        screenshot_result = result.get("screenshots", {}).get(mirror_side_name, {})
        print(f"\n{mirror_side_name}后视镜:")
        print(f"视点: {screenshot_result.get('target_point')}")
        print(f"视角方向: {screenshot_result.get('view_direction')}")
        print(f"上方向: {screenshot_result.get('up_direction')}")
        print(f"视距: {screenshot_result.get('view_distance')} mm")
        print(f"截图时隐藏结构树: {screenshot_result.get('spec_tree_hidden_for_capture')}")
        print(f"截图后恢复结构树: {screenshot_result.get('spec_tree_restored_after_capture')}")
        if screenshot_result.get("success"):
            print(f"截图文件: {screenshot_result.get('saved_path')}")
        else:
            print(f"截图失败: {screenshot_result.get('error')}")


def _docx_text_run(text: Any, *, bold: bool = False, size: int = 22) -> str:
    """功能: 生成 WordprocessingML 文本 run；输入: 文本、是否加粗和字号；输出: XML 字符串。"""
    value = "" if text is None else str(text)
    properties = (
        '<w:rPr>'
        '<w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="Microsoft YaHei"/>'
        f'<w:sz w:val="{size}"/><w:szCs w:val="{size}"/>'
        f'{"<w:b/><w:bCs/>" if bold else ""}'
        '</w:rPr>'
    )
    parts = value.splitlines() or [""]
    runs = []
    for index, part in enumerate(parts):
        if index:
            runs.append("<w:br/>")
        runs.append(f'<w:t xml:space="preserve">{escape_xml(part)}</w:t>')
    return f'<w:r>{properties}{"".join(runs)}</w:r>'


def _docx_paragraph(
    text: Any,
    *,
    bold: bool = False,
    size: int = 22,
    style: str | None = None,
    spacing_after: int = 120,
    alignment: str | None = None,
) -> str:
    """功能: 生成 WordprocessingML 段落；输入: 文本、加粗、字号、样式、段后距和对齐；输出: XML 字符串。"""
    style_xml = f'<w:pStyle w:val="{style}"/>' if style else ""
    align_xml = f'<w:jc w:val="{alignment}"/>' if alignment else ""
    return (
        '<w:p>'
        f'<w:pPr>{style_xml}<w:spacing w:after="{spacing_after}"/>{align_xml}</w:pPr>'
        f'{_docx_text_run(text, bold=bold, size=size)}'
        '</w:p>'
    )


def _docx_table(rows: list[list[Any]]) -> str:
    """功能: 生成 WordprocessingML 表格；输入: 二维表格行数据；输出: XML 字符串。"""
    max_columns = max((len(row) for row in rows), default=3)
    if max_columns <= 3:
        column_widths = [2400, 3480, 3480]
    elif max_columns == 4:
        column_widths = [1000, 3300, 3700, 1360]
    else:
        column_widths = [max(900, int(9360 / max_columns)) for _ in range(max_columns)]
    row_xml = []
    for row_index, row in enumerate(rows):
        cells = []
        for column_index, width in enumerate(column_widths):
            value = row[column_index] if column_index < len(row) else ""
            fill = '<w:shd w:fill="EAF1FF"/>' if row_index == 0 else ""
            cells.append(
                '<w:tc>'
                f'<w:tcPr><w:tcW w:w="{width}" w:type="dxa"/>{fill}</w:tcPr>'
                f'{_docx_paragraph(value, bold=row_index == 0, spacing_after=0)}'
                '</w:tc>'
            )
        row_xml.append(f'<w:tr>{"".join(cells)}</w:tr>')
    grid = "".join(f'<w:gridCol w:w="{width}"/>' for width in column_widths)
    return (
        '<w:tbl>'
        '<w:tblPr>'
        '<w:tblW w:w="9360" w:type="dxa"/>'
        '<w:tblBorders>'
        '<w:top w:val="single" w:sz="4" w:color="D9E2F3"/>'
        '<w:left w:val="single" w:sz="4" w:color="D9E2F3"/>'
        '<w:bottom w:val="single" w:sz="4" w:color="D9E2F3"/>'
        '<w:right w:val="single" w:sz="4" w:color="D9E2F3"/>'
        '<w:insideH w:val="single" w:sz="4" w:color="D9E2F3"/>'
        '<w:insideV w:val="single" w:sz="4" w:color="D9E2F3"/>'
        '</w:tblBorders>'
        '<w:tblCellMar>'
        '<w:top w:w="120" w:type="dxa"/><w:left w:w="120" w:type="dxa"/>'
        '<w:bottom w:w="120" w:type="dxa"/><w:right w:w="120" w:type="dxa"/>'
        '</w:tblCellMar>'
        '</w:tblPr>'
        f'<w:tblGrid>{grid}</w:tblGrid>'
        f'{"".join(row_xml)}'
        '</w:tbl>'
    )


def _docx_image_paragraph(image: dict[str, Any]) -> str:
    """功能: 生成 Word 图片段落 XML；输入: 包含路径、关系ID、名称和尺寸的图片字典；输出: XML 字符串。"""
    width_emu = int(float(image.get("width_px", 600)) * 9525)
    height_emu = int(float(image.get("height_px", 338)) * 9525)
    rel_id = str(image["rel_id"])
    name = escape_xml(str(image.get("name", "截图")))
    image_id = int(image.get("id", 1))
    return (
        '<w:p><w:pPr><w:spacing w:after="240"/><w:jc w:val="center"/></w:pPr><w:r><w:drawing>'
        '<wp:inline distT="0" distB="0" distL="0" distR="0" '
        'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">'
        f'<wp:extent cx="{width_emu}" cy="{height_emu}"/>'
        f'<wp:docPr id="{image_id}" name="{name}"/>'
        '<wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1" '
        'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/>'
        '</wp:cNvGraphicFramePr>'
        '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
        '<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
        '<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">'
        '<pic:nvPicPr>'
        f'<pic:cNvPr id="{image_id}" name="{name}"/>'
        '<pic:cNvPicPr><a:picLocks noChangeAspect="1"/></pic:cNvPicPr>'
        '</pic:nvPicPr>'
        '<pic:blipFill>'
        f'<a:blip r:embed="{rel_id}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>'
        '<a:stretch><a:fillRect/></a:stretch>'
        '</pic:blipFill>'
        '<pic:spPr>'
        '<a:xfrm><a:off x="0" y="0"/>'
        f'<a:ext cx="{width_emu}" cy="{height_emu}"/></a:xfrm>'
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
        '</pic:spPr>'
        '</pic:pic>'
        '</a:graphicData>'
        '</a:graphic>'
        '</wp:inline>'
        '</w:drawing></w:r></w:p>'
    )


def _docx_regulation_vision_section() -> list[str]:
    """功能: 生成报告中的法规视野固定章节；输入: 无；输出: 段落 XML 列表。"""
    return [
        _docx_paragraph("法规视野", bold=True, size=28, spacing_after=120),
        _docx_paragraph("校核要求", bold=True, size=24, spacing_after=80),
        _docx_paragraph(
            "依据 GB 15084 和 ECE R046 法规，驾驶员视野需满足以下两项要求：",
            size=22,
            spacing_after=80,
        ),
        _docx_paragraph(
            "1.驾驶员至少应能看到宽度为 4 米的水平道路区域。该区域由平行于车辆垂直纵向中间平面、"
            "且通过驾驶员一侧车身最外点的平面所界定，并从驾驶员眼点后方 20 米处开始，"
            "一直延伸至地平线。",
            size=22,
            spacing_after=80,
        ),
        _docx_paragraph(
            "2.驾驶员还必须能看到宽度超过 1 米的路面区域。该区域由平行于车辆垂直纵向中间平面、"
            "且通过车辆最外点的平面所界定，并从通过驾驶员两眼点的垂面后方 4 米处开始延伸。",
            size=22,
            spacing_after=160,
        ),
    ]


def build_regulation_check_rows(regulation_vision_check: dict[str, Any]) -> list[list[Any]]:
    """功能: 将法规视野检查结果转换为报告表格行；输入: 法规视野检查结果；输出: 表格行列表。"""
    rows: list[list[Any]] = [["子项编号", "检查内容", "检查依据与测量值", "是否合格"]]
    for check in regulation_vision_check.get("checks", []):
        rows.append(
            [
                check.get("item_no", ""),
                check.get("content", ""),
                f"{check.get('basis', '')}\n{check.get('measure', '')}",
                "合格" if check.get("pass") else "不合格",
            ]
        )
    return rows


def write_simple_docx(
    path: Path,
    *,
    title: str,
    table_rows: list[list[Any]],
    regulation_check_rows: list[list[Any]] | None = None,
    regulation_image: dict[str, Any] | None = None,
    regulation_vision_image: dict[str, Any] | None = None,
    images: list[dict[str, Any]] | None = None,
) -> Path:
    """功能: 组装并写入简易 Word 报告；输入: 保存路径、标题、表格、法规图和截图；输出: 报告路径。"""
    images = images or []
    document_images = (
        ([regulation_image] if regulation_image else [])
        + ([regulation_vision_image] if regulation_vision_image else [])
        + images
    )
    relationships = [
        '<Relationship Id="rIdStyles" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" '
        'Target="styles.xml"/>',
        '<Relationship Id="rIdSettings" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" '
        'Target="settings.xml"/>',
    ]
    content_defaults = [
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>',
        '<Default Extension="xml" ContentType="application/xml"/>',
        '<Default Extension="png" ContentType="image/png"/>',
        '<Default Extension="jpg" ContentType="image/jpeg"/>',
        '<Default Extension="jpeg" ContentType="image/jpeg"/>',
        '<Default Extension="bmp" ContentType="image/bmp"/>',
    ]
    media_files: list[tuple[str, Path]] = []
    for index, image in enumerate(document_images, start=1):
        image_path = Path(image["path"])
        extension = image_path.suffix.lower().lstrip(".") or "png"
        if extension == "jpeg":
            extension = "jpg"
        media_name = f"image{index}.{extension}"
        image["rel_id"] = f"rIdImage{index}"
        image["id"] = index
        relationships.append(
            f'<Relationship Id="{image["rel_id"]}" '
            'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" '
            f'Target="media/{media_name}"/>'
        )
        media_files.append((media_name, image_path))

    body = [
        _docx_paragraph(title, bold=True, size=32, spacing_after=220),
        _docx_paragraph(f"生成时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}", size=20),
        *_docx_regulation_vision_section(),
        *(
            [
                _docx_paragraph(str(regulation_image.get("name", "法规图")), bold=True, size=24, spacing_after=80),
                _docx_image_paragraph(regulation_image),
            ]
            if regulation_image
            else []
        ),
        *(
            [
                _docx_paragraph(str(regulation_vision_image.get("name", "法规视野截图")), bold=True, size=24, spacing_after=80),
                _docx_image_paragraph(regulation_vision_image),
            ]
            if regulation_vision_image
            else []
        ),
        _docx_paragraph("法规视野检查结果", bold=True, size=26, spacing_after=120),
        _docx_table(regulation_check_rows or [["子项编号", "检查内容", "检查依据与测量值", "是否合格"]]),
    ]
    for image in images:
        body.append(_docx_paragraph(str(image.get("name", "截图")), bold=True, size=24, spacing_after=80))
        body.append(_docx_image_paragraph(image))
    body.extend(
        [
            _docx_paragraph("后视镜偏转结果", bold=True, size=26, spacing_after=120),
            _docx_table(table_rows),
        ]
    )

    document_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        f'<w:body>{"".join(body)}'
        '<w:sectPr>'
        '<w:pgSz w:w="12240" w:h="15840"/>'
        '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" '
        'w:header="720" w:footer="720" w:gutter="0"/>'
        '</w:sectPr>'
        '</w:body></w:document>'
    )

    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as document:
        document.writestr(
            "[Content_Types].xml",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            f'{"".join(content_defaults)}'
            '<Override PartName="/word/document.xml" '
            'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
            '<Override PartName="/word/styles.xml" '
            'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
            '<Override PartName="/word/settings.xml" '
            'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>'
            '</Types>',
        )
        document.writestr(
            "_rels/.rels",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" '
            'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" '
            'Target="word/document.xml"/></Relationships>',
        )
        document.writestr(
            "word/_rels/document.xml.rels",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            f'{"".join(relationships)}</Relationships>',
        )
        document.writestr(
            "word/styles.xml",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            '<w:style w:type="paragraph" w:default="1" w:styleId="Normal">'
            '<w:name w:val="Normal"/><w:rPr>'
            '<w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="Microsoft YaHei"/>'
            '<w:sz w:val="22"/><w:szCs w:val="22"/>'
            '</w:rPr></w:style></w:styles>',
        )
        document.writestr(
            "word/settings.xml",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>',
        )
        document.writestr("word/document.xml", document_xml)
        for media_name, image_path in media_files:
            document.writestr(f"word/media/{media_name}", image_path.read_bytes())
    return path


def generate_report_file(result: dict[str, Any], output_dir: Path) -> Path:
    """功能: 根据算法结果生成后视镜校核报告；输入: 算法结果字典和输出目录；输出: 报告路径。"""
    report_path = build_report_save_path(output_dir)
    screenshots = result.get("screenshots", {})
    regulation_vision_check = result.get("regulation_vision_check", {})
    temporary = result.get("temporary_reflection_points", {})
    temporary_mirrors = temporary.get("mirrors", {})
    left_temporary = temporary_mirrors.get("左", {})
    right_temporary = temporary_mirrors.get("右", {})
    regulation_image: dict[str, Any] | None = None
    if REGULATION_IMAGE_PATH.is_file():
        width_px, height_px = scaled_image_ratio(REGULATION_IMAGE_PATH, 0.63)
        regulation_image = {
            "path": str(REGULATION_IMAGE_PATH),
            "name": "后视镜法规图",
            "width_px": width_px,
            "height_px": height_px,
        }
    regulation_vision_image: dict[str, Any] | None = None
    regulation_vision_screenshot_path = screenshots.get("法规视野", {}).get("saved_path")
    if regulation_vision_screenshot_path and Path(regulation_vision_screenshot_path).is_file():
        width_px, height_px = scaled_image_size(regulation_vision_screenshot_path, 650)
        regulation_vision_image = {
            "path": str(regulation_vision_screenshot_path),
            "name": "法规视野截图",
            "width_px": width_px,
            "height_px": height_px,
        }

    report_rows: list[list[Any]] = [
        ["项目", "左镜片", "右镜片"],
        [
            "校核结果",
            format_calibration_status(left_temporary),
            format_calibration_status(right_temporary),
        ],
        [
            "偏转说明",
            format_angle_offset(left_temporary.get("horizontal_angle"), left_temporary.get("vertical_angle")),
            format_angle_offset(right_temporary.get("horizontal_angle"), right_temporary.get("vertical_angle")),
        ],
        [
            "水平偏移角(deg)",
            left_temporary.get("horizontal_angle"),
            right_temporary.get("horizontal_angle"),
        ],
        [
            "竖直偏移角(deg)",
            left_temporary.get("vertical_angle"),
            right_temporary.get("vertical_angle"),
        ],
        [
            "搜索组合数",
            left_temporary.get("searched_angle_count"),
            right_temporary.get("searched_angle_count"),
        ],
    ]
    screenshot_images: list[dict[str, Any]] = []
    left_screenshot_path = screenshots.get("左", {}).get("saved_path")
    right_screenshot_path = screenshots.get("右", {}).get("saved_path")
    image_width_px = 600
    if left_screenshot_path and Path(left_screenshot_path).is_file():
        width_px, height_px = scaled_image_size(left_screenshot_path, image_width_px)
        screenshot_images.append(
            {
                "path": str(left_screenshot_path),
                "name": "左后视镜截图",
                "width_px": width_px,
                "height_px": height_px,
            }
        )
    if right_screenshot_path and Path(right_screenshot_path).is_file():
        width_px, height_px = scaled_image_size(right_screenshot_path, image_width_px)
        screenshot_images.append(
            {
                "path": str(right_screenshot_path),
                "name": "右后视镜截图",
                "width_px": width_px,
                "height_px": height_px,
            }
        )

    return write_simple_docx(
        report_path,
        title="外后视镜视野校核报告",
        table_rows=report_rows,
        regulation_check_rows=build_regulation_check_rows(regulation_vision_check),
        regulation_image=regulation_image,
        regulation_vision_image=regulation_vision_image,
        images=screenshot_images,
    )


def run_rearview_analysis(
    read_file_path: str | Path,
    input_parameter_geo_set_name: str,
    regulation_line_geo_set_name: str,
    parametric_rearview_mirror_geo_set_name: str,
    regulation_reflection_point_geo_set_name: str,
    gap_check_geo_set_name: str,
    left_mirror_feature_name: str,
    right_mirror_feature_name: str,
    left_eye_point_feature_name: str,
    right_eye_point_feature_name: str,
    ground_feature_name: str,
    left_vehicle_width_line_feature_name: str,
    right_vehicle_width_line_feature_name: str,
) -> dict[str, Any]:
    """
    功能: 执行后视镜法规校核完整流程，供服务端直接调用；输入: CATPart 路径、几何图形集名称和输入特征名称；输出: 可 JSON 序列化的结果字典。

    所有输入均为文件路径或 CATIA 中需要查找、创建的对象名称。返回值可直接
    序列化为 JSON。失败时不会向外抛出异常，而是在返回值中提供 error。
    """
    try:
        output_dir = get_output_dir()
        print("[1] 连接 CATIA")
        catia = start_or_connect_catia()
        print("[2] 打开 CATPart")
        document, part = open_target_document(catia, read_file_path)

        print("[3] 读取后视镜 Part")
        print("[输入检查][info] 正在进行输入项检查")
        input_hybrid_body = checked_hybrid_body(
            part,
            "输入参数几何图形集名称",
            input_parameter_geo_set_name,
        )

        print("[4] 读取镜片与眼点输入")
        features = {}
        feature_errors: list[str] = []
        feature_checks = [
            ("left_mirror", "左镜片特征名称", left_mirror_feature_name),
            ("right_mirror", "右镜片特征名称", right_mirror_feature_name),
            ("left_eye", "左眼点特征名称", left_eye_point_feature_name),
            ("right_eye", "右眼点特征名称", right_eye_point_feature_name),
            ("ground", "地面特征名称", ground_feature_name),
            ("left_vehicle_width_line", "左车宽线特征名称", left_vehicle_width_line_feature_name),
            ("right_vehicle_width_line", "右车宽线特征名称", right_vehicle_width_line_feature_name),
        ]
        for key, label, feature_name in feature_checks:
            try:
                features[key] = checked_feature(input_hybrid_body, label, feature_name)
            except Exception as exc:
                feature_errors.append(str(exc))
        if feature_errors:
            raise LookupError("输入检查未通过: " + "；".join(feature_errors))

        print("[5] 确认车辆方向")
        directions = calculate_vehicle_directions(document, part, features)
        regulation_line_hybrid_body = create_hybrid_body(part, regulation_line_geo_set_name)
        eye_to_ground_projection = create_projection(
            part,
            regulation_line_hybrid_body,
            features["right_eye"],
            features["ground"],
            "眼点到地面投影点",
        )
        regulation_projection_references = create_regulation_projection_references(
            document=document,
            part=part,
            regulation_line_hybrid_body=regulation_line_hybrid_body,
            eye_to_ground_projection=eye_to_ground_projection,
            left_vehicle_width_line=features["left_vehicle_width_line"],
            right_vehicle_width_line=features["right_vehicle_width_line"],
            up_direction=directions["地面法线/上方向"],
            rough_rear_direction=directions["车辆粗略后方向"],
        )
        directions.update(regulation_projection_references["directions"])
        left_regulation_line = create_regulation_line(
            document=document,
            part=part,
            regulation_line_hybrid_body=regulation_line_hybrid_body,
            mirror=features["left_mirror"],
            eye=None,
            ground=features["ground"],
            vehicle_width_line=features["left_vehicle_width_line"],
            up_direction=directions["地面法线/上方向"],
            rear_direction=directions["车辆后方向"],
            side_direction=directions["车辆左方向"],
            side_name="左",
            eye_to_ground_projection=eye_to_ground_projection,
            ground_to_width_projection=regulation_projection_references["left_projection"],
            rear_reference_line=regulation_projection_references["left_extension_line"],
        )
        left_regulation_line["left_extension"] = regulation_projection_references["left_extension"]
        right_regulation_line = create_regulation_line(
            document=document,
            part=part,
            regulation_line_hybrid_body=regulation_line_hybrid_body,
            mirror=features["right_mirror"],
            eye=None,
            ground=features["ground"],
            vehicle_width_line=features["right_vehicle_width_line"],
            up_direction=directions["地面法线/上方向"],
            rear_direction=directions["车辆后方向"],
            side_direction=directions["车辆右方向"],
            side_name="右",
            eye_to_ground_projection=eye_to_ground_projection,
            ground_to_width_projection=regulation_projection_references["right_projection"],
            rear_reference_line=regulation_projection_references["right_extension_line"],
        )
        right_regulation_line["right_extension"] = regulation_projection_references["right_extension"]
        regulation_vision_check = check_regulation_vision_requirements(
            document=document,
            part=part,
            regulation_line_hybrid_body=regulation_line_hybrid_body,
            left_regulation_line=left_regulation_line,
            right_regulation_line=right_regulation_line,
            rear_direction=directions["车辆后方向"],
            up_direction=directions["地面法线/上方向"],
        )
        print("[6] 创建参数化后视镜")
        parametric_rearview_mirror = create_parametric_rearview_mirror(
            document=document,
            part=part,
            parametric_hybrid_body_name=parametric_rearview_mirror_geo_set_name,
            left_mirror=features["left_mirror"],
            right_mirror=features["right_mirror"],
            up_direction=directions["地面法线/上方向"],
        )
        parametric_rearview_hybrid_body = find_hybrid_body(
            part,
            parametric_rearview_mirror["geo_set_name"],
        )
        print("[7] 隐藏过程特征")
        print("[8] 法规线创建")
        print("[9] 法规反射取点")
        temporary_reflection_points = create_temporary_reflection_points(
            document=document,
            part=part,
            regulation_line_hybrid_body=regulation_line_hybrid_body,
            parametric_rearview_hybrid_body=parametric_rearview_hybrid_body,
            left_eye=features["left_eye"],
            right_eye=features["right_eye"],
        )
        for mirror_side_name in ("左", "右"):
            temporary_geo_set_name = (
                temporary_reflection_points.get("mirrors", {})
                .get(mirror_side_name, {})
                .get("geo_set_name")
            )
            if temporary_geo_set_name:
                hide_object(document, find_hybrid_body(part, temporary_geo_set_name))

        selected_mirror_sides = {
            mirror_side_name
            for mirror_side_name, mirror_result in temporary_reflection_points.get(
                "mirrors",
                {},
            ).items()
            if mirror_result.get("reflection_point_count", 0) > 0
        }
        rotated_left_mirror = find_feature(parametric_rearview_hybrid_body, "左镜片旋转")
        rotated_right_mirror = find_feature(parametric_rearview_hybrid_body, "右镜片旋转")

        regulation_reflection_points = create_regulation_reflection_points(
            document=document,
            part=part,
            reflection_point_hybrid_body_name=regulation_reflection_point_geo_set_name,
            parametric_rearview_hybrid_body=parametric_rearview_hybrid_body,
            regulation_line_hybrid_body=regulation_line_hybrid_body,
            left_mirror=rotated_left_mirror,
            right_mirror=rotated_right_mirror,
            left_eye=features["left_eye"],
            right_eye=features["right_eye"],
            selected_mirror_sides=selected_mirror_sides,
            temporary_reflection_points=temporary_reflection_points,
        )
        regulation_reflection_hybrid_body = find_hybrid_body(
            part,
            regulation_reflection_points["geo_set_name"],
        )
        print("[10] 间隙校验")
        gap_check = create_gap_check(
            document=document,
            part=part,
            gap_check_hybrid_body_name=gap_check_geo_set_name,
            left_mirror=rotated_left_mirror,
            right_mirror=rotated_right_mirror,
            reflection_point_hybrid_body=regulation_reflection_hybrid_body,
            reflection_points=regulation_reflection_points["reflection_points"],
            selected_mirror_sides=selected_mirror_sides,
        )
        print("[11] 二次校核")
        measurement_comparison = compare_gap_check_with_temporary_points(
            gap_check,
            temporary_reflection_points,
        )
        print("[12] 法规视野截图")
        screenshots = {
            "法规视野": capture_regulation_vision_screenshot(
                document=document,
                left_regulation_line=left_regulation_line,
                right_regulation_line=right_regulation_line,
                ground_up_direction=directions["地面法线/上方向"],
                vehicle_right_direction=directions["车辆右方向"],
                output_dir=output_dir,
                view_distance=REGULATION_VISION_SCREENSHOT_VIEW_DISTANCE,
            ),
        }
        print("[13] 结果另存")
        saved_as_path = build_result_save_path(read_file_path, output_dir)
        document.SaveAs(str(saved_as_path))
        print("[14] 创建标注装配")
        annotation_product_result = create_annotation_product_file(
            catia=catia,
            saved_catpart_path=saved_as_path,
            annotation_data=regulation_reflection_points.get(
                "best_view_distance_annotation_data",
                {},
            ),
            output_dir=output_dir,
            vehicle_right_direction=directions["车辆右方向"],
        )
        print("[15] 左/右后视镜截图")
        annotation_product_document = annotation_product_result.get("_product_document")
        try:
            annotation_product_document.Activate()
        except Exception:
            pass
        screenshots.update(
            {
                "左": capture_mirror_screenshot(
                    document=annotation_product_document or document,
                    part=part,
                    hybrid_body=regulation_reflection_hybrid_body,
                    mirror_feature=rotated_left_mirror,
                    mirror_side_name="左",
                    up_direction=directions["地面法线/上方向"],
                    output_dir=output_dir,
                    view_distance=SCREENSHOT_VIEW_DISTANCE,
                ),
                "右": capture_mirror_screenshot(
                    document=annotation_product_document or document,
                    part=part,
                    hybrid_body=regulation_reflection_hybrid_body,
                    mirror_feature=rotated_right_mirror,
                    mirror_side_name="右",
                    up_direction=directions["地面法线/上方向"],
                    output_dir=output_dir,
                    view_distance=SCREENSHOT_VIEW_DISTANCE,
                ),
            }
        )
        print("[16] 保存后视镜校核CATPart和标注装配")
        try:
            document.Save()
            print(f"后视镜校核 CATPart 已保存: {saved_as_path}", flush=True)
            annotation_product_document.Save()
            print(
                f"标注 CATProduct 已保存: {annotation_product_result['annotation_product_path']}",
                flush=True,
            )
        except Exception as exc:
            raise RuntimeError("左/右后视镜截图后保存后视镜校核 CATPart 或标注 CATProduct 失败。") from exc

        result = {
            "success": True,
            "document_path": str(Path(read_file_path).expanduser().resolve()),
            "output_dir": str(output_dir),
            "saved_as_path": str(saved_as_path),
            "annotation_product_path": annotation_product_result.get("annotation_product_path"),
            "annotation_part_path": annotation_product_result.get("annotation_part_path"),
            "annotation_product": {
                key: value
                for key, value in annotation_product_result.items()
                if not key.startswith("_")
            },
            "input_parameter_geo_set_name": str(input_hybrid_body.Name),
            "input_feature_names": {
                key: str(feature.Name)
                for key, feature in features.items()
            },
            "regulation_line_geo_set_name": str(regulation_line_hybrid_body.Name),
            "parametric_rearview_mirror_geo_set_name": parametric_rearview_mirror["geo_set_name"],
            "regulation_reflection_point_geo_set_name": regulation_reflection_points["geo_set_name"],
            "gap_check_geo_set_name": gap_check["geo_set_name"],
            "left_regulation_line": left_regulation_line,
            "right_regulation_line": right_regulation_line,
            "regulation_vision_check": regulation_vision_check,
            "parametric_rearview_mirror": parametric_rearview_mirror,
            "temporary_reflection_points": temporary_reflection_points,
            "regulation_reflection_points": regulation_reflection_points,
            "gap_check": gap_check,
            "measurement_comparison": measurement_comparison,
            "screenshots": screenshots,
            "directions": {
                direction_name: round_vector(vector)
                for direction_name, vector in directions.items()
            },
        }
        report_path = generate_report_file(result, output_dir)
        result["report_path"] = str(report_path)
        return result
    except Exception as exc:
        error_message = str(exc)
        return {
            "success": False,
            "error": error_message,
        }


def main() -> int:
    """功能: 命令行入口，按全局配置执行后视镜法规校核；输入: 无；输出: 进程退出码。"""
    configure_console_encoding()
    result = run_rearview_analysis(
        read_file_path=READ_FILE_PATH,
        input_parameter_geo_set_name=INPUT_PARAMETER_GEO_SET_NAME,
        regulation_line_geo_set_name=REGULATION_LINE_GEO_SET_NAME,
        parametric_rearview_mirror_geo_set_name=PARAMETRIC_REARVIEW_MIRROR_GEO_SET_NAME,
        regulation_reflection_point_geo_set_name=REGULATION_REFLECTION_POINT_GEO_SET_NAME,
        gap_check_geo_set_name=GAP_CHECK_GEO_SET_NAME,
        left_mirror_feature_name=LEFT_MIRROR_FEATURE_NAME,
        right_mirror_feature_name=RIGHT_MIRROR_FEATURE_NAME,
        left_eye_point_feature_name=LEFT_EYE_POINT_FEATURE_NAME,
        right_eye_point_feature_name=RIGHT_EYE_POINT_FEATURE_NAME,
        ground_feature_name=GROUND_FEATURE_NAME,
        left_vehicle_width_line_feature_name=LEFT_VEHICLE_WIDTH_LINE_FEATURE_NAME,
        right_vehicle_width_line_feature_name=RIGHT_VEHICLE_WIDTH_LINE_FEATURE_NAME,
    )
    print_result_summary(result)
    return 0 if result["success"] else 1


if __name__ == "__main__":
    main()
