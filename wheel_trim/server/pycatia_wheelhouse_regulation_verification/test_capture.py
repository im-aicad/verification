"""
测试 CATIA 通用截图工具。

运行前准备:
1. 打开 CATIA。
2. 确保当前活动文档是包含 Front_Wheelhouse 的 CATProduct。
3. 运行本脚本。

测试内容:
1. 遍历结构树查找 Front_Wheelhouse，选择该对象 Reframe 居中后截图。
2. 固定坐标点截图暂时关闭，方便单独验证对象居中截图。
"""

from __future__ import annotations

from datetime import datetime
from pathlib import Path
from pprint import pprint
from typing import Any, Iterable

import win32com.client as win32

from catia_picture_capture import (
    capture_active_view,
    ensure_document_active,
    get_viewpoint_state,
    set_viewpoint,
)


FRONT_WHEELHOUSE_KEYWORD = "Front_Wheelhouse"
SIGHT_DIRECTION = (0.0, 0.0, -1.0)
UP_DIRECTION = (-1.0, 0.0, 0.0)
VIEW_DISTANCE = 10000.0
POINT_CENTER = (2950.0, -775.0, 519.0)


def iter_collection(collection: Any) -> Iterable[Any]:
    """遍历 CATIA COM Collection。"""
    try:
        count = int(collection.Count)
    except Exception:
        return
    for index in range(1, count + 1):
        try:
            yield collection.Item(index)
        except Exception:
            continue


def safe_name(target: Any, attr_name: str = "Name") -> str:
    """安全读取 CATIA 对象名称。"""
    try:
        return str(getattr(target, attr_name))
    except Exception:
        return ""


def product_text(product: Any) -> str:
    """拼接 Product 可搜索文本。"""
    return " ".join(
        text
        for text in (
            safe_name(product, "Name"),
            safe_name(product, "PartNumber"),
        )
        if text
    )


def find_product_by_keyword(root_product: Any, keyword: str) -> Any | None:
    """递归遍历 Product 结构树，按名称或 PartNumber 查找组件。"""
    keyword_lower = keyword.casefold()
    if keyword_lower in product_text(root_product).casefold():
        return root_product
    try:
        children = root_product.Products
    except Exception:
        return None
    for child in iter_collection(children):
        matched = find_product_by_keyword(child, keyword)
        if matched is not None:
            return matched
    return None


def reframe_product_component(product_document: Any, product_component: Any) -> dict[str, Any]:
    """选择指定 Product 组件并执行 Reframe，返回居中后的视图状态。"""
    product_document.Activate()
    selection = product_document.Selection
    viewer = product_document.Application.ActiveWindow.ActiveViewer
    try:
        selection.Clear()
        selection.Add(product_component)
        viewer.Reframe()
        viewer.Update()
        view_state = get_viewpoint_state(product_document.Application)
    finally:
        try:
            selection.Clear()
        except Exception:
            pass
    return {
        "component_name": safe_name(product_component, "Name"),
        "component_part_number": safe_name(product_component, "PartNumber"),
        "view": view_state,
    }


def desktop_output_path(label: str, timestamp: str) -> Path:
    """构造桌面截图路径。"""
    desktop = Path.home() / "Desktop"
    return desktop / f"result_{label}_{timestamp}.png"


def main() -> dict[str, Any]:
    """执行对象居中截图。"""
    catia = win32.GetActiveObject("CATIA.Application")
    product_document = catia.ActiveDocument
    ensure_document_active(catia, product_document)
    root_product = product_document.Product
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

    front_wheelhouse = find_product_by_keyword(root_product, FRONT_WHEELHOUSE_KEYWORD)
    if front_wheelhouse is None:
        raise RuntimeError(f"未在当前 CATProduct 结构树中找到组件: {FRONT_WHEELHOUSE_KEYWORD}")

    object_reframe = reframe_product_component(product_document, front_wheelhouse)
    object_view_point = object_reframe.get("view", {}).get("view_point")
    if object_view_point is None:
        raise RuntimeError("Front_Wheelhouse 居中后无法读取 Viewpoint3D Origin。")
    set_viewpoint(
        catia,
        object_view_point,
        SIGHT_DIRECTION,
        UP_DIRECTION,
        None,
    )
    object_path = capture_active_view(
        catia,
        desktop_output_path("front_wheelhouse_center", timestamp),
        image_format="png",
    )

    set_viewpoint(
        catia,
        POINT_CENTER,
        SIGHT_DIRECTION,
        UP_DIRECTION,
        VIEW_DISTANCE,
    )
    point_path = capture_active_view(
        catia,
        desktop_output_path("point_center", timestamp),
        image_format="png",
    )

    result = {
        "status": "success",
        "active_document": safe_name(product_document),
        "front_wheelhouse": {
            "name": object_reframe.get("component_name"),
            "part_number": object_reframe.get("component_part_number"),
            "reframe_view_point": object_view_point,
        },
        "settings": {
            "sight_direction": SIGHT_DIRECTION,
            "up_direction": UP_DIRECTION,
            "view_distance": "preserve_reframe_distance",
            "point_center": POINT_CENTER,
        },
        "screenshots": {
            "front_wheelhouse_center": object_path,
            # "point_center": point_path,
        },
    }
    pprint(result)
    return result


if __name__ == "__main__":
    main()
