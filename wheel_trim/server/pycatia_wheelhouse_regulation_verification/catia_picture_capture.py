"""
CATIA 通用截图工具。

核心流程:
1. 激活目标文档。
2. 通过 ActiveWindow.ActiveViewer.Viewpoint3D 设置视点、视线方向、上方向和视距。
3. 清理选择高亮并刷新视图。
4. 调用 ActiveViewer.CaptureToFile 导出图片。

说明:
- view_point 是旋转/观察中心，对应 CATIA Viewpoint3D Origin。
- sight_direction 是视线方向，对应 CATIA Viewpoint3D SightDirection。
- up_direction 是画面上方向，对应 CATIA Viewpoint3D UpDirection。
- view_distance 默认映射到 Viewpoint3D.FocusDistance，数值越大通常视图越远。
"""

from __future__ import annotations

import math
import time
from pathlib import Path
from typing import Any, Iterable

Vector = tuple[float, float, float]

CAPTURE_FORMATS = {
    "bmp": 0,
    "tiff": 1,
    "jpg": 2,
    "jpeg": 2,
    "emf": 3,
    "png": 4,
}

PRESET_VIEWS: dict[str, dict[str, Vector]] = {
    "back": {"sight_direction": (1.0, 0.0, 0.0), "up_direction": (0.0, 0.0, 1.0)},
    "front": {"sight_direction": (-1.0, 0.0, 0.0), "up_direction": (0.0, 0.0, 1.0)},
    "left": {"sight_direction": (0.0, 1.0, 0.0), "up_direction": (0.0, 0.0, 1.0)},
    "right": {"sight_direction": (0.0, -1.0, 0.0), "up_direction": (0.0, 0.0, 1.0)},
    "top": {"sight_direction": (0.0, 0.0, -1.0), "up_direction": (1.0, 0.0, 0.0)},
    "bottom": {"sight_direction": (0.0, 0.0, 1.0), "up_direction": (1.0, 0.0, 0.0)},
    "right_back_top_45": {"sight_direction": (1.0, -1.0, -1.0), "up_direction": (0.0, 0.0, 1.0)},
}


def _unwrap_com_object(target: Any) -> Any:
    """兼容 pycatia 包装对象和 win32com 原生 COM 对象。"""
    return getattr(target, "com_object", target)


def _as_vector(values: Iterable[float], label: str) -> Vector:
    """将输入转为三维向量。"""
    items = tuple(float(value) for value in values)
    if len(items) != 3:
        raise ValueError(f"{label} 必须是三维坐标/向量。")
    return items  # type: ignore[return-value]


def _vector_length(vector: Vector) -> float:
    return math.sqrt(sum(value * value for value in vector))


def _normalize_vector(vector: Vector, label: str) -> Vector:
    length = _vector_length(vector)
    if length <= 1.0e-12:
        raise ValueError(f"{label} 不能是零向量。")
    return tuple(value / length for value in vector)  # type: ignore[return-value]


def _dot(first: Vector, second: Vector) -> float:
    return sum(a * b for a, b in zip(first, second))


def _subtract(first: Vector, second: Vector) -> Vector:
    return tuple(a - b for a, b in zip(first, second))  # type: ignore[return-value]


def _scale(vector: Vector, factor: float) -> Vector:
    return tuple(value * factor for value in vector)  # type: ignore[return-value]


def _validated_directions(sight_direction: Iterable[float], up_direction: Iterable[float]) -> tuple[Vector, Vector]:
    """校验视线方向和上方向，并将上方向自动修正为垂直于视线方向。"""
    sight = _normalize_vector(_as_vector(sight_direction, "sight_direction"), "sight_direction")
    up = _normalize_vector(_as_vector(up_direction, "up_direction"), "up_direction")
    up_projection_on_sight = _scale(sight, _dot(up, sight))
    orthogonal_up = _subtract(up, up_projection_on_sight)
    if _vector_length(orthogonal_up) <= 1.0e-9:
        raise ValueError("sight_direction 和 up_direction 不能平行，无法自动修正上方向。")
    return sight, _normalize_vector(orthogonal_up, "orthogonal_up_direction")


def _get_attr(target: Any, *names: str) -> Any:
    """按多个候选名称读取属性，兼容 pycatia snake_case 和 COM PascalCase。"""
    for name in names:
        try:
            return getattr(target, name)
        except Exception:
            pass
    raise AttributeError(f"对象缺少属性: {', '.join(names)}")


def _call_vector_method(target: Any, names: tuple[str, ...], vector: Vector) -> None:
    """调用 CATIA 向量设置方法，兼容 tuple/list/展开参数。"""
    last_error: Exception | None = None
    for name in names:
        try:
            method = getattr(target, name)
        except Exception as exc:
            last_error = exc
            continue
        for args in ((vector,), (list(vector),), vector):
            try:
                if len(args) == 3 and not isinstance(args[0], (tuple, list)):
                    method(*args)
                else:
                    method(*args)
                return
            except Exception as exc:
                last_error = exc
    if last_error is not None:
        raise last_error
    raise AttributeError(f"对象缺少方法: {', '.join(names)}")


def _set_property(target: Any, names: tuple[str, ...], value: float) -> bool:
    """设置属性，成功返回 True。"""
    for name in names:
        try:
            setattr(target, name, value)
            return True
        except Exception:
            pass
    return False


def _get_catia_application(catia_or_caa: Any) -> Any:
    """获取 CATIA.Application COM 对象。"""
    catia = _unwrap_com_object(catia_or_caa)
    try:
        _ = catia.ActiveWindow
        return catia
    except Exception:
        pass
    try:
        return _unwrap_com_object(catia_or_caa.application)
    except Exception:
        return catia


def _get_active_viewer(catia_or_caa: Any) -> Any:
    """获取当前活动 Viewer。"""
    try:
        active_window = _get_attr(catia_or_caa, "active_window", "ActiveWindow")
        viewer = _get_attr(active_window, "active_viewer", "ActiveViewer")
        return _unwrap_com_object(viewer)
    except Exception:
        catia = _get_catia_application(catia_or_caa)
        return _unwrap_com_object(catia.ActiveWindow.ActiveViewer)


def _get_viewpoint_3d(viewer: Any) -> Any:
    """获取 Viewer3D 的 Viewpoint3D。"""
    try:
        return _get_attr(viewer, "Viewpoint3D", "viewpoint_3d")
    except Exception:
        try:
            from pycatia.in_interfaces.viewer_3d import Viewer3D

            return Viewer3D(viewer).viewpoint_3d
        except Exception as exc:
            raise RuntimeError("无法获取 CATIA Viewpoint3D。") from exc


def _safe_name(target: Any) -> str:
    """读取 CATIA 对象名称。"""
    try:
        return str(_unwrap_com_object(target).Name)
    except Exception:
        return ""


def _get_active_document(catia_or_caa: Any) -> Any | None:
    """获取当前活动文档。"""
    try:
        return _get_catia_application(catia_or_caa).ActiveDocument
    except Exception:
        return None


def _read_vector_method(target: Any, method_names: tuple[str, ...], property_names: tuple[str, ...]) -> Vector | None:
    """读取 CATIA 三维向量，兼容 GetXxx(outArray)、GetXxx() 和属性读取。"""
    for property_name in property_names:
        try:
            value = getattr(target, property_name)
            return _as_vector(value, property_name)
        except Exception:
            pass
    for method_name in method_names:
        try:
            method = getattr(target, method_name)
        except Exception:
            continue
        try:
            value = method()
            return _as_vector(value, method_name)
        except Exception:
            pass
        buffer = [0.0, 0.0, 0.0]
        try:
            method(buffer)
            return _as_vector(buffer, method_name)
        except Exception:
            pass
    return None


def _read_float_property(target: Any, names: tuple[str, ...]) -> float | None:
    """读取浮点属性。"""
    for name in names:
        try:
            return float(getattr(target, name))
        except Exception:
            pass
    return None


def get_viewpoint_state(catia_or_caa: Any) -> dict[str, Any]:
    """读取当前 CATIA 视图状态。"""
    viewer = _get_active_viewer(catia_or_caa)
    viewpoint = _get_viewpoint_3d(viewer)
    return {
        "view_point": _read_vector_method(viewpoint, ("GetOrigin", "get_origin"), ("Origin", "origin")),
        "sight_direction": _read_vector_method(
            viewpoint,
            ("GetSightDirection", "get_sight_direction"),
            ("SightDirection", "sight_direction"),
        ),
        "up_direction": _read_vector_method(
            viewpoint,
            ("GetUpDirection", "get_up_direction"),
            ("UpDirection", "up_direction"),
        ),
        "view_distance": _read_float_property(viewpoint, ("FocusDistance", "focus_distance", "Zoom", "zoom")),
    }


def get_document_root_object(document: Any) -> Any:
    """获取 CATProduct/CATPart 文档的根对象。"""
    document_com = _unwrap_com_object(document)
    for attr_name in ("Product", "Part"):
        try:
            return getattr(document_com, attr_name)
        except Exception:
            pass
    raise RuntimeError(f"无法获取文档根节点: {_safe_name(document_com) or '<unknown>'}")


def is_document_active(catia_or_caa: Any, document: Any) -> bool:
    """判断当前活动文档是否为指定文档。"""
    active_document = _get_active_document(catia_or_caa)
    if active_document is None:
        return False
    document_com = _unwrap_com_object(document)
    try:
        return active_document is document_com
    except Exception:
        pass
    active_name = _safe_name(active_document)
    target_name = _safe_name(document_com)
    return bool(active_name and target_name and active_name == target_name)


def activate_document(catia_or_caa: Any, document: Any | None = None) -> None:
    """激活指定 CATIA 文档；document 为空时保持当前活动文档。"""
    if document is None:
        return
    document_com = _unwrap_com_object(document)
    try:
        document_com.Activate()
        time.sleep(0.2)
        return
    except Exception:
        pass
    try:
        catia = _get_catia_application(catia_or_caa)
        catia.ActiveDocument = document_com
        time.sleep(0.2)
    except Exception:
        pass


def ensure_document_active(catia_or_caa: Any, document: Any | None = None) -> dict[str, Any]:
    """确保指定文档为活动文档，并返回切换结果。"""
    active_before = _get_active_document(catia_or_caa)
    if document is None:
        return {
            "status": "success",
            "changed": False,
            "active_before": _safe_name(active_before),
            "active_after": _safe_name(active_before),
        }
    document_com = _unwrap_com_object(document)
    if is_document_active(catia_or_caa, document_com):
        return {
            "status": "success",
            "changed": False,
            "active_before": _safe_name(active_before),
            "active_after": _safe_name(active_before),
            "target_document": _safe_name(document_com),
        }
    activate_document(catia_or_caa, document_com)
    active_after = _get_active_document(catia_or_caa)
    return {
        "status": "success" if is_document_active(catia_or_caa, document_com) else "unknown",
        "changed": True,
        "active_before": _safe_name(active_before),
        "active_after": _safe_name(active_after),
        "target_document": _safe_name(document_com),
    }


def clear_active_selection(catia_or_caa: Any) -> None:
    """清理当前活动文档选择集，避免截图里出现高亮。"""
    try:
        catia = _get_catia_application(catia_or_caa)
        catia.ActiveDocument.Selection.Clear()
    except Exception:
        pass


def reframe_document_root(
    catia_or_caa: Any,
    document: Any | None = None,
    *,
    clear_selection_after: bool = True,
    wait_seconds: float = 0.5,
) -> dict[str, Any]:
    """
    将文档根节点居中显示，并读取居中后的视点状态。

    CATProduct 会选择 Product 根节点，CATPart 会选择 Part 根节点，然后调用 ActiveViewer.Reframe()。
    """
    activation = ensure_document_active(catia_or_caa, document)
    catia = _get_catia_application(catia_or_caa)
    active_document = _get_active_document(catia_or_caa)
    if active_document is None:
        raise RuntimeError("无法获取当前活动文档，不能执行根节点居中。")
    root_object = get_document_root_object(active_document)
    selection = active_document.Selection
    viewer = _get_active_viewer(catia)
    selection.Clear()
    selection.Add(root_object)
    try:
        viewer.Reframe()
    except Exception:
        try:
            catia.StartCommand("Fit All In")
        except Exception as exc:
            raise RuntimeError("根节点居中失败: ActiveViewer.Reframe 和 Fit All In 均不可用。") from exc
    try:
        viewer.Update()
    except Exception:
        pass
    time.sleep(wait_seconds)
    state = get_viewpoint_state(catia)
    if clear_selection_after:
        try:
            selection.Clear()
        except Exception:
            pass
    return {
        "status": "success",
        "activation": activation,
        "root_name": _safe_name(root_object),
        "document_name": _safe_name(active_document),
        "view": state,
    }


def resolve_view_directions(
    *,
    sight_direction: Iterable[float] | None = None,
    up_direction: Iterable[float] | None = None,
    preset_view: str | None = None,
    fallback_sight_direction: Iterable[float] | None = None,
    fallback_up_direction: Iterable[float] | None = None,
) -> tuple[Vector, Vector, str]:
    """解析截图方向，支持预设视角、显式方向和当前视图方向回退。"""
    source = "explicit"
    base_sight = fallback_sight_direction
    base_up = fallback_up_direction
    if preset_view:
        preset_key = str(preset_view).strip().casefold()
        preset = PRESET_VIEWS.get(preset_key)
        if preset is None:
            raise ValueError(f"不支持的预设视图: {preset_view}")
        base_sight = preset["sight_direction"]
        base_up = preset["up_direction"]
        source = f"preset:{preset_key}"
    final_sight = sight_direction if sight_direction is not None else base_sight
    final_up = up_direction if up_direction is not None else base_up
    if final_sight is None:
        final_sight = (1.0, 0.0, 0.0)
        source = "default"
    if final_up is None:
        final_up = (0.0, 0.0, 1.0)
        source = "default"
    sight, up = _validated_directions(final_sight, final_up)
    return sight, up, source


def set_viewpoint(
    catia_or_caa: Any,
    view_point: Iterable[float],
    sight_direction: Iterable[float],
    up_direction: Iterable[float],
    view_distance: float | None = None,
    *,
    update: bool = True,
    settle_seconds: float = 0.3,
) -> dict[str, Any]:
    """
    设置当前 CATIA 视图。

    参数:
        catia_or_caa: CATIA.Application、pycatia caa 或其它带 com_object 的包装对象。
        view_point: 视点/旋转中心 (x, y, z)。
        sight_direction: 视角方向/视线方向 (x, y, z)。
        up_direction: 画面上方向 (x, y, z)。
            如果它和 sight_direction 不完全垂直，函数会自动去掉 up_direction
            在 sight_direction 上的投影分量，使最终上方向垂直于视角方向。
        view_distance: 视距，默认设置到 Viewpoint3D.FocusDistance。
        update: 是否刷新 Viewer。
        settle_seconds: 设置后等待时间。
    """
    origin = _as_vector(view_point, "view_point")
    sight, up = _validated_directions(sight_direction, up_direction)
    viewer = _get_active_viewer(catia_or_caa)
    viewpoint = _get_viewpoint_3d(viewer)

    _call_vector_method(viewpoint, ("PutOrigin", "put_origin"), origin)
    _call_vector_method(viewpoint, ("PutSightDirection", "put_sight_direction"), sight)
    _call_vector_method(viewpoint, ("PutUpDirection", "put_up_direction"), up)

    view_distance_set = False
    if view_distance is not None:
        view_distance_set = _set_property(viewpoint, ("FocusDistance", "focus_distance"), float(view_distance))
        if not view_distance_set:
            view_distance_set = _set_property(viewpoint, ("Zoom", "zoom"), float(view_distance))

    if update:
        try:
            viewer.Update()
        except Exception:
            pass
        time.sleep(settle_seconds)

    return {
        "view_point": origin,
        "sight_direction": sight,
        "up_direction": up,
        "view_distance": view_distance,
        "view_distance_set": view_distance_set,
    }


def set_preset_viewpoint(
    catia_or_caa: Any,
    view_point: Iterable[float],
    preset_view: str,
    view_distance: float | None = None,
    *,
    update: bool = True,
    settle_seconds: float = 0.3,
) -> dict[str, Any]:
    """按预设视图设置当前 CATIA 视角。"""
    preset_key = str(preset_view).strip().casefold()
    preset = PRESET_VIEWS.get(preset_key)
    if preset is None:
        raise ValueError(f"不支持的预设视图: {preset_view}")
    result = set_viewpoint(
        catia_or_caa,
        view_point,
        preset["sight_direction"],
        preset["up_direction"],
        view_distance,
        update=update,
        settle_seconds=settle_seconds,
    )
    result["preset_view"] = preset_key
    return result


def capture_active_view(
    catia_or_caa: Any,
    output_path: str | Path,
    *,
    image_format: str | None = None,
    clear_selection: bool = True,
    update: bool = True,
    wait_seconds: float = 0.5,
) -> str:
    """截取当前 CATIA 活动 Viewer 并保存为图片。"""
    path = Path(output_path)
    path.parent.mkdir(parents=True, exist_ok=True)
    fmt_key = (image_format or path.suffix.lstrip(".") or "png").casefold()
    fmt = CAPTURE_FORMATS.get(fmt_key)
    if fmt is None:
        raise ValueError(f"不支持的截图格式: {fmt_key}")

    viewer = _get_active_viewer(catia_or_caa)
    if clear_selection:
        clear_active_selection(catia_or_caa)
    if update:
        try:
            viewer.Update()
        except Exception:
            pass
    time.sleep(wait_seconds)
    viewer.CaptureToFile(fmt, str(path))
    return str(path)


def capture_viewpoint(
    catia_or_caa: Any,
    output_path: str | Path,
    *,
    view_point: Iterable[float],
    sight_direction: Iterable[float],
    up_direction: Iterable[float],
    view_distance: float | None = None,
    document: Any | None = None,
    image_format: str | None = None,
    clear_selection: bool = True,
    wait_seconds: float = 0.5,
) -> dict[str, Any]:
    """
    设置视图并截图。

    这是后续主程序最常用的入口:
        capture_viewpoint(catia, "out.png",
                          view_point=(0, 0, 0),
                          view_distance=0.01,
                          sight_direction=(1, 0, 0),
                          up_direction=(0, 0, 1))
    """
    activate_document(catia_or_caa, document)
    view_result = set_viewpoint(
        catia_or_caa,
        view_point,
        sight_direction,
        up_direction,
        view_distance,
        update=True,
        settle_seconds=wait_seconds,
    )
    screenshot_path = capture_active_view(
        catia_or_caa,
        output_path,
        image_format=image_format,
        clear_selection=clear_selection,
        update=True,
        wait_seconds=wait_seconds,
    )
    return {
        "status": "success",
        "screenshot_path": screenshot_path,
        "view": view_result,
    }


def capture_preset_viewpoint(
    catia_or_caa: Any,
    output_path: str | Path,
    *,
    view_point: Iterable[float],
    preset_view: str,
    view_distance: float | None = None,
    document: Any | None = None,
    image_format: str | None = None,
    clear_selection: bool = True,
    wait_seconds: float = 0.5,
) -> dict[str, Any]:
    """使用预设视角设置视图并截图。"""
    activate_document(catia_or_caa, document)
    view_result = set_preset_viewpoint(
        catia_or_caa,
        view_point,
        preset_view,
        view_distance,
        update=True,
        settle_seconds=wait_seconds,
    )
    screenshot_path = capture_active_view(
        catia_or_caa,
        output_path,
        image_format=image_format,
        clear_selection=clear_selection,
        update=True,
        wait_seconds=wait_seconds,
    )
    return {
        "status": "success",
        "screenshot_path": screenshot_path,
        "view": view_result,
    }


def capture_document_fixed_view(
    catia_or_caa: Any,
    document: Any,
    output_path: str | Path,
    *,
    view_point: Iterable[float],
    view_distance: float | None = None,
    sight_direction: Iterable[float] | None = None,
    up_direction: Iterable[float] | None = None,
    preset_view: str | None = None,
    image_format: str | None = None,
    clear_selection: bool = True,
    wait_seconds: float = 0.5,
) -> dict[str, Any]:
    """
    方式 1: 按固定视点/视角/视距/上方向截图。

    参数:
        document: 要截图的 CATProduct 或 CATPart 文档。
        view_point: 固定视点/旋转中心。
        sight_direction/up_direction: 固定视角方向和上方向。
        preset_view: 可选预设视角；提供后可不传 sight_direction/up_direction。
    """
    activation = ensure_document_active(catia_or_caa, document)
    sight, up, direction_source = resolve_view_directions(
        sight_direction=sight_direction,
        up_direction=up_direction,
        preset_view=preset_view,
    )
    view_result = set_viewpoint(
        catia_or_caa,
        view_point,
        sight,
        up,
        view_distance,
        update=True,
        settle_seconds=wait_seconds,
    )
    screenshot_path = capture_active_view(
        catia_or_caa,
        output_path,
        image_format=image_format,
        clear_selection=clear_selection,
        update=True,
        wait_seconds=wait_seconds,
    )
    return {
        "status": "success",
        "mode": "fixed",
        "activation": activation,
        "screenshot_path": screenshot_path,
        "direction_source": direction_source,
        "view": view_result,
    }


def capture_document_root_center_view(
    catia_or_caa: Any,
    document: Any,
    output_path: str | Path,
    *,
    view_distance: float | None = None,
    sight_direction: Iterable[float] | None = None,
    up_direction: Iterable[float] | None = None,
    preset_view: str | None = None,
    fallback_view_point: Iterable[float] | None = None,
    image_format: str | None = None,
    clear_selection: bool = True,
    wait_seconds: float = 0.5,
) -> dict[str, Any]:
    """
    方式 2: 先选择文档根节点并居中，再使用居中后的视点作为旋转中心截图。

    参数:
        document: 要截图的 CATProduct 或 CATPart 文档。
        view_distance: 基础缩放/视距；为空时不写入 FocusDistance，保留 Reframe 后的自适应缩放。
        preset_view: 可选视角类型，如 back/right/top。
        sight_direction/up_direction: 显式方向，可覆盖 preset 或当前视图方向。
        fallback_view_point: 如果 CATIA 无法读取居中后的 Viewpoint Origin，则使用该点。
    """
    reframe_result = reframe_document_root(
        catia_or_caa,
        document,
        clear_selection_after=clear_selection,
        wait_seconds=wait_seconds,
    )
    current_view = reframe_result.get("view") or {}
    view_point = current_view.get("view_point")
    if view_point is None:
        if fallback_view_point is None:
            raise RuntimeError("根节点居中后无法读取视点，请传入 fallback_view_point。")
        view_point = _as_vector(fallback_view_point, "fallback_view_point")
    sight, up, direction_source = resolve_view_directions(
        sight_direction=sight_direction,
        up_direction=up_direction,
        preset_view=preset_view,
        fallback_sight_direction=current_view.get("sight_direction"),
        fallback_up_direction=current_view.get("up_direction"),
    )
    view_result = set_viewpoint(
        catia_or_caa,
        view_point,
        sight,
        up,
        view_distance,
        update=True,
        settle_seconds=wait_seconds,
    )
    screenshot_path = capture_active_view(
        catia_or_caa,
        output_path,
        image_format=image_format,
        clear_selection=clear_selection,
        update=True,
        wait_seconds=wait_seconds,
    )
    return {
        "status": "success",
        "mode": "root_center",
        "reframe": reframe_result,
        "screenshot_path": screenshot_path,
        "direction_source": direction_source,
        "reframe_view_distance": current_view.get("view_distance"),
        "view_distance_preserved": view_distance is None,
        "view": view_result,
    }


def capture_document_view(
    catia_or_caa: Any,
    document: Any,
    output_path: str | Path,
    *,
    mode: str = "fixed",
    view_point: Iterable[float] | None = None,
    view_distance: float | None = None,
    sight_direction: Iterable[float] | None = None,
    up_direction: Iterable[float] | None = None,
    preset_view: str | None = None,
    fallback_view_point: Iterable[float] | None = None,
    image_format: str | None = None,
    clear_selection: bool = True,
    wait_seconds: float = 0.5,
) -> dict[str, Any]:
    """
    总入口: 根据 mode 选择截图方式。

    参数:
        catia_or_caa:
            CATIA 应用对象。可以传 win32com 的 CATIA.Application，也可以传 pycatia/caa 包装对象；
            函数内部会自动兼容 com_object。
        document:
            要截图的 CATIA 文档。CATProduct Document 和 CATPart Document 都可以。
            函数会先判断当前活动文档是不是它，不是则自动激活。
        output_path:
            截图输出路径。后缀通常为 .png/.jpg/.bmp/.tiff/.emf；
            如果父目录不存在会自动创建。
        mode:
            截图模式。
            fixed: 使用传入的固定 view_point 作为旋转中心。
            root_center: 先选择文档根节点并居中，再读取居中后的视点作为旋转中心。
            root 或 reframe 是 root_center 的别名。
        view_point:
            固定视点/旋转中心，三维坐标 (x, y, z)。
            mode="fixed" 时必填；mode="root_center" 时通常不用传。
        view_distance:
            视距/缩放值，优先写入 CATIA Viewpoint3D.FocusDistance。
            数值含义和 CATIA Viewer 缩放一致，一般越大视图越远。
            mode="root_center" 且不传时，不会改写视距，保留 Reframe 后 CATIA
            自动适配对象得到的缩放。
        sight_direction:
            视角方向/视线方向，三维向量 (x, y, z)。
            例如 (1, 0, 0) 表示沿 +X 方向看。
            如果传了 preset_view，可省略；如果同时传，会覆盖 preset_view 的视线方向。
        up_direction:
            画面上方向，三维向量 (x, y, z)。
            例如 (0, 0, 1) 表示屏幕上方对应模型 +Z。
            不能和 sight_direction 平行；如果不完全垂直，会自动正交化修正。
        preset_view:
            预设视角名称，可选 back/front/left/right/top/bottom/right_back_top_45。
            用于快速填充 sight_direction 和 up_direction。
        fallback_view_point:
            mode="root_center" 时的兜底视点。
            如果 CATIA 居中后无法读取 Viewpoint3D.Origin，就用这个点作为旋转中心。
        image_format:
            截图格式。为空时根据 output_path 后缀判断；无后缀时默认 png。
            支持 png/jpg/jpeg/bmp/tiff/emf。
        clear_selection:
            截图前是否清空选择集，默认 True，用于避免截图中出现选择高亮。
        wait_seconds:
            设置视角、刷新 Viewer 后等待的秒数。CATIA 刷新慢时可以调大。

    mode:
        fixed: 使用固定 view_point + 方向/预设视角 + view_distance 截图。
        root_center: 先根节点居中，读取当前视点作为旋转中心，再按方向/预设视角截图。

    返回:
        dict，成功时包含 status、mode、screenshot_path、view 等字段；
        view 中会记录最终使用的 view_point、sight_direction、up_direction、view_distance。
    """
    mode_key = str(mode).strip().casefold()
    if mode_key == "fixed":
        if view_point is None:
            raise ValueError("mode='fixed' 时必须传入 view_point。")
        return capture_document_fixed_view(
            catia_or_caa,
            document,
            output_path,
            view_point=view_point,
            view_distance=view_distance,
            sight_direction=sight_direction,
            up_direction=up_direction,
            preset_view=preset_view,
            image_format=image_format,
            clear_selection=clear_selection,
            wait_seconds=wait_seconds,
        )
    if mode_key in {"root_center", "root", "reframe"}:
        return capture_document_root_center_view(
            catia_or_caa,
            document,
            output_path,
            view_distance=view_distance,
            sight_direction=sight_direction,
            up_direction=up_direction,
            preset_view=preset_view,
            fallback_view_point=fallback_view_point,
            image_format=image_format,
            clear_selection=clear_selection,
            wait_seconds=wait_seconds,
        )
    raise ValueError(f"不支持的截图模式: {mode}")


def capture_document_screenshots(
    catia_or_caa: Any,
    document: Any,
    output_dir: str | Path,
    shots: list[dict[str, Any]],
    *,
    default_mode: str = "fixed",
    default_view_distance: float | None = None,
    image_format: str = "png",
    wait_seconds: float = 0.5,
) -> dict[str, Any]:
    """
    批量总入口: 对同一个文档连续截多张图。

    参数:
        catia_or_caa:
            CATIA 应用对象，含义同 capture_document_view。
        document:
            要截图的 CATProduct/CATPart 文档。
        output_dir:
            默认输出目录。单张截图未指定 output_path 时，会保存为 output_dir/label.image_format。
        shots:
            截图配置列表。每个元素都是一个 dict，常用字段包括:
            label: 结果键名和默认文件名。
            output_path: 单张截图输出路径；不传则使用 output_dir/label.image_format。
            mode: fixed 或 root_center；不传则使用 default_mode。
            view_point: fixed 模式下的固定视点/旋转中心。
            view_distance: 本张截图的视距/缩放值。
            sight_direction: 本张截图的视角方向。
            up_direction: 本张截图的画面上方向。
            preset_view: 本张截图的预设视角。
            fallback_view_point: root_center 模式读取居中视点失败时的兜底点。
            image_format: 本张截图格式；不传则使用批量入口的 image_format。
            clear_selection: 本张截图前是否清空选择集。
            wait_seconds: 本张截图设置视角后的等待秒数。
        default_mode:
            shots 中未写 mode 时使用的模式。
        default_view_distance:
            shots 中未写 view_distance 时使用的默认视距。
        image_format:
            默认截图格式。
        wait_seconds:
            默认等待秒数。

    shots 示例:
        [
            {
                "label": "front_fixed",
                "mode": "fixed",
                "view_point": (0, 0, 0),
                "preset_view": "front",
                "view_distance": 0.01,
            },
            {
                "label": "root_back",
                "mode": "root_center",
                "preset_view": "back",
                "view_distance": 0.02,
            },
        ]
    """
    output_root = Path(output_dir)
    results: dict[str, Any] = {}
    for index, shot in enumerate(shots, start=1):
        label = str(shot.get("label") or f"capture_{index:03d}")
        output_path = shot.get("output_path") or output_root / f"{label}.{image_format}"
        try:
            results[label] = capture_document_view(
                catia_or_caa,
                document,
                output_path,
                mode=str(shot.get("mode") or default_mode),
                view_point=shot.get("view_point"),
                view_distance=shot.get("view_distance", default_view_distance),
                sight_direction=shot.get("sight_direction"),
                up_direction=shot.get("up_direction"),
                preset_view=shot.get("preset_view"),
                fallback_view_point=shot.get("fallback_view_point"),
                image_format=shot.get("image_format") or image_format,
                clear_selection=bool(shot.get("clear_selection", True)),
                wait_seconds=float(shot.get("wait_seconds", wait_seconds)),
            )
        except Exception as exc:
            results[label] = {
                "status": "failed",
                "message": str(exc),
                "mode": str(shot.get("mode") or default_mode),
            }
    return results
