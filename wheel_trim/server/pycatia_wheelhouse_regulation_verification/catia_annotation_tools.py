"""
CATIA 通用距离标注工具。

核心流程:
1. 确认当前活动文档是 CATProduct。
2. 使用传入的过程 Part，或在当前 CATProduct 中新建一个过程 Part。
3. 在过程 Part 的指定几何图形集中创建两个点和一条连线。
4. 在 CATProduct 的 Marker3Ds 中创建 3D 文本标记。

典型调用:
    result = create_distance_annotation(
        catia,
        point1=(0, 0, 0),
        point2=(100, 0, 0),
        text_offset_direction=(0, 0, 1),
        text_offset_distance=20,
        annotation_text=None,  # None 时自动使用距离 + mm
    )
"""

from __future__ import annotations

import math
import time
from pathlib import Path
from typing import Any, Iterable

Vector = tuple[float, float, float]

DEFAULT_PROCESS_PART_NUMBER = "CATIA_Distance_Annotation_Process"
DEFAULT_PROCESS_PART_NAME = "距离标注过程Part"
DEFAULT_HYBRID_BODY_NAME = "距离标注"
DEFAULT_FEATURE_COLOR = (0, 255, 0)
DEFAULT_TEXT_COLOR = (0, 255, 0)
DEFAULT_LINE_WIDTH = 2
DEFAULT_TEXT_SIZE = 8.0
CATPART_SUFFIX = ".CATPart"


def _unwrap_com_object(target: Any) -> Any:
    """兼容 pycatia 包装对象和 win32com 原生 COM 对象。"""
    return getattr(target, "com_object", target)


def _get_catia_application(catia_or_caa: Any | None = None) -> Any:
    """获取 CATIA.Application COM 对象。"""
    if catia_or_caa is not None:
        catia = _unwrap_com_object(catia_or_caa)
        try:
            _ = catia.ActiveDocument
            return catia
        except Exception:
            pass
        try:
            return _unwrap_com_object(catia_or_caa.application)
        except Exception:
            pass
    try:
        import win32com.client as win32

        return win32.GetActiveObject("CATIA.Application")
    except Exception as exc:
        raise RuntimeError("无法获取 CATIA.Application，请传入 catia 对象或确认 CATIA 已启动。") from exc


def _safe_name(target: Any) -> str:
    try:
        return str(_unwrap_com_object(target).Name)
    except Exception:
        return ""


def _as_vector(values: Iterable[float], label: str) -> Vector:
    items = tuple(float(value) for value in values)
    if len(items) != 3:
        raise ValueError(f"{label} 必须是三维坐标/向量。")
    return items  # type: ignore[return-value]


def _add(first: Vector, second: Vector) -> Vector:
    return tuple(a + b for a, b in zip(first, second))  # type: ignore[return-value]


def _subtract(first: Vector, second: Vector) -> Vector:
    return tuple(a - b for a, b in zip(first, second))  # type: ignore[return-value]


def _scale(vector: Vector, factor: float) -> Vector:
    return tuple(value * factor for value in vector)  # type: ignore[return-value]


def _length(vector: Vector) -> float:
    return math.sqrt(sum(value * value for value in vector))


def _normalize(vector: Iterable[float], label: str) -> Vector:
    raw = _as_vector(vector, label)
    length = _length(raw)
    if length <= 1.0e-12:
        raise ValueError(f"{label} 不能是零向量。")
    return tuple(value / length for value in raw)  # type: ignore[return-value]


def _average(first: Vector, second: Vector) -> Vector:
    return tuple((a + b) / 2.0 for a, b in zip(first, second))  # type: ignore[return-value]


def _round_vector(vector: Vector, digits: int = 6) -> tuple[float, float, float]:
    return tuple(round(value, digits) for value in vector)  # type: ignore[return-value]


def _iter_collection(collection: Any) -> Iterable[Any]:
    try:
        count = int(collection.Count)
    except Exception:
        return
    for index in range(1, count + 1):
        try:
            yield collection.Item(index)
        except Exception:
            continue


def _is_product_document(document: Any) -> bool:
    try:
        _ = document.Product
        return True
    except Exception:
        return False


def get_active_product_document(catia_or_caa: Any | None = None) -> tuple[Any, Any]:
    """获取当前活动 CATProduct 文档和根 Product。"""
    catia = _get_catia_application(catia_or_caa)
    try:
        document = catia.ActiveDocument
    except Exception as exc:
        raise RuntimeError("无法获取当前活动文档。") from exc
    if not _is_product_document(document):
        raise RuntimeError(
            f"当前活动文档不是 CATProduct，无法创建 Product 级 3D 文本标注: {_safe_name(document)}"
        )
    return document, document.Product


def get_or_create_hybrid_body(part: Any, body_name: str) -> Any:
    """获取或创建 Part 下的几何图形集。"""
    try:
        hybrid_bodies = part.HybridBodies
    except Exception as exc:
        raise RuntimeError("过程 Part 不支持 HybridBodies，无法创建标注几何。") from exc
    for body in _iter_collection(hybrid_bodies):
        if _safe_name(body) == body_name:
            return body
    body = hybrid_bodies.Add()
    try:
        body.Name = body_name
    except Exception:
        pass
    try:
        part.InWorkObject = body
    except Exception:
        pass
    return body


def create_reference(part: Any, feature: Any) -> Any:
    """创建 CATIA Reference。"""
    return part.CreateReferenceFromObject(feature)


def set_object_visual_style(
    document: Any,
    feature: Any,
    color: tuple[int, int, int] = DEFAULT_FEATURE_COLOR,
    *,
    width: int | None = None,
) -> None:
    """设置对象颜色和可选线宽。"""
    selection = document.Selection
    try:
        selection.Clear()
        selection.Add(feature)
        selection.VisProperties.SetRealColor(int(color[0]), int(color[1]), int(color[2]), 1)
        if width is not None:
            try:
                selection.VisProperties.SetRealWidth(int(width), 1)
            except Exception:
                pass
    finally:
        try:
            selection.Clear()
        except Exception:
            pass


def _set_marker_text_size(marker: Any, text_size: float) -> dict[str, Any]:
    """优先通过 pycatia wrapper 属性设置 Marker3D 字号，再回退到 COM 属性。"""
    requested = float(text_size)
    applied_via = None
    error = None
    readback = None
    readback_error = None
    for attr_name, setter in (
        ("text_size", lambda obj, value: setattr(obj, "text_size", value)),
        ("TextSize", lambda obj, value: setattr(obj, "TextSize", value)),
        ("com.TextSize", lambda obj, value: setattr(_unwrap_com_object(obj), "TextSize", value)),
    ):
        try:
            setter(marker, requested)
            applied_via = attr_name
            break
        except Exception as exc:
            error = str(exc)
    try:
        readback = float(getattr(marker, "text_size"))
    except Exception:
        try:
            readback = float(getattr(marker, "TextSize"))
        except Exception as exc:
            readback_error = str(exc)
    return {
        "text_size_requested": requested,
        "text_size_applied": applied_via is not None,
        "text_size_applied_via": applied_via,
        "text_size_readback": readback,
        "text_size_readback_error": readback_error,
        "text_size_error": error,
    }


def _refresh_marker_display(product_document: Any, root_product: Any, marker: Any) -> None:
    """尽量强制 Marker3D 和当前视图刷新。"""
    try:
        marker.Update()
    except Exception:
        try:
            _unwrap_com_object(marker).Update()
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
        viewer.Update()
        try:
            viewer.Reframe()
        except Exception:
            pass
        viewer.Update()
    except Exception:
        pass


def _safe_filename(name: str, fallback: str = "CATIA_Distance_Annotation_Process") -> str:
    cleaned = "".join(
        char if char not in r'<>:"/\|?*' and ord(char) >= 32 else "_"
        for char in str(name or "").strip()
    ).strip(" .")
    return cleaned or fallback


def _normalize_catpart_path(path: str | Path | None, product_document: Any, part_number: str) -> Path | None:
    if path is None:
        return None
    save_path = Path(path).expanduser()
    if save_path.suffix.lower() != CATPART_SUFFIX.lower():
        if save_path.suffix:
            save_path = save_path.with_suffix(CATPART_SUFFIX)
        else:
            save_path = save_path / f"{_safe_filename(part_number)}{CATPART_SUFFIX}" if save_path.exists() and save_path.is_dir() else save_path.with_suffix(CATPART_SUFFIX)
    if not save_path.is_absolute():
        base_dir: Path
        try:
            product_full_name = str(product_document.FullName or "").strip()
        except Exception:
            product_full_name = ""
        if product_full_name:
            base_dir = Path(product_full_name).expanduser().resolve().parent
        else:
            base_dir = Path.cwd()
        save_path = base_dir / save_path
    return save_path.resolve()


def _same_resolved_path(first: str | Path | None, second: str | Path | None) -> bool:
    if not first or not second:
        return False
    try:
        return Path(first).expanduser().resolve() == Path(second).expanduser().resolve()
    except Exception:
        return str(first).strip().lower() == str(second).strip().lower()


def add_component_from_file_to_product(product_document: Any, root_product: Any, file_path: str | Path) -> Any:
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
        product_document.Application.SystemService.Evaluate(
            vba_code,
            0,
            "add_component",
            [root_product, str(file_path)],
        )
    except Exception as exc:
        raise RuntimeError(f"无法装配标注过程 Part: {file_path}") from exc
    try:
        after_count = int(root_product.Products.Count)
        if after_count > before_count:
            return root_product.Products.Item(after_count)
    except Exception:
        pass
    raise RuntimeError(f"标注过程 Part 已尝试装配，但无法定位新增组件: {file_path}")


def _part_document_from_component(component: Any) -> tuple[Any | None, Any | None]:
    attempts = (
        lambda: component.ReferenceProduct.Parent,
        lambda: component.Parent,
    )
    for attempt in attempts:
        try:
            part_document = attempt()
            return part_document, part_document.Part
        except Exception:
            continue
    return None, None


def create_saved_process_part_and_add_to_product(
    product_document: Any,
    root_product: Any,
    *,
    save_path: str | Path,
    part_number: str = DEFAULT_PROCESS_PART_NUMBER,
    part_name: str = DEFAULT_PROCESS_PART_NAME,
) -> dict[str, Any]:
    save_path = Path(save_path).expanduser().resolve()
    save_path.parent.mkdir(parents=True, exist_ok=True)
    catia = product_document.Application
    temporary_document = None
    try:
        temporary_document = catia.Documents.Add("Part")
        temporary_part = temporary_document.Part
        try:
            temporary_part.PartNumber = part_number
        except Exception:
            pass
        try:
            temporary_part.Name = part_name
        except Exception:
            pass
        try:
            temporary_document.Product.PartNumber = part_number
            temporary_document.Product.Name = part_name
        except Exception:
            pass
        try:
            temporary_part.Update()
        except Exception:
            pass
        temporary_document.SaveAs(str(save_path))
    except Exception as exc:
        raise RuntimeError(f"无法创建并保存标注过程 CATPart: {save_path}") from exc
    finally:
        if temporary_document is not None:
            try:
                temporary_document.Close()
            except Exception:
                pass

    try:
        product_document.Activate()
    except Exception:
        pass
    component = add_component_from_file_to_product(product_document, root_product, save_path)
    try:
        component.PartNumber = part_number
    except Exception:
        pass
    try:
        component.Name = part_name
    except Exception:
        pass
    part_document, part = _part_document_from_component(component)
    if part_document is None or part is None:
        try:
            part_document = product_document.Application.Documents.Open(str(save_path))
            part = part_document.Part
        except Exception as exc:
            raise RuntimeError(f"标注过程 CATPart 已装配，但无法获取 PartDocument: {save_path}") from exc
    return {
        "component": component,
        "part_document": part_document,
        "part": part,
        "created": True,
        "external_file": True,
        "save_path": str(save_path),
        "initial_save_result": {
            "status": "success",
            "path": str(save_path),
            "message": "标注过程 CATPart 已预创建并装配。",
        },
    }


def add_new_process_part_to_product(
    product_document: Any,
    root_product: Any,
    *,
    part_number: str = DEFAULT_PROCESS_PART_NUMBER,
    part_name: str = DEFAULT_PROCESS_PART_NAME,
) -> dict[str, Any]:
    """在当前 CATProduct 中新建一个用于承载标注几何的过程 Part。"""
    try:
        component = root_product.Products.AddNewComponent("Part", part_number)
    except Exception as exc:
        raise RuntimeError("无法在 CATProduct 中新建标注过程 Part。") from exc
    try:
        component.Name = part_name
    except Exception:
        pass
    part_document, part = _part_document_from_component(component)
    if part is None or part_document is None:
        raise RuntimeError("已新建过程 Part，但无法获取 PartDocument/Part。")
    try:
        part.PartNumber = part_number
    except Exception:
        pass
    try:
        part_document.Product.PartNumber = part_number
        part_document.Product.Name = part_name
    except Exception:
        pass
    return {
        "component": component,
        "part_document": part_document,
        "part": part,
        "created": True,
    }


def resolve_process_part(
    product_document: Any,
    root_product: Any,
    *,
    process_part: Any | None = None,
    process_part_document: Any | None = None,
    process_component: Any | None = None,
    create_if_missing: bool = True,
    process_part_number: str = DEFAULT_PROCESS_PART_NUMBER,
    process_part_name: str = DEFAULT_PROCESS_PART_NAME,
    process_part_save_path: str | Path | None = None,
) -> dict[str, Any]:
    """解析或创建过程 Part。"""
    if process_part is not None:
        normalized_save_path = _normalize_catpart_path(
            process_part_save_path,
            product_document,
            process_part_number,
        )
        return {
            "component": process_component,
            "part_document": process_part_document,
            "part": process_part,
            "created": False,
            "external_file": normalized_save_path is not None,
            "save_path": str(normalized_save_path) if normalized_save_path is not None else None,
        }
    if process_component is not None:
        part_document, part = _part_document_from_component(process_component)
        if part_document is not None and part is not None:
            normalized_save_path = _normalize_catpart_path(
                process_part_save_path,
                product_document,
                process_part_number,
            )
            return {
                "component": process_component,
                "part_document": part_document,
                "part": part,
                "created": False,
                "external_file": normalized_save_path is not None,
                "save_path": str(normalized_save_path) if normalized_save_path is not None else None,
            }
    if not create_if_missing:
        raise RuntimeError("未传入过程 Part，且 create_if_missing=False。")
    normalized_save_path = _normalize_catpart_path(
        process_part_save_path,
        product_document,
        process_part_number,
    )
    if normalized_save_path is not None:
        return create_saved_process_part_and_add_to_product(
            product_document,
            root_product,
            save_path=normalized_save_path,
            part_number=process_part_number,
            part_name=process_part_name,
        )
    return add_new_process_part_to_product(
        product_document,
        root_product,
        part_number=process_part_number,
        part_name=process_part_name,
    )


def create_annotation_line_geometry(
    part_document: Any,
    part: Any,
    hybrid_body: Any,
    point1: Vector,
    point2: Vector,
    *,
    annotation_name: str,
    feature_color: tuple[int, int, int] = DEFAULT_FEATURE_COLOR,
    line_width: int = DEFAULT_LINE_WIDTH,
) -> dict[str, Any]:
    """在过程 Part 中创建两个点和一条标注线。"""
    factory = part.HybridShapeFactory
    point_feature1 = factory.AddNewPointCoord(*point1)
    point_feature2 = factory.AddNewPointCoord(*point2)
    try:
        point_feature1.Name = f"{annotation_name}_P1"
        point_feature2.Name = f"{annotation_name}_P2"
    except Exception:
        pass
    hybrid_body.AppendHybridShape(point_feature1)
    hybrid_body.AppendHybridShape(point_feature2)
    try:
        part.Update()
    except Exception:
        pass
    line_feature = factory.AddNewLinePtPt(
        create_reference(part, point_feature1),
        create_reference(part, point_feature2),
    )
    try:
        line_feature.Name = f"{annotation_name}_Line"
    except Exception:
        pass
    hybrid_body.AppendHybridShape(line_feature)
    try:
        part.UpdateObject(line_feature)
    except Exception:
        try:
            part.Update()
        except Exception:
            pass
    for feature in (point_feature1, point_feature2):
        try:
            set_object_visual_style(part_document, feature, feature_color)
        except Exception:
            pass
    try:
        set_object_visual_style(part_document, line_feature, feature_color, width=line_width)
    except Exception:
        pass
    return {
        "point1_feature": point_feature1,
        "point2_feature": point_feature2,
        "line_feature": line_feature,
        "point1_name": _safe_name(point_feature1),
        "point2_name": _safe_name(point_feature2),
        "line_name": _safe_name(line_feature),
    }


def get_marker3ds(root_product: Any) -> Any | None:
    """获取 Product 的 3D 标记集合。"""
    try:
        return root_product.GetTechnologicalObject("Marker3Ds")
    except Exception:
        return None


def activate_feature_for_display_refresh(document: Any, feature: Any) -> None:
    """通过选择对象触发 CATIA 显示刷新。"""
    try:
        document.Activate()
    except Exception:
        pass
    selection = document.Selection
    try:
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
    finally:
        try:
            selection.Clear()
        except Exception:
            pass


def refresh_product_view(product_document: Any, root_product: Any | None = None) -> None:
    """刷新 Product 更新和当前视图。"""
    if root_product is None:
        try:
            root_product = product_document.Product
        except Exception:
            root_product = None
    try:
        if root_product is not None:
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


def reopen_product_document_for_marker_refresh(
    catia_or_caa: Any | None = None,
    product_document: Any | None = None,
    *,
    product_save_path: str | Path | None = None,
    close_before_open: bool = True,
    wait_seconds: float = 0.2,
) -> dict[str, Any]:
    """
    保存、关闭并重新打开 CATProduct，用于强制刷新 Product 级 Marker3Ds 显示缓存。

    单条标注可在创建完成后调用一次；批量标注应在全部标注完成后只调用一次。
    """
    catia = _get_catia_application(catia_or_caa)
    if product_document is None:
        product_document, _root_product = get_active_product_document(catia)
    else:
        product_document = _unwrap_com_object(product_document)
    try:
        product_document.Activate()
    except Exception:
        pass

    document_name = _safe_name(product_document)
    save_path_text = ""
    if product_save_path is not None:
        save_path = Path(product_save_path).expanduser().resolve()
        save_path.parent.mkdir(parents=True, exist_ok=True)
        save_path_text = str(save_path)
        try:
            current_full_name = str(product_document.FullName or "").strip()
        except Exception:
            current_full_name = ""
        try:
            if _same_resolved_path(current_full_name, save_path_text):
                product_document.Save()
            elif save_path.exists():
                return {
                    "status": "failed",
                    "message": (
                        "目标 CATProduct 已存在且当前文档不是该路径，"
                        "为避免 CATIA 覆盖确认弹窗，已跳过 SaveAs。"
                    ),
                    "document": document_name,
                    "path": save_path_text,
                    "current_full_name": current_full_name,
                    "_product_document": product_document,
                    "_root_product": getattr(product_document, "Product", None),
                }
            else:
                product_document.SaveAs(save_path_text)
        except Exception as exc:
            return {
                "status": "failed",
                "message": f"CATProduct 保存失败: {exc}",
                "document": document_name,
                "path": save_path_text,
                "current_full_name": current_full_name,
                "_product_document": product_document,
                "_root_product": getattr(product_document, "Product", None),
            }
    else:
        try:
            save_path_text = str(product_document.FullName or "").strip()
        except Exception:
            save_path_text = ""
        if not save_path_text:
            return {
                "status": "skipped",
                "message": "CATProduct 尚未保存且未提供 product_save_path，无法关闭重开刷新 Marker3Ds。",
                "document": document_name,
                "_product_document": product_document,
                "_root_product": getattr(product_document, "Product", None),
            }
        try:
            product_document.Save()
        except Exception as exc:
            return {
                "status": "failed",
                "message": f"CATProduct Save 失败: {exc}",
                "document": document_name,
                "path": save_path_text,
                "_product_document": product_document,
                "_root_product": getattr(product_document, "Product", None),
            }

    reopened_document = product_document
    close_warning = None
    if close_before_open:
        try:
            product_document.Close()
            time.sleep(max(0.0, float(wait_seconds)))
            reopened_document = catia.Documents.Open(save_path_text)
        except Exception as exc:
            close_warning = str(exc)
            try:
                reopened_document = catia.Documents.Open(save_path_text)
            except Exception:
                reopened_document = product_document
    try:
        reopened_root_product = reopened_document.Product
    except Exception:
        reopened_root_product = None
    refresh_product_view(reopened_document, reopened_root_product)
    return {
        "status": "success" if close_warning is None else "warning",
        "message": (
            "CATProduct 已保存、关闭并重新打开以刷新 Marker3Ds。"
            if close_warning is None
            else f"CATProduct 保存后重开存在警告: {close_warning}"
        ),
        "document": _safe_name(reopened_document),
        "path": save_path_text,
        "closed_and_reopened": bool(close_before_open and close_warning is None),
        "close_warning": close_warning,
        "_product_document": reopened_document,
        "_root_product": reopened_root_product,
    }


def create_annotation_text_marker(
    product_document: Any,
    root_product: Any,
    process_part: Any,
    support_feature: Any,
    text_position: Vector,
    anchor_position: Vector,
    label: str,
    name: str,
    *,
    text_color: tuple[int, int, int] = DEFAULT_TEXT_COLOR,
    text_size: float = DEFAULT_TEXT_SIZE,
) -> dict[str, Any]:
    """在 CATProduct 中创建 3D 文本标记。"""
    marker3ds = get_marker3ds(root_product)
    if marker3ds is None:
        return {
            "text_created": False,
            "message": "当前 CATProduct 不支持 Marker3Ds，无法创建 3D 文本。",
        }
    try:
        product_document.Activate()
    except Exception:
        pass
    try:
        process_part.Update()
    except Exception:
        pass
    try:
        root_product.Update()
    except Exception:
        pass
    activate_feature_for_display_refresh(product_document, support_feature)
    marker = marker3ds.Add3DText(
        tuple(float(value) for value in text_position),
        str(label),
        tuple(float(value) for value in anchor_position),
        support_feature,
    )
    try:
        marker.Name = name
    except Exception:
        pass
    text_size_result = _set_marker_text_size(marker, text_size)
    try:
        set_object_visual_style(product_document, marker, text_color, width=DEFAULT_LINE_WIDTH)
    except Exception:
        pass
    _refresh_marker_display(product_document, root_product, marker)
    activate_feature_for_display_refresh(product_document, marker)
    return {
        "text_created": True,
        "text_name": _safe_name(marker),
        "text_position": _round_vector(text_position),
        "anchor_position": _round_vector(anchor_position),
        "text": str(label),
        **text_size_result,
        "_text_marker": marker,
    }


def save_external_process_part_if_needed(process_info: dict[str, Any], stage: str) -> dict[str, Any] | None:
    if not process_info.get("external_file") or process_info.get("part_document") is None:
        return None
    try:
        process_info["part_document"].Save()
        return {
            "status": "success",
            "stage": stage,
            "path": process_info.get("save_path"),
            "message": f"标注过程 CATPart 已保存: {stage}",
        }
    except Exception as exc:
        return {
            "status": "failed",
            "stage": stage,
            "path": process_info.get("save_path"),
            "message": str(exc),
        }


def default_distance_label(point1: Vector, point2: Vector, digits: int = 3) -> str:
    """按两点距离生成默认标注文本。"""
    return f"{_length(_subtract(point2, point1)):.{digits}f} mm"


def create_distance_annotation(
    catia_or_caa: Any | None = None,
    *,
    point1: Iterable[float],
    point2: Iterable[float],
    annotation_text: str | None = None,
    annotation_name: str = "DistanceAnnotation",
    hybrid_body_name: str = DEFAULT_HYBRID_BODY_NAME,
    text_offset_direction: Iterable[float] = (0.0, 0.0, 1.0),
    text_offset_distance: float = 5.0,
    feature_color: tuple[int, int, int] = DEFAULT_FEATURE_COLOR,
    text_color: tuple[int, int, int] = DEFAULT_TEXT_COLOR,
    line_width: int = DEFAULT_LINE_WIDTH,
    text_size: float = DEFAULT_TEXT_SIZE,
    process_part: Any | None = None,
    process_part_document: Any | None = None,
    process_component: Any | None = None,
    create_process_part_if_missing: bool = True,
    process_part_number: str = DEFAULT_PROCESS_PART_NUMBER,
    process_part_name: str = DEFAULT_PROCESS_PART_NAME,
    process_part_save_path: str | Path | None = None,
    reopen_product_after_create: bool = False,
    product_save_path: str | Path | None = None,
) -> dict[str, Any]:
    """
    创建一条两点距离标注。

    参数:
        catia_or_caa:
            CATIA.Application 或 pycatia/caa 包装对象；不传时尝试连接当前 CATIA。
        point1 / point2:
            标注线两个端点的装配坐标，三维坐标 (x, y, z)。
        annotation_text:
            标注文本；不传时默认为两点距离，格式为 "xxx.xxx mm"。
        annotation_name:
            标注名称，会作为点、线和文本名称前缀。
        hybrid_body_name:
            过程 Part 中承载标注点线的几何图形集名称。
        text_offset_direction:
            文本相对两点中点的偏移参考方向。
        text_offset_distance:
            文本沿偏移参考方向移动的距离，单位按 CATIA 模型单位，通常为 mm。
        feature_color:
            点和线颜色，RGB 元组，例如 (0, 255, 0)。
        text_color:
            3D 文本颜色，RGB 元组。
        line_width:
            标注线宽。
        text_size:
            3D 文本字号。
        process_part/process_part_document/process_component:
            可选的过程 Part。传入后直接在该 Part 里创建点线。
        create_process_part_if_missing:
            未传过程 Part 时是否自动在当前 CATProduct 下新建一个 Part。
        process_part_number/process_part_name:
            自动新建过程 Part 时使用的 PartNumber 和名称。
        process_part_save_path:
            可选的过程 CATPart 保存路径。传入后工具会先单独创建并 SaveAs 该 CATPart，
            再将它装配到当前 CATProduct 中承载标注几何；每次创建标注后会保存该 CATPart。
            不传时保持兼容旧逻辑，直接在当前 CATProduct 下 AddNewComponent 新建过程 Part。
        reopen_product_after_create/product_save_path:
            可选的 Product 级 Marker3Ds 刷新流程。单次标注时可设为 True，工具会在
            创建完成后保存 CATProduct、关闭并重新打开，再刷新视图。未保存的 Product
            需要提供 product_save_path。批量标注建议通过 create_distance_annotations 统一执行一次。

    运行环境:
        当前活动文档必须是 CATProduct，因为 3D 文本通过 root_product.Marker3Ds 创建。
    """
    product_document, root_product = get_active_product_document(catia_or_caa)
    p1 = _as_vector(point1, "point1")
    p2 = _as_vector(point2, "point2")
    label = str(annotation_text) if annotation_text is not None else default_distance_label(p1, p2)
    offset_direction = _normalize(text_offset_direction, "text_offset_direction")
    midpoint = _average(p1, p2)
    text_position = _add(midpoint, _scale(offset_direction, float(text_offset_distance)))

    process_info = resolve_process_part(
        product_document,
        root_product,
        process_part=process_part,
        process_part_document=process_part_document,
        process_component=process_component,
        create_if_missing=create_process_part_if_missing,
        process_part_number=process_part_number,
        process_part_name=process_part_name,
        process_part_save_path=process_part_save_path,
    )
    part = process_info["part"]
    part_document = process_info.get("part_document") or product_document
    hybrid_body = get_or_create_hybrid_body(part, hybrid_body_name)
    geometry = create_annotation_line_geometry(
        part_document,
        part,
        hybrid_body,
        p1,
        p2,
        annotation_name=annotation_name,
        feature_color=feature_color,
        line_width=line_width,
    )
    geometry_save_result = save_external_process_part_if_needed(process_info, "geometry_before_text")
    if geometry_save_result is not None and geometry_save_result.get("status") == "failed":
        return {
            "status": "failed",
            "message": f"标注线已创建，但创建文字前保存过程 CATPart 失败: {geometry_save_result.get('message')}",
            "product_document": _safe_name(product_document),
            "root_product": _safe_name(root_product),
            "process_part": _safe_name(part),
            "process_part_created": bool(process_info.get("created")),
            "process_part_external_file": bool(process_info.get("external_file")),
            "process_part_save_path": process_info.get("save_path"),
            "process_part_geometry_save_result": geometry_save_result,
            "process_part_save_result": None,
            "hybrid_body_name": hybrid_body_name,
            "annotation_name": annotation_name,
            "point1": _round_vector(p1),
            "point2": _round_vector(p2),
            "distance": round(_length(_subtract(p2, p1)), 6),
            "annotation_text": label,
            "text_offset_direction": _round_vector(offset_direction),
            "text_offset_distance": float(text_offset_distance),
            "text_position": _round_vector(text_position),
            "anchor_position": _round_vector(midpoint),
            "geometry": {
                key: value
                for key, value in geometry.items()
                if not key.endswith("_feature")
            },
            "text": {"text_created": False, "message": "创建文字前过程 CATPart 保存失败。"},
        }
    text = create_annotation_text_marker(
        product_document,
        root_product,
        part,
        geometry["line_feature"],
        text_position,
        midpoint,
        label,
        f"{annotation_name}_Text",
        text_color=text_color,
        text_size=text_size,
    )
    try:
        part.Update()
    except Exception:
        pass
    part_save_result = save_external_process_part_if_needed(process_info, "after_text")
    try:
        root_product.Update()
    except Exception:
        pass
    status = "success"
    message = None
    if part_save_result is not None and part_save_result.get("status") == "failed":
        status = "failed"
        message = f"标注几何已创建，但过程 CATPart 保存失败: {part_save_result.get('message')}"
    product_reopen_result = None
    if reopen_product_after_create and status == "success":
        product_reopen_result = reopen_product_document_for_marker_refresh(
            catia_or_caa,
            product_document,
            product_save_path=product_save_path,
        )
    return {
        "status": status,
        "message": message,
        "product_document": _safe_name(product_document),
        "root_product": _safe_name(root_product),
        "process_part": _safe_name(part),
        "process_part_created": bool(process_info.get("created")),
        "process_part_external_file": bool(process_info.get("external_file")),
        "process_part_save_path": process_info.get("save_path"),
        "process_part_geometry_save_result": geometry_save_result,
        "process_part_save_result": part_save_result,
        "product_reopen_result": (
            {
                key: value
                for key, value in product_reopen_result.items()
                if not str(key).startswith("_")
            }
            if product_reopen_result is not None
            else None
        ),
        "hybrid_body_name": hybrid_body_name,
        "annotation_name": annotation_name,
        "point1": _round_vector(p1),
        "point2": _round_vector(p2),
        "distance": round(_length(_subtract(p2, p1)), 6),
        "annotation_text": label,
        "text_offset_direction": _round_vector(offset_direction),
        "text_offset_distance": float(text_offset_distance),
        "text_position": _round_vector(text_position),
        "anchor_position": _round_vector(midpoint),
        "geometry": {
            key: value
            for key, value in geometry.items()
            if not key.endswith("_feature")
        },
        "text": text,
        "_annotation_geometry": geometry,
        "_text_marker": text.get("_text_marker") if isinstance(text, dict) else None,
        "_process_part": part,
        "_process_part_document": process_info.get("part_document"),
        "_process_component": process_info.get("component"),
        "_product_document": (
            product_reopen_result.get("_product_document")
            if product_reopen_result is not None
            else product_document
        ),
        "_root_product": (
            product_reopen_result.get("_root_product")
            if product_reopen_result is not None
            else root_product
        ),
    }


def create_distance_annotations(
    catia_or_caa: Any | None = None,
    annotations: list[dict[str, Any]] | None = None,
    **shared_kwargs: Any,
) -> dict[str, Any]:
    """
    批量创建多条距离标注。

    annotations 中每个 dict 支持 create_distance_annotation 的同名参数。
    shared_kwargs 会作为默认参数传给每一条标注。
    如果在 shared_kwargs 中传入 process_part_save_path，批量标注会共用同一个
    外部过程 CATPart，第一条创建时装配，后续标注继续写入并保存该 Part。
    如果传入 reopen_product_after_create=True，批量标注会在全部条目创建完后
    只保存、关闭并重新打开一次 CATProduct，而不是每条标注都重开。
    """
    if annotations is None:
        annotations = []
    shared_kwargs = dict(shared_kwargs)
    reopen_product_after_create = bool(shared_kwargs.pop("reopen_product_after_create", False))
    product_save_path = shared_kwargs.pop("product_save_path", None)
    results: list[dict[str, Any]] = []
    process_part = shared_kwargs.get("process_part")
    process_part_document = shared_kwargs.get("process_part_document")
    process_component = shared_kwargs.get("process_component")
    product_document = None
    root_product = None
    for index, item in enumerate(annotations, start=1):
        kwargs = {**shared_kwargs, **item}
        if process_part is not None:
            kwargs["process_part"] = process_part
        if process_part_document is not None:
            kwargs["process_part_document"] = process_part_document
        if process_component is not None:
            kwargs["process_component"] = process_component
        kwargs.setdefault("annotation_name", f"DistanceAnnotation_{index:03d}")
        try:
            result = create_distance_annotation(catia_or_caa, **kwargs)
            results.append(result)
            process_part = result.get("_process_part") or process_part
            process_part_document = result.get("_process_part_document") or process_part_document
            process_component = result.get("_process_component") or process_component
            product_document = result.get("_product_document") or product_document
            root_product = result.get("_root_product") or root_product
        except Exception as exc:
            results.append(
                {
                    "status": "failed",
                    "annotation_name": kwargs.get("annotation_name"),
                    "message": str(exc),
                }
            )
    product_reopen_result = None
    if reopen_product_after_create and any(result.get("status") == "success" for result in results):
        try:
            if product_document is None:
                product_document, root_product = get_active_product_document(catia_or_caa)
            product_reopen_result = reopen_product_document_for_marker_refresh(
                catia_or_caa,
                product_document,
                product_save_path=product_save_path,
            )
            product_document = product_reopen_result.get("_product_document") or product_document
            root_product = product_reopen_result.get("_root_product") or root_product
        except Exception as exc:
            product_reopen_result = {
                "status": "failed",
                "message": str(exc),
            }
    return {
        "status": "success",
        "count": len(results),
        "success_count": sum(1 for result in results if result.get("status") == "success"),
        "results": results,
        "product_reopen_result": (
            {
                key: value
                for key, value in product_reopen_result.items()
                if not str(key).startswith("_")
            }
            if product_reopen_result is not None
            else None
        ),
        "_process_part": process_part,
        "_process_part_document": process_part_document,
        "_process_component": process_component,
        "_product_document": product_document,
        "_root_product": root_product,
    }
