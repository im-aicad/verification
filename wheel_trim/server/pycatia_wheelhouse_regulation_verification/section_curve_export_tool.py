"""Standalone CATIA section curve export tool.

This file is intentionally self-contained so the whole folder can be copied to a
CATIA workstation and executed without the main catia_agent_v3 package.
"""

from __future__ import annotations

import argparse
import json
import sys
import time
from datetime import datetime
from pathlib import Path
from typing import Any


TOOL_ROOT = Path(__file__).resolve().parent
ORIGIN_PLANE_NAMES = {"xy", "yz", "zx"}
NAMED_COLORS: dict[str, tuple[int, int, int]] = {
    "gray": (150, 150, 150),
    "grey": (150, 150, 150),
    "yellow": (255, 220, 40),
    "green": (70, 190, 90),
    "orange": (255, 145, 40),
    "red": (235, 55, 55),
}


def ok(message: str, **extra: Any) -> dict[str, Any]:
    payload = {"status": "success", "message": message}
    payload.update(extra)
    return payload


def err(message: str, *, error_code: str = "ERROR", **extra: Any) -> dict[str, Any]:
    payload = {"status": "error", "message": message, "error_code": error_code}
    payload.update(extra)
    return payload


def print_json(payload: dict[str, Any]) -> None:
    print(json.dumps(json_safe(payload), ensure_ascii=False, indent=2))


def json_safe(value: Any) -> Any:
    if isinstance(value, dict):
        return {str(json_safe(key)): json_safe(item) for key, item in value.items()}
    if isinstance(value, list):
        return [json_safe(item) for item in value]
    if isinstance(value, tuple):
        return [json_safe(item) for item in value]
    if isinstance(value, str):
        return value.encode("utf-8", errors="replace").decode("utf-8", errors="replace")
    if isinstance(value, (int, float, bool)) or value is None:
        return value
    return str(value)


def write_json(path: Path, payload: dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(json_safe(payload), ensure_ascii=True, indent=2), encoding="utf-8")


def write_markdown(path: Path, payload: dict[str, Any]) -> None:
    lines = [
        "# Section Curve Export Report",
        "",
        f"- Status: `{payload.get('status')}`",
        f"- Document: `{payload.get('document_name')}`",
        f"- Part: `{payload.get('part_name')}`",
        f"- Target: `{payload.get('target_name')}`",
        f"- Section plane: `{payload.get('section_plane_name')}`",
        f"- Section plane definition: `{payload.get('section_plane_definition')}`",
        f"- Section curve: `{payload.get('section_curve_name')}`",
        f"- Exported CATPart: `{payload.get('exported_catpart')}`",
        "",
        "## Attempts",
        "",
    ]
    for attempt in payload.get("attempts", []):
        lines.append(
            "- target={target}, plane={plane}, existing={existing}, status={status}, length={length}".format(
                target=attempt.get("target_name"),
                plane=attempt.get("plane_name"),
                existing=attempt.get("use_existing_plane"),
                status=attempt.get("status"),
                length=attempt.get("curve_length"),
            )
        )
    lines.extend(["", "## Health", ""])
    for health in payload.get("health", []):
        lines.append(f"- `{health.get('feature_name')}`: `{health.get('health_status')}`")
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text("\n".join(lines) + "\n", encoding="utf-8")


def _get_catia_from_rot(win32com_client: Any) -> Any | None:
    try:
        import pythoncom  # type: ignore
        import pywintypes  # type: ignore
    except Exception:
        return None

    pythoncom.CoInitialize()
    for attempt in range(3):
        try:
            rot = pythoncom.GetRunningObjectTable()
            monikers = rot.EnumRunning()
        except Exception:
            return None

        entries: list[tuple[Any, str]] = []
        while True:
            batch = monikers.Next(1)
            if not batch:
                break
            moniker = batch[0]
            try:
                display_name = str(moniker.GetDisplayName(pythoncom.CreateBindCtx(0), None))
            except Exception:
                display_name = ""
            entries.append((moniker, display_name))
        entries.sort(key=lambda item: 0 if item[1].startswith("!{") else 1)

        for moniker, display_name in entries:
            try:
                if display_name.startswith("!{") and display_name.endswith("}"):
                    unknown = pythoncom.GetActiveObject(pywintypes.IID(display_name[1:]))
                else:
                    unknown = rot.GetObject(moniker)
                dispatch = unknown.QueryInterface(pythoncom.IID_IDispatch)
                candidate = win32com_client.Dispatch(dispatch)
            except Exception:
                continue

            for active_candidate in [candidate, getattr(candidate, "Application", None)]:
                if active_candidate is None:
                    continue
                try:
                    document = active_candidate.ActiveDocument
                    document_name = str(getattr(document, "Name", ""))
                    if document_name.lower().endswith(".catpart"):
                        return active_candidate
                except Exception:
                    continue
        if attempt < 2:
            time.sleep(0.1)
    return None


def get_active_catia_part() -> tuple[Any, Any, Any]:
    try:
        import win32com.client  # type: ignore
    except Exception as exc:
        raise RuntimeError("pywin32 is required. Install with: pip install pywin32") from exc

    catia = _get_catia_from_rot(win32com.client)
    active_object_error: Exception | None = None
    if catia is None:
        try:
            catia = win32com.client.GetActiveObject("CATIA.Application")
        except Exception as exc:
            active_object_error = exc
    if catia is None:
        raise RuntimeError("CATIA is not running or no active CATIA session is available.") from active_object_error

    try:
        document = catia.ActiveDocument
    except Exception as exc:
        raise RuntimeError("CATIA does not have an active document.") from exc

    document_name = str(getattr(document, "Name", ""))
    if not document_name.lower().endswith(".catpart"):
        raise RuntimeError(f"Active CATIA document is not a CATPart: {document_name}")

    try:
        part = document.Part
    except Exception as exc:
        raise RuntimeError("Active CATIA CATPart does not expose Part.") from exc
    return catia, document, part


def count(container: Any) -> int:
    for attr_name in ("Count", "Count2", "count"):
        try:
            return int(getattr(container, attr_name))
        except Exception:
            pass
    return 0


def item(container: Any, index_or_name: Any) -> Any:
    for method_name in ("Item", "Item2", "item"):
        method = getattr(container, method_name, None)
        if callable(method):
            return method(index_or_name)
    raise RuntimeError(f"CATIA collection item {index_or_name!r} is unavailable.")


def safe_name(obj: Any, attr_name: str = "Name") -> str | None:
    try:
        value = getattr(obj, attr_name)
    except Exception:
        return None
    try:
        return str(value)
    except Exception:
        return None


def iter_hybrid_bodies(container: Any):
    if container is None:
        return
    for index in range(1, count(container) + 1):
        try:
            body = item(container, index)
        except Exception:
            continue
        yield body
        nested = getattr(body, "HybridBodies", None) or getattr(body, "hybrid_bodies", None)
        if nested is not None:
            yield from iter_hybrid_bodies(nested)


def find_hybrid_body_by_name(part: Any, body_name: str) -> Any | None:
    for body in iter_hybrid_bodies(getattr(part, "HybridBodies", None)):
        if safe_name(body) == body_name:
            return body
    return None


def get_or_create_hybrid_body(part: Any, body_name: str) -> Any:
    existing = find_hybrid_body_by_name(part, body_name)
    if existing is not None:
        return existing
    bodies = getattr(part, "HybridBodies", None)
    if bodies is None:
        raise RuntimeError("Active CATIA part does not expose HybridBodies.")
    body = bodies.Add()
    body.Name = body_name
    return body


def find_hybrid_shape_by_name(part: Any, feature_name: str, prefer_latest: bool = True) -> tuple[Any, Any, dict[str, Any]]:
    matches: list[tuple[Any, Any, int]] = []
    for body in iter_hybrid_bodies(getattr(part, "HybridBodies", None)):
        shapes = getattr(body, "HybridShapes", None)
        if shapes is None:
            continue
        for index in range(1, count(shapes) + 1):
            try:
                shape = item(shapes, index)
            except Exception:
                continue
            if safe_name(shape) == feature_name:
                matches.append((shape, body, index))
    if not matches:
        raise ValueError(f"Hybrid shape '{feature_name}' was not found.")
    selected = matches[-1] if prefer_latest else matches[0]
    shape, body, index = selected
    return shape, body, {
        "matched_count": len(matches),
        "selected_index": index - 1,
        "body_name": safe_name(body),
        "ambiguous": len(matches) > 1,
    }


def find_hybrid_sketch_by_name(part: Any, sketch_name: str) -> tuple[Any, Any, dict[str, Any]]:
    for body in iter_hybrid_bodies(getattr(part, "HybridBodies", None)):
        sketches = getattr(body, "HybridSketches", None)
        if sketches is None:
            continue
        for index in range(1, count(sketches) + 1):
            try:
                sketch = item(sketches, index)
            except Exception:
                continue
            if safe_name(sketch) == sketch_name:
                return sketch, body, {"body_name": safe_name(body), "selected_index": index - 1}
    raise ValueError(f"Hybrid sketch '{sketch_name}' was not found.")


def find_body_by_name(part: Any, body_name: str) -> Any | None:
    bodies = getattr(part, "Bodies", None)
    if bodies is None:
        return None
    try:
        return item(bodies, body_name)
    except Exception:
        pass
    for index in range(1, count(bodies) + 1):
        try:
            body = item(bodies, index)
        except Exception:
            continue
        if safe_name(body) == body_name:
            return body
    return None


def create_reference_from_name(part: Any, name: str) -> tuple[Any, dict[str, Any]]:
    normalized = name.lower()
    origin = getattr(part, "OriginElements", None)
    origin_map = {
        "xy": getattr(origin, "PlaneXY", None),
        "yz": getattr(origin, "PlaneYZ", None),
        "zx": getattr(origin, "PlaneZX", None),
    }
    target = origin_map.get(normalized)
    metadata: dict[str, Any] = {"source_type": "origin_plane" if target is not None else "hybrid_shape", "source_name": name}
    if target is None:
        try:
            target, _parent, shape_metadata = find_hybrid_shape_by_name(part, name)
            metadata.update(shape_metadata)
        except ValueError as shape_error:
            try:
                target, _parent, sketch_metadata = find_hybrid_sketch_by_name(part, name)
            except ValueError:
                raise shape_error
            metadata["source_type"] = "hybrid_sketch"
            metadata.update(sketch_metadata)
    try:
        return part.CreateReferenceFromObject(target), metadata
    except Exception as exc:
        raise RuntimeError(f"Failed to create CATIA reference for '{name}'.") from exc


def reference_from_any_name(part: Any, name: str) -> tuple[Any, dict[str, Any]]:
    try:
        return create_reference_from_name(part, name)
    except Exception as hybrid_exc:
        body = find_body_by_name(part, name)
        source_type = "body"
        if body is None:
            body = find_hybrid_body_by_name(part, name)
            source_type = "hybrid_body"
        if body is None:
            raise hybrid_exc
        return part.CreateReferenceFromObject(body), {"source_type": source_type, "source_name": name}


def append_update_feature(part: Any, target_body: Any, feature: Any, result_name: str) -> dict[str, Any]:
    try:
        feature.Name = result_name
        target_body.AppendHybridShape(feature)
        part.InWorkObject = feature
        update_object = getattr(part, "UpdateObject", None)
        if callable(update_object):
            update_object(feature)
        else:
            part.Update()
    except Exception as exc:
        raise RuntimeError(f"CATIA update failed for live feature '{result_name}'.") from exc
    return {"result_name": result_name, "target_body": safe_name(target_body), "feature_name": safe_name(feature) or result_name}


def measure_feature(document: Any, part: Any, feature: Any) -> dict[str, Any]:
    result: dict[str, Any] = {"name": safe_name(feature), "type_name": safe_name(feature, "Type") or type(feature).__name__}
    try:
        ref = part.CreateReferenceFromObject(feature)
        measurable = document.GetWorkbench("SPAWorkbench").GetMeasurable(ref)
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
        result["cog"] = [float(value) for value in cog]
    except Exception:
        try:
            coords = [0.0, 0.0, 0.0]
            measurable.GetCOG(coords)
            result["cog"] = [float(value) for value in coords]
        except Exception:
            pass
    return result


def parse_float_list(raw: str | None, *, expected_count: int, option_name: str) -> list[float] | None:
    if raw is None:
        return None
    normalized = raw.replace(";", ",").replace(" ", ",")
    values = [item for item in normalized.split(",") if item != ""]
    if len(values) != expected_count:
        raise ValueError(f"{option_name} requires {expected_count} numeric values.")
    try:
        return [float(value) for value in values]
    except ValueError as exc:
        raise ValueError(f"{option_name} contains a non-numeric value: {raw}") from exc


def vector_length(vector: list[float]) -> float:
    return sum(value * value for value in vector) ** 0.5


def normalized_vector(vector: list[float], *, option_name: str) -> list[float]:
    length = vector_length(vector)
    if length <= 1e-12:
        raise ValueError(f"{option_name} cannot be a zero vector.")
    return [value / length for value in vector]


def base_plane_normal(section_plane: str) -> list[float]:
    plane = section_plane.lower()
    if plane == "yz":
        return [1.0, 0.0, 0.0]
    if plane == "zx":
        return [0.0, 1.0, 0.0]
    if plane == "xy":
        return [0.0, 0.0, 1.0]
    raise ValueError("--through-point without --normal requires --section-plane xy, yz, or zx.")


def has_explicit_plane_definition(args: argparse.Namespace) -> bool:
    return any(
        [
            args.offset_distance is not None,
            args.through_point is not None,
            args.normal is not None,
            args.plane_equation is not None,
            args.axis_name is not None,
            args.angle_deg is not None,
            args.angle_reverse,
        ]
    )


def section_plane_definition_from_args(args: argparse.Namespace) -> dict[str, Any]:
    if (args.axis_name is None) != (args.angle_deg is None):
        raise ValueError("--axis-name and --angle-deg must be provided together.")
    equation = parse_float_list(args.plane_equation, expected_count=4, option_name="--plane-equation")
    through_point = parse_float_list(args.through_point, expected_count=3, option_name="--through-point")
    normal = parse_float_list(args.normal, expected_count=3, option_name="--normal")

    if equation is not None and (through_point is not None or normal is not None):
        raise ValueError("Use either --plane-equation or --through-point/--normal, not both.")
    if equation is not None:
        normal_part = normalized_vector(equation[:3], option_name="--plane-equation normal")
        scale = vector_length(equation[:3])
        return {
            "mode": "equation",
            "equation": [normal_part[0], normal_part[1], normal_part[2], float(equation[3]) / scale],
            "raw_equation": equation,
            "equation_format": "A*x + B*y + C*z = D",
        }

    if through_point is not None or normal is not None:
        if through_point is None:
            raise ValueError("--normal requires --through-point.")
        normal_vector = normalized_vector(normal if normal is not None else base_plane_normal(args.section_plane), option_name="--normal")
        d_value = sum(normal_vector[index] * through_point[index] for index in range(3))
        return {
            "mode": "point_normal",
            "through_point": through_point,
            "normal": normal_vector,
            "equation": [normal_vector[0], normal_vector[1], normal_vector[2], d_value],
            "equation_format": "A*x + B*y + C*z = D",
        }

    if has_explicit_plane_definition(args):
        if args.section_plane == "auto":
            raise ValueError("Explicit offset/angle positioning requires --section-plane xy, yz, zx, or an existing plane name.")
        return {
            "mode": "offset_angle",
            "base_plane_name": args.section_plane,
            "offset_distance": float(args.offset_distance or 0.0),
            "reverse": bool(args.reverse),
            "axis_name": args.axis_name,
            "angle_deg": args.angle_deg,
            "angle_reverse": bool(args.angle_reverse),
        }

    return {"mode": "auto", "section_plane": args.section_plane}


def create_section_plane_by_equation(
    result_name: str,
    *,
    equation: list[float],
    target_body_name: str = "section_results",
    source_definition: dict[str, Any] | None = None,
) -> dict[str, Any]:
    _catia, document, part = get_active_catia_part()
    target_body = get_or_create_hybrid_body(part, target_body_name)
    factory = part.HybridShapeFactory
    create_method = getattr(factory, "AddNewPlaneEquation", None)
    if not callable(create_method):
        raise RuntimeError("CATIA HybridShapeFactory does not expose AddNewPlaneEquation.")
    a_value, b_value, c_value, d_value = [float(value) for value in equation]
    plane = create_method(a_value, b_value, c_value, d_value)
    metadata = append_update_feature(part, target_body, plane, result_name)
    return ok(
        f"Section plane '{result_name}' created from equation.",
        tool="create_section_plane_by_equation",
        result_name=result_name,
        document_name=safe_name(document),
        created_features=[result_name],
        raw_result={
            **metadata,
            "equation": [a_value, b_value, c_value, d_value],
            "equation_format": "A*x + B*y + C*z = D",
            "source_definition": source_definition or {},
            "attempted_methods": ["AddNewPlaneEquation"],
        },
    )


def create_section_plane_by_axis(
    result_name: str,
    *,
    base_plane_name: str,
    offset_distance: float = 0.0,
    reverse: bool = False,
    axis_name: str | None = None,
    angle_deg: float | None = None,
    angle_reverse: bool = False,
    target_body_name: str = "section_results",
) -> dict[str, Any]:
    _catia, document, part = get_active_catia_part()
    target_body = get_or_create_hybrid_body(part, target_body_name)
    factory = part.HybridShapeFactory
    base_ref, base_metadata = create_reference_from_name(part, base_plane_name)
    working_ref = base_ref
    construction: list[str] = []
    attempted: list[str] = []

    if float(offset_distance) != 0.0:
        offset_feature = factory.AddNewPlaneOffset(base_ref, abs(float(offset_distance)), bool(reverse) or float(offset_distance) < 0.0)
        seed_name = f"_tmp_{result_name}_offset_seed"
        append_update_feature(part, target_body, offset_feature, seed_name)
        working_ref = part.CreateReferenceFromObject(offset_feature)
        construction.append(seed_name)
        attempted.append("AddNewPlaneOffset")

    if axis_name and angle_deg is not None:
        axis_ref, axis_metadata = reference_from_any_name(part, axis_name)
        try:
            plane = factory.AddNewPlaneAngle(working_ref, axis_ref, float(angle_deg), bool(angle_reverse))
            attempted.append("AddNewPlaneAngle(ref, axis, angle, reverse)")
        except TypeError:
            plane = factory.AddNewPlaneAngle(working_ref, axis_ref, float(angle_deg))
            attempted.append("AddNewPlaneAngle(ref, axis, angle)")
    else:
        axis_metadata = None
        datum_method = getattr(factory, "AddNewPlaneDatum", None)
        if callable(datum_method):
            try:
                plane = datum_method(working_ref)
                attempted.append("AddNewPlaneDatum")
            except Exception:
                plane = factory.AddNewPlaneOffset(working_ref, 0.001, False)
                attempted.append("AddNewPlaneOffset(0.001 fallback)")
        else:
            plane = factory.AddNewPlaneOffset(working_ref, 0.001, False)
            attempted.append("AddNewPlaneOffset(0.001 fallback)")

    metadata = append_update_feature(part, target_body, plane, result_name)
    return ok(
        f"Section plane '{result_name}' created.",
        tool="create_section_plane_by_axis",
        result_name=result_name,
        document_name=safe_name(document),
        created_features=construction + [result_name],
        raw_result={
            **metadata,
            "base_plane_name": base_plane_name,
            "base_metadata": base_metadata,
            "axis_name": axis_name,
            "axis_metadata": axis_metadata,
            "offset_distance": offset_distance,
            "reverse": reverse,
            "angle_deg": angle_deg,
            "angle_reverse": angle_reverse,
            "attempted_methods": attempted,
        },
    )


def create_section_intersection_curves(
    *,
    section_plane_name: str,
    target_element_names: list[str],
    result_name: str,
    target_body_name: str = "section_results",
    extend_mode: bool = False,
) -> dict[str, Any]:
    _catia, document, part = get_active_catia_part()
    target_body = get_or_create_hybrid_body(part, target_body_name)
    factory = part.HybridShapeFactory
    section_ref, section_metadata = reference_from_any_name(part, section_plane_name)
    created: list[str] = []
    intersections: list[dict[str, Any]] = []
    for index, target_name in enumerate(target_element_names, start=1):
        target_ref, target_metadata = reference_from_any_name(part, target_name)
        feature = factory.AddNewIntersection(section_ref, target_ref)
        if hasattr(feature, "ExtendMode"):
            feature.ExtendMode = bool(extend_mode)
        curve_name = result_name if len(target_element_names) == 1 else f"{result_name}_{index:03d}"
        metadata = append_update_feature(part, target_body, feature, curve_name)
        measurement = measure_feature(document, part, feature)
        created.append(curve_name)
        intersections.append({"curve_name": curve_name, "target_name": target_name, "metadata": metadata, "target_metadata": target_metadata, "measurement": measurement})
    return ok(
        f"Created {len(created)} section intersection curve(s).",
        tool="create_section_intersection_curves",
        result_name=result_name,
        document_name=safe_name(document),
        created_features=created,
        raw_result={"section_plane_name": section_plane_name, "section_metadata": section_metadata, "intersections": intersections},
    )


def export_section_curves_as_catpart(
    *,
    curve_names: list[str],
    result_name: str,
    output_path: Path,
    target_body_name: str = "SectionResult",
    close_exported_document: bool = False,
    reactivate_source_document: bool = True,
) -> dict[str, Any]:
    catia, source_document, source_part = get_active_catia_part()
    source_selection = source_document.Selection
    selected: list[dict[str, Any]] = []
    source_selection.Clear()
    try:
        for name in curve_names:
            shape, body, metadata = find_hybrid_shape_by_name(source_part, name)
            source_selection.Add(shape)
            selected.append({"curve_name": name, "body_name": safe_name(body), "metadata": metadata})
        source_selection.Copy()
    finally:
        try:
            source_selection.Clear()
        except Exception:
            pass

    exported_document = catia.Documents.Add("Part")
    exported_part = exported_document.Part
    target_body = get_or_create_hybrid_body(exported_part, target_body_name)
    target_shapes = getattr(target_body, "HybridShapes", None)
    before_shape_count = count(target_shapes) if target_shapes is not None else 0
    paste_selection = exported_document.Selection
    paste_errors: list[dict[str, str]] = []
    used_paste_method = None
    for fmt in ["CATPrtResultWithOutLink", "CATPrtResult", "CATPrtCont"]:
        try:
            paste_selection.Clear()
            paste_selection.Add(target_body)
            paste_selection.PasteSpecial(fmt)
            used_paste_method = f"PasteSpecial({fmt})"
            break
        except Exception as exc:
            paste_errors.append({"method": f"PasteSpecial({fmt})", "message": str(exc)})
    if used_paste_method is None:
        paste_selection.Clear()
        paste_selection.Add(target_body)
        paste_selection.Paste()
        used_paste_method = "Paste()"
    paste_selection.Clear()

    try:
        exported_part.Update()
    except Exception:
        pass
    exported_features: list[dict[str, Any]] = []
    try:
        target_shapes = getattr(target_body, "HybridShapes", None)
        after_shape_count = count(target_shapes) if target_shapes is not None else 0
        pasted_count = max(0, after_shape_count - before_shape_count)
        for index in range(1, pasted_count + 1):
            shape = item(target_shapes, before_shape_count + index)
            requested_name = curve_names[index - 1] if index <= len(curve_names) else f"{result_name}_{index:03d}"
            try:
                shape.Name = requested_name
            except Exception:
                pass
            exported_features.append(
                {
                    "body_name": target_body_name,
                    "feature_name": safe_name(shape) or requested_name,
                    "source_curve_name": curve_names[index - 1] if index <= len(curve_names) else None,
                    "index": before_shape_count + index,
                }
            )
        try:
            exported_part.Update()
        except Exception:
            pass
    except Exception as exc:
        exported_features.append({"status": "rename_failed", "message": str(exc)})
    output_path.parent.mkdir(parents=True, exist_ok=True)
    exported_document.SaveAs(str(output_path))
    if close_exported_document:
        try:
            exported_document.Close()
        except Exception:
            pass
    if reactivate_source_document:
        try:
            source_document.Activate()
        except Exception:
            pass
    return ok(
        f"Exported {len(curve_names)} curve(s) to '{output_path}'.",
        tool="export_section_curves_as_catpart",
        result_name=result_name,
        document_name=safe_name(source_document),
        exported_path=str(output_path),
        raw_result={
            "source_document": safe_name(source_document),
            "exported_document": safe_name(exported_document),
            "curve_names": curve_names,
            "selected": selected,
            "exported_features": exported_features,
            "target_body_name": target_body_name,
            "used_paste_method": used_paste_method,
            "paste_errors": paste_errors,
            "file_exists": output_path.exists(),
            "file_size_bytes": output_path.stat().st_size if output_path.exists() else None,
        },
    )


def set_visual_color(feature_name: str, color_name: str, opacity: int | None = None) -> dict[str, Any]:
    rgb = NAMED_COLORS.get(color_name.lower())
    if rgb is None:
        return err(f"Unsupported color '{color_name}'.", error_code="UNSUPPORTED_COLOR")
    try:
        _catia, document, part = get_active_catia_part()
        shape, body, metadata = find_hybrid_shape_by_name(part, feature_name)
        selection = document.Selection
        selection.Clear()
        selection.Add(shape)
        selection.VisProperties.SetRealColor(int(rgb[0]), int(rgb[1]), int(rgb[2]), 1)
        if opacity is not None:
            selection.VisProperties.SetRealOpacity(int(opacity), 1)
        selection.Clear()
        return ok("Styled feature.", feature_name=feature_name, body_name=safe_name(body), metadata=metadata)
    except Exception as exc:
        return err(str(exc), error_code="STYLE_FAILED", feature_name=feature_name)


def validate_feature_health(feature_name: str, *, expected_kind: str = "feature", min_length: float | None = None) -> dict[str, Any]:
    report: dict[str, Any] = {"feature_name": feature_name, "expected_kind": expected_kind, "health_status": "unknown", "measurements": {}, "checks": {}}
    if feature_name.lower() in ORIGIN_PLANE_NAMES:
        report["health_status"] = "not_applicable"
        report["message"] = "Origin plane used directly."
        return report
    try:
        _catia, document, part = get_active_catia_part()
        feature, body, metadata = find_hybrid_shape_by_name(part, feature_name)
        report["found"] = True
        report["feature_metadata"] = metadata
        report["body_name"] = safe_name(body)
        ref = part.CreateReferenceFromObject(feature)
        report["checks"]["reference"] = True
        try:
            part.UpdateObject(feature)
            report["checks"]["update"] = True
        except Exception:
            part.Update()
            report["checks"]["update"] = True
            report["used_global_update"] = True
        measurement = measure_feature(document, part, feature)
        report["measurements"] = {k: v for k, v in measurement.items() if k in {"length", "area", "volume"}}
        if min_length is not None:
            length = report["measurements"].get("length")
            if length is None or float(length) <= float(min_length):
                report["health_status"] = "unhealthy"
                report["message"] = f"Length {length} is not greater than {min_length}."
                return report
        report["health_status"] = "healthy"
        report["message"] = "Feature health check completed."
        return report
    except Exception as exc:
        report["health_status"] = "unhealthy"
        report["message"] = str(exc)
        return report


def scan_surfaces(document: Any, part: Any, limit: int = 20) -> list[dict[str, Any]]:
    candidates: list[dict[str, Any]] = []
    for body in iter_hybrid_bodies(getattr(part, "HybridBodies", None)):
        body_name = safe_name(body) or ""
        shapes = getattr(body, "HybridShapes", None)
        if shapes is None:
            continue
        for index in range(1, count(shapes) + 1):
            try:
                shape = item(shapes, index)
            except Exception:
                continue
            measurement = measure_feature(document, part, shape)
            area = measurement.get("area")
            if isinstance(area, (int, float)) and float(area) > 1e-6:
                candidates.append({"kind": "surface", "body_name": body_name, "index": index, "name": measurement.get("name"), "area": float(area), "cog": measurement.get("cog")})
    candidates.sort(key=lambda row: float(row.get("area") or 0.0), reverse=True)
    return candidates[:limit]


def scan_bodies(document: Any, part: Any) -> list[dict[str, Any]]:
    candidates: list[dict[str, Any]] = []
    bodies = getattr(part, "Bodies", None)
    if bodies is None:
        return candidates
    for index in range(1, count(bodies) + 1):
        try:
            body = item(bodies, index)
        except Exception:
            continue
        name = safe_name(body)
        if not name:
            continue
        measurement = measure_feature(document, part, body)
        candidates.append({"kind": "body", "name": name, "index": index, "volume": measurement.get("volume"), "area": measurement.get("area"), "cog": measurement.get("cog"), "measurement_error": measurement.get("measurement_error")})
    candidates.sort(key=lambda row: (0 if row.get("cog") else 1, -(float(row.get("volume") or row.get("area") or 0.0))))
    return candidates


def scan_active_part(limit: int = 20) -> dict[str, Any]:
    _catia, document, part = get_active_catia_part()
    bodies = scan_bodies(document, part)
    surfaces = scan_surfaces(document, part, limit=limit)
    hybrid_body_names = [safe_name(body) for body in iter_hybrid_bodies(getattr(part, "HybridBodies", None))]
    return ok(
        "Scanned active CATPart.",
        document_name=safe_name(document),
        part_name=safe_name(part),
        hybrid_bodies=[name for name in hybrid_body_names if name],
        body_candidates=bodies,
        surface_candidates=surfaces,
    )


def build_attempts(candidate: dict[str, Any], args: argparse.Namespace, plane_definition: dict[str, Any]) -> list[dict[str, Any]]:
    if plane_definition["mode"] in {"equation", "point_normal"}:
        return [
            {
                "plane_mode": plane_definition["mode"],
                "base_plane_name": args.section_plane,
                "offset_distance": None,
                "reverse": False,
                "use_existing_plane": False,
                "use_positioned_plane": True,
                "plane_definition": plane_definition,
            }
        ]
    if plane_definition["mode"] == "offset_angle":
        return [
            {
                "plane_mode": "offset_angle",
                "base_plane_name": plane_definition["base_plane_name"],
                "offset_distance": plane_definition["offset_distance"],
                "reverse": plane_definition["reverse"],
                "axis_name": plane_definition["axis_name"],
                "angle_deg": plane_definition["angle_deg"],
                "angle_reverse": plane_definition["angle_reverse"],
                "use_existing_plane": False,
                "use_positioned_plane": True,
                "plane_definition": plane_definition,
            }
        ]

    section_plane = args.section_plane
    planes = ["yz", "zx", "xy"] if section_plane == "auto" else [section_plane.lower()]
    attempts: list[dict[str, Any]] = []
    for plane in planes:
        attempts.append({"base_plane_name": plane, "plane_name": plane, "offset_distance": 0.0, "reverse": False, "use_existing_plane": True})
    cog = candidate.get("cog")
    if cog and len(cog) >= 3:
        x, y, z = float(cog[0]), float(cog[1]), float(cog[2])
        offset_attempts = [
            ("yz", abs(x), x < 0.0),
            ("yz", abs(x), x >= 0.0),
            ("zx", abs(y), y < 0.0),
            ("xy", abs(z), z < 0.0),
        ]
        for plane, offset, reverse in offset_attempts:
            if section_plane == "auto" or section_plane.lower() == plane:
                attempts.append({"base_plane_name": plane, "offset_distance": offset, "reverse": reverse, "use_existing_plane": False})
    unique: list[dict[str, Any]] = []
    seen: set[tuple[str, float, bool, bool]] = set()
    for attempt in attempts:
        key = (str(attempt["base_plane_name"]), round(float(attempt["offset_distance"]), 6), bool(attempt["reverse"]), bool(attempt["use_existing_plane"]))
        if key not in seen:
            seen.add(key)
            unique.append(attempt)
    return unique


def candidate_label(candidate: dict[str, Any]) -> str:
    return str(candidate.get("name") or "")


def target_candidates(document: Any, part: Any, target_name: str | None, surface_limit: int) -> tuple[list[dict[str, Any]], list[dict[str, Any]], list[dict[str, Any]]]:
    surfaces = scan_surfaces(document, part, limit=surface_limit)
    bodies = scan_bodies(document, part)
    if target_name:
        try:
            shape, body, _metadata = find_hybrid_shape_by_name(part, target_name)
            measurement = measure_feature(document, part, shape)
            return [{"kind": "surface", "name": target_name, "body_name": safe_name(body), "area": measurement.get("area"), "cog": measurement.get("cog"), "provided_by_user": True}], surfaces, bodies
        except Exception:
            return [{"kind": "body_or_named_reference", "name": target_name, "provided_by_user": True}], surfaces, bodies
    candidates: list[dict[str, Any]] = []
    candidates.extend(bodies)
    candidates.extend(surfaces)
    return candidates, surfaces, bodies


def curve_length_from_result(result: dict[str, Any]) -> float | None:
    intersections = (result.get("raw_result") or {}).get("intersections") or []
    if not intersections:
        return None
    measurement = intersections[0].get("measurement") or {}
    length = measurement.get("length")
    return float(length) if isinstance(length, (int, float)) else None


def run_export(args: argparse.Namespace) -> dict[str, Any]:
    try:
        plane_definition = section_plane_definition_from_args(args)
    except Exception as exc:
        return err(str(exc), error_code="INVALID_SECTION_PLANE_DEFINITION")

    if args.dry_run:
        return ok(
            "Dry-run plan only. No CATIA connection or model edit was performed.",
            section_plane=args.section_plane,
            section_plane_definition=plane_definition,
            target_name=args.target_name,
            output_dir=str(resolve_output_dir(args)),
            steps=[
                "scan active CATPart",
                "create/use requested section plane" if plane_definition["mode"] != "auto" else "try origin plane intersections",
                "fallback to generated section plane when needed" if plane_definition["mode"] == "auto" else "intersect target with positioned section plane",
                "health check generated curve",
                "export curve to standalone CATPart",
            ],
        )
    if not args.user_confirmed:
        return err("Live run requires --user-confirmed.", error_code="LIVE_RUN_NOT_CONFIRMED")

    _catia, document, part = get_active_catia_part()
    candidates, surfaces, bodies = target_candidates(document, part, args.target_name, args.surface_limit)
    if not candidates:
        return err("No target candidate found in active CATPart.", error_code="NO_TARGET_CANDIDATE")

    run_id = args.run_id or f"section_export_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
    token = run_id.lower().replace("-", "_")
    output_dir = resolve_output_dir(args)
    export_path = output_dir / f"{run_id}_SectionResult.CATPart"
    base_plane_result_name = f"s10_section_plane_{token}_offset_001"
    base_curve_result_name = f"s20_section_curve_{token}_intersect_001"
    export_result_name = f"s90_section_result_{token}_extract_001"

    attempts_report: list[dict[str, Any]] = []
    accepted_curve: str | None = None
    accepted_plane: str | None = None
    accepted_target: dict[str, Any] | None = None
    accepted_plane_result: dict[str, Any] | None = None
    accepted_curve_result: dict[str, Any] | None = None
    attempt_counter = 0

    for candidate_index, candidate in enumerate(candidates, start=1):
        current_target_name = candidate_label(candidate)
        if not current_target_name:
            continue
        for attempt in build_attempts(candidate, args, plane_definition):
            attempt_counter += 1
            curve_name = base_curve_result_name if attempt_counter == 1 else f"{base_curve_result_name}_{attempt_counter:03d}"
            if attempt["use_existing_plane"]:
                plane_name = str(attempt["base_plane_name"])
                plane_result = ok(f"Using existing origin plane '{plane_name}'.", raw_result={"created_feature": False})
            else:
                plane_name = base_plane_result_name if attempt_counter == 1 else f"{base_plane_result_name}_{attempt_counter:03d}"
                try:
                    if attempt.get("plane_mode") in {"equation", "point_normal"}:
                        plane_result = create_section_plane_by_equation(
                            result_name=plane_name,
                            equation=list(attempt["plane_definition"]["equation"]),
                            target_body_name=args.work_body,
                            source_definition=attempt["plane_definition"],
                        )
                    else:
                        plane_result = create_section_plane_by_axis(
                            result_name=plane_name,
                            base_plane_name=str(attempt["base_plane_name"]),
                            offset_distance=float(attempt["offset_distance"] or 0.0),
                            reverse=bool(attempt["reverse"]),
                            axis_name=attempt.get("axis_name"),
                            angle_deg=attempt.get("angle_deg"),
                            angle_reverse=bool(attempt.get("angle_reverse")),
                            target_body_name=args.work_body,
                        )
                except Exception as exc:
                    plane_result = err(str(exc), error_code="PLANE_CREATE_FAILED")

            report_item = {
                **attempt,
                "target_name": current_target_name,
                "target_kind": candidate.get("kind"),
                "target_index": candidate_index,
                "plane_name": plane_name,
                "curve_name": curve_name,
                "plane_status": plane_result.get("status"),
                "status": "plane_failed",
                "plane_error": plane_result.get("message") if plane_result.get("status") != "success" else None,
            }
            if plane_result.get("status") == "success":
                try:
                    curve_result = create_section_intersection_curves(
                        section_plane_name=plane_name,
                        target_element_names=[current_target_name],
                        result_name=curve_name,
                        target_body_name=args.work_body,
                        extend_mode=args.extend_mode,
                    )
                except Exception as exc:
                    curve_result = err(str(exc), error_code="CURVE_CREATE_FAILED")
                length = curve_length_from_result(curve_result)
                report_item.update(
                    {
                        "curve_status": curve_result.get("status"),
                        "curve_length": length,
                        "curve_error": curve_result.get("message") if curve_result.get("status") != "success" else None,
                    }
                )
                if curve_result.get("status") == "success" and length is not None and length > args.min_length:
                    report_item["status"] = "accepted"
                    accepted_curve = curve_name
                    accepted_plane = plane_name
                    accepted_target = candidate
                    accepted_plane_result = plane_result
                    accepted_curve_result = curve_result
                    attempts_report.append(report_item)
                    break
                report_item["status"] = "curve_failed_or_too_short"
            attempts_report.append(report_item)
        if accepted_curve:
            break

    if not accepted_curve or not accepted_plane:
        report = err(
            "No section attempt produced a measurable curve above the minimum length.",
            error_code="SECTION_CURVE_NOT_FOUND",
            document_name=safe_name(document),
            part_name=safe_name(part),
            target_name=args.target_name,
            section_plane_definition=plane_definition,
            attempts=attempts_report,
            target_candidates=candidates[:12],
            surface_candidates=surfaces,
            body_candidates=bodies,
        )
        persist_report(output_dir, report)
        return report

    if args.color:
        if accepted_plane.lower() not in ORIGIN_PLANE_NAMES:
            set_visual_color(accepted_plane, "yellow", opacity=110)
        set_visual_color(accepted_curve, "orange")

    if args.no_export:
        export_result = ok("Export disabled by --no-export.", exported_path=None, raw_result={})
    else:
        export_result = export_section_curves_as_catpart(
            curve_names=[accepted_curve],
            result_name=export_result_name,
            output_path=export_path,
            target_body_name=args.export_body,
            close_exported_document=args.close_exported_document,
            reactivate_source_document=True,
        )

    health = [
        validate_feature_health(accepted_plane, expected_kind="origin_plane" if accepted_plane.lower() in ORIGIN_PLANE_NAMES else "surface"),
        validate_feature_health(accepted_curve, expected_kind="curve", min_length=args.min_length),
    ]
    report = ok(
        "Section curve export completed.",
        document_name=safe_name(document),
        part_name=safe_name(part),
        target_name=candidate_label(accepted_target or {}),
        selected_target=accepted_target,
        section_plane_name=accepted_plane,
        section_plane_definition=plane_definition,
        section_curve_name=accepted_curve,
        exported_catpart=str(export_path) if not args.no_export else None,
        attempts=attempts_report,
        plane_result=accepted_plane_result,
        curve_result=accepted_curve_result,
        export_result=export_result,
        health=health,
        target_candidates=candidates[:12],
        surface_candidates=surfaces,
        body_candidates=bodies,
    )
    persist_report(output_dir, report)
    return report


def resolve_output_dir(args: argparse.Namespace) -> Path:
    run_id = getattr(args, "run_id", None) or "dry_run"
    raw = Path(getattr(args, "output_dir", None) or (TOOL_ROOT / "output" / run_id))
    if not raw.is_absolute():
        raw = TOOL_ROOT / raw
    return raw.resolve()


def persist_report(output_dir: Path, report: dict[str, Any]) -> None:
    output_dir.mkdir(parents=True, exist_ok=True)
    report_json = output_dir / "section_curve_export_report.json"
    report_md = output_dir / "section_curve_export_report.md"
    report["report_json"] = str(report_json)
    report["report_md"] = str(report_md)
    write_json(report_json, report)
    write_markdown(report_md, report)


def activate_document(document_name: str) -> dict[str, Any]:
    catia, active_document, _part = get_active_catia_part()
    active_before = safe_name(active_document)
    documents = catia.Documents
    target = None
    available: list[str] = []
    for index in range(1, count(documents) + 1):
        document = item(documents, index)
        name = safe_name(document) or ""
        available.append(name)
        if name == document_name:
            target = document
    if target is None:
        return err(f"Document '{document_name}' is not open.", error_code="DOCUMENT_NOT_FOUND", active_before=active_before, available_documents=available)
    target.Activate()
    return ok("Document activated.", active_before=active_before, active_after=safe_name(catia.ActiveDocument), available_documents=available)


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Standalone CATIA section curve export tool.")
    subparsers = parser.add_subparsers(dest="command", required=True)

    scan_parser = subparsers.add_parser("scan", help="Scan active CATPart target candidates.")
    scan_parser.add_argument("--limit", type=int, default=20)
    scan_parser.add_argument("--output")

    run_parser = subparsers.add_parser("run", help="Create section curve and export it to a standalone CATPart.")
    run_parser.add_argument("--target-name")
    run_parser.add_argument("--section-plane", default="auto", help="auto, xy, yz, zx, or an existing plane feature name.")
    run_parser.add_argument("--offset-distance", type=float, help="Create a section plane offset from --section-plane. Signed values are accepted.")
    run_parser.add_argument("--reverse", action="store_true", help="Reverse offset direction for --offset-distance.")
    run_parser.add_argument("--through-point", help="Create a positioned plane through point X,Y,Z. If --normal is omitted, the normal comes from --section-plane.")
    run_parser.add_argument("--normal", help="Plane normal vector NX,NY,NZ for --through-point.")
    run_parser.add_argument("--plane-equation", help="Create a positioned plane from A,B,C,D using A*x + B*y + C*z = D.")
    run_parser.add_argument("--axis-name", help="Rotate the offset/base plane around this CATIA axis/line feature.")
    run_parser.add_argument("--angle-deg", type=float, help="Rotation angle in degrees for --axis-name.")
    run_parser.add_argument("--angle-reverse", action="store_true", help="Reverse angle orientation for --axis-name/--angle-deg.")
    run_parser.add_argument("--run-id")
    run_parser.add_argument("--output-dir")
    run_parser.add_argument("--min-length", type=float, default=0.001)
    run_parser.add_argument("--surface-limit", type=int, default=8)
    run_parser.add_argument("--work-body", default="section_results")
    run_parser.add_argument("--export-body", default="SectionResult")
    run_parser.add_argument("--extend-mode", action="store_true")
    run_parser.add_argument("--color", action="store_true")
    run_parser.add_argument("--no-export", action="store_true")
    run_parser.add_argument("--close-exported-document", action="store_true")
    run_parser.add_argument("--dry-run", action="store_true")
    run_parser.add_argument("--user-confirmed", action="store_true")

    health_parser = subparsers.add_parser("health", help="Validate a generated CATIA feature.")
    health_parser.add_argument("feature_name")
    health_parser.add_argument("--expected-kind", default="feature")
    health_parser.add_argument("--min-length", type=float)

    activate_parser = subparsers.add_parser("activate", help="Activate an already open CATIA document by name.")
    activate_parser.add_argument("document_name")
    return parser


def main() -> int:
    parser = build_parser()
    args = parser.parse_args()
    try:
        if args.command == "scan":
            payload = scan_active_part(limit=args.limit)
            if args.output:
                write_json(Path(args.output), payload)
        elif args.command == "run":
            payload = run_export(args)
        elif args.command == "health":
            payload = validate_feature_health(args.feature_name, expected_kind=args.expected_kind, min_length=args.min_length)
        elif args.command == "activate":
            payload = activate_document(args.document_name)
        else:
            parser.error("Unknown command.")
            return 2
    except Exception as exc:
        payload = err(str(exc), error_code="UNHANDLED_EXCEPTION")
    print_json(payload)
    return 0 if payload.get("status") == "success" or payload.get("health_status") == "healthy" else 1


if __name__ == "__main__":
    raise SystemExit(main())
