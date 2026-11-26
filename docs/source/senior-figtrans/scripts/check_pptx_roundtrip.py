"""Read-only comparison of saved native PPTX nodes and connector geometry.

Combine these DrawingML measurements with native-host receipts and render review.
Usage: python check_pptx_roundtrip.py BEFORE.pptx AFTER.pptx --moved-name receiver --output proof.json
Optional --moved-name matches a semantic shape name. Coordinates are slide points.
Nested group transforms, including scale, flip and rotation, are applied explicitly.
Copyright (c) 2026 NCEPU-Senior. Author: XiaoN.
"""
from __future__ import annotations

import argparse
from collections import Counter
from hashlib import sha256
import json
import math
from pathlib import Path
import re
from zipfile import ZipFile
import xml.etree.ElementTree as ET


NS = {"p": "http://schemas.openxmlformats.org/presentationml/2006/main",
      "a": "http://schemas.openxmlformats.org/drawingml/2006/main"}
EMU_PER_POINT = 12700
CONNECTOR_PRESETS = {"line", "straightConnector1", "bentConnector2", "bentConnector3",
                     "bentConnector4", "bentConnector5"}


def apply(matrix, coordinates):
    a, b, c, d, e, f = matrix
    x, y = coordinates
    return [round(a*x+c*y+e, 6), round(b*x+d*y+f, 6)]


def multiply(left, right):
    a, b, c, d, e, f = left
    g, h, i, j, k, l = right
    return (a*g+c*h, b*g+d*h, a*i+c*j, b*i+d*j, a*k+c*l+e, b*k+d*l+f)


def shape_matrix(transform):
    angle = math.radians(transform["rotation_deg"])
    sx = -1 if transform["flip_h"] else 1
    sy = -1 if transform["flip_v"] else 1
    a, b = sx*math.cos(angle), sx*math.sin(angle)
    c, d = -sy*math.sin(angle), sy*math.cos(angle)
    w, h = transform["width"], transform["height"]
    return (a, b, c, d, transform["left"]+w/2-a*w/2-c*h/2,
            transform["top"]+h/2-b*w/2-d*h/2)


def group_child_matrix(element, transform):
    node = element.find("p:grpSpPr/a:xfrm", NS)
    offset, extent = node.find("a:chOff", NS), node.find("a:chExt", NS)
    if offset is None or extent is None:
        raise ValueError("Group has no child coordinate system")
    cx, cy = int(extent.get("cx"))/EMU_PER_POINT, int(extent.get("cy"))/EMU_PER_POINT
    if cx <= 0 or cy <= 0:
        raise ValueError("Group child coordinate extent must be positive")
    sx, sy = transform["width"]/cx, transform["height"]/cy
    return (sx, 0, 0, sy, -int(offset.get("x"))/EMU_PER_POINT*sx,
            -int(offset.get("y"))/EMU_PER_POINT*sy)


def point(transform: dict, u: float, v: float) -> list[float]:
    """Apply the saved shape flip and clockwise rotation about its centre."""
    width, height = transform["width"], transform["height"]
    x = width-u if transform["flip_h"] else u
    y = height-v if transform["flip_v"] else v
    x -= width/2
    y -= height/2
    angle = math.radians(transform["rotation_deg"])
    return [round(transform["left"]+width/2+x*math.cos(angle)-y*math.sin(angle), 6),
            round(transform["top"]+height/2+x*math.sin(angle)+y*math.cos(angle), 6)]


def transform_of(element, group=False):
    path = "p:grpSpPr/a:xfrm" if group else "p:spPr/a:xfrm"
    node = element.find(path, NS)
    if node is None:
        node = element.find("p:xfrm", NS)
    if node is None:
        return None
    off, ext = node.find("a:off", NS), node.find("a:ext", NS)
    if off is None or ext is None:
        return None
    return {"left": int(off.get("x"))/EMU_PER_POINT, "top": int(off.get("y"))/EMU_PER_POINT,
            "width": int(ext.get("cx"))/EMU_PER_POINT, "height": int(ext.get("cy"))/EMU_PER_POINT,
            "rotation_deg": int(node.get("rot", 0))/60000,
            "flip_h": node.get("flipH") in {"1", "true"}, "flip_v": node.get("flipV") in {"1", "true"}}


def shape_text(element):
    paragraphs = []
    for paragraph in element.findall("p:txBody/a:p", NS):
        fragments = []
        for node in paragraph.iter():
            if node.tag == "{"+NS["a"]+"}t":
                fragments.append(node.text or "")
            elif node.tag == "{"+NS["a"]+"}br":
                fragments.append("\n")
        paragraphs.append("".join(fragments))
    return "\n".join(paragraphs)


def inspect(path: Path) -> dict:
    slides = []
    with ZipFile(path) as package:
        parts = sorted((name for name in package.namelist() if re.fullmatch(r"ppt/slides/slide\d+\.xml", name)),
                       key=lambda name: int(re.search(r"slide(\d+)", name).group(1)))
        for part in parts:
            root = ET.fromstring(package.read(part))
            objects = []

            def walk(tree, parents=(), parent_matrix=(1, 0, 0, 1, 0, 0)):
                for element in tree:
                    kind = element.tag.rsplit("}", 1)[-1]
                    if kind not in {"sp", "cxnSp", "grpSp", "pic", "graphicFrame"}:
                        continue
                    identity = element.find(".//p:cNvPr", NS)
                    if identity is None:
                        continue
                    record = {"id": identity.get("id"), "name": identity.get("name"), "kind": kind,
                              "parent_groups": list(parents), "text": shape_text(element)}
                    transform = transform_of(element, kind == "grpSp")
                    record["transform_pt"] = transform
                    record["coordinate_scope"] = "slide (group transforms applied)"
                    geometry = element.find("p:spPr/a:prstGeom", NS)
                    record["preset"] = geometry.get("prst") if geometry is not None else None
                    custom = element.find("p:spPr/a:custGeom", NS)
                    if custom is not None:
                        record["custom_geometry"] = {
                            "sha256": sha256(ET.tostring(custom)).hexdigest(),
                            "cubic_segments": len(custom.findall(".//a:cubicBezTo", NS)),
                            "line_segments": len(custom.findall(".//a:lnTo", NS)),
                        }
                    matrix = None
                    if transform is not None:
                        matrix = multiply(parent_matrix, shape_matrix(transform))
                        w, h = transform["width"], transform["height"]
                        record["centre_pt"] = apply(matrix, (w/2, h/2))
                        record["corners_pt"] = [apply(matrix, p) for p in ((0, 0), (w, 0), (w, h), (0, h))]
                    if kind == "cxnSp":
                        attachments = {}
                        for endpoint, tag in (("begin", "stCxn"), ("end", "endCxn")):
                            nodes = element.findall("p:nvCxnSpPr/p:cNvCxnSpPr/a:"+tag, NS)
                            if len(nodes) == 1:
                                attachments[endpoint] = {"shape_id": nodes[0].get("id"),
                                    "xml_site_index": int(nodes[0].get("idx")),
                                    "powerpoint_site_number": int(nodes[0].get("idx"))+1}
                            elif len(nodes) > 1:
                                raise ValueError(f"Repeated {tag} in {part}/{record['name']}")
                        record["attachments"] = attachments
                        if matrix is not None and record["preset"] in CONNECTOR_PRESETS:
                            record["endpoints_pt"] = {"begin": apply(matrix, (0, 0)),
                                "end": apply(matrix, (transform["width"], transform["height"]))}
                        else:
                            record["endpoints_pt"] = None
                    objects.append(record)
                    if kind == "grpSp":
                        if matrix is None:
                            raise ValueError(f"Group has no transform: {record['name']}")
                        walk(element, parents+(record["id"],), multiply(matrix, group_child_matrix(element, transform)))

            tree = root.find("p:cSld/p:spTree", NS)
            if tree is not None:
                walk(tree)
            identities = Counter(row["id"] for row in objects)
            duplicates = [identity for identity, count in identities.items() if count > 1]
            if duplicates:
                raise ValueError(f"Duplicate shape IDs in {part}: {duplicates}")
            by_id = {row["id"]: row for row in objects}
            for row in objects:
                for endpoint in row.get("attachments", {}).values():
                    target = by_id.get(endpoint["shape_id"])
                    if target is None:
                        raise ValueError(f"Missing attachment target in {part}/{row['name']}")
                    endpoint["shape_name"] = target["name"]
            kinds = Counter(row["kind"] for row in objects)
            slides.append({"part": part, "counts": dict(kinds), "text_bearing_objects": sum(bool(row["text"]) for row in objects),
                           "two_end_attachments": sum(set(row.get("attachments", {})) == {"begin", "end"} for row in objects),
                           "objects": objects})
    return {"file": str(path.resolve()), "sha256": sha256(path.read_bytes()).hexdigest(), "slides": slides}


def subtract(after, before):
    return [round(a-b, 6) for a, b in zip(after, before)]


def close(actual, expected, tolerance):
    return all(abs(a-b) <= tolerance for a, b in zip(actual, expected))


def compare(before: dict, after: dict, moved_ids=(), moved_names=(), tolerance=.02,
            expected_text=None, expected_delta=None) -> dict:
    errors, moved, text_changes, connector_checks = [], [], [], []
    requested_ids, requested_names = set(moved_ids), set(moved_names)
    found_ids, found_names = set(), set()
    for old_slide, new_slide in zip(before["slides"], after["slides"]):
        old = {row["id"]: row for row in old_slide["objects"]}
        new = {row["id"]: row for row in new_slide["objects"]}
        part = old_slide["part"]
        if old_slide["counts"] != new_slide["counts"]:
            errors.append(f"Structural counts changed on {part}")
        if set(old) != set(new):
            errors.append(f"Shape ID roster changed on {part}")
        selected = set()
        for identity, previous in old.items():
            current = new.get(identity)
            if current is None:
                continue
            if previous["text"] != current["text"]:
                text_changes.append({"slide": part, "id": identity, "name": previous["name"],
                                     "before": previous["text"], "after": current["text"]})
            if previous["name"] != current["name"] or previous["parent_groups"] != current["parent_groups"]:
                errors.append(f"Shape name or group membership changed: {part}/{previous['name']}")
            if identity in requested_ids or previous["name"] in requested_names:
                found_ids.add(identity); found_names.add(previous["name"]); selected.add(identity)
                record = {"slide": part, "id": identity, "name": previous["name"],
                          "before": previous["transform_pt"], "after": current["transform_pt"]}
                if "centre_pt" in previous and "centre_pt" in current:
                    record["delta_pt"] = subtract(current["centre_pt"], previous["centre_pt"])
                    if close(record["delta_pt"], [0, 0], tolerance):
                        errors.append(f"Selected moved object did not move: {part}/{previous['name']}")
                    if expected_delta is not None and not close(record["delta_pt"], expected_delta, tolerance):
                        errors.append(f"Selected object did not move by expected delta: {part}/{previous['name']}")
                else:
                    errors.append(f"Selected object {identity} has no measurable saved transform")
                moved.append(record)
        moving = selected | {row["id"] for row in old.values() if selected.intersection(row["parent_groups"])}
        for identity in moving:
            previous, current = old[identity], new.get(identity)
            if current is None or "corners_pt" not in previous or "corners_pt" not in current:
                errors.append(f"Moved member has no measurable transform: {part}/{previous['name']}")
                continue
            owners = selected.intersection(previous["parent_groups"]+[identity])
            owner = old[next(iter(owners))]
            expected = subtract(new[owner["id"]]["centre_pt"], owner["centre_pt"])
            if any(not close(subtract(a, b), expected, tolerance)
                   for a, b in zip(current["corners_pt"], previous["corners_pt"])):
                errors.append(f"Group member did not translate with selected receiver: {part}/{previous['name']}")
            if previous.get("custom_geometry") != current.get("custom_geometry"):
                errors.append(f"Moved custom geometry changed: {part}/{previous['name']}")
        if requested_ids or requested_names:
            for identity, previous in old.items():
                current = new.get(identity)
                if identity in moving or current is None or previous["kind"] == "cxnSp":
                    continue
                if "corners_pt" in previous and "corners_pt" in current:
                    if any(not close(a, b, tolerance) for a, b in zip(current["corners_pt"], previous["corners_pt"])):
                        errors.append(f"Unselected object moved: {part}/{previous['name']}")
        for identity, previous in old.items():
            if previous["kind"] != "cxnSp" or identity not in new:
                continue
            current = new[identity]
            attachments_same = previous["attachments"] == current["attachments"]
            record = {"slide": part, "id": identity, "name": previous["name"],
                      "before_attachments": previous["attachments"], "after_attachments": current["attachments"],
                      "attachments_preserved": attachments_same, "endpoint_checks": []}
            if not attachments_same:
                errors.append(f"Attachment IDs, names or sites changed on {part}/{previous['name']}")
            for endpoint, binding in previous["attachments"].items():
                target_id = binding["shape_id"]
                if target_id not in old or target_id not in new:
                    continue
                if not previous["endpoints_pt"] or not current["endpoints_pt"]:
                    record["endpoint_checks"].append({"endpoint": endpoint, "geometry": "unsupported preset or grouped coordinates"})
                    errors.append(f"Attached endpoint needs explicit path interpretation: {previous['name']}/{endpoint}")
                    continue
                node_before, node_after = old[target_id], new[target_id]
                if "centre_pt" not in node_before or "centre_pt" not in node_after:
                    continue
                endpoint_delta = subtract(current["endpoints_pt"][endpoint], previous["endpoints_pt"][endpoint])
                node_delta = subtract(node_after["centre_pt"], node_before["centre_pt"])
                pure_translation = all(close(subtract(a, b), node_delta, tolerance)
                                       for a, b in zip(node_after["corners_pt"], node_before["corners_pt"]))
                follows = pure_translation and close(endpoint_delta, node_delta, tolerance)
                record["endpoint_checks"].append({"endpoint": endpoint, "target_id": target_id,
                    "target_name": binding["shape_name"], "target_selected": target_id in moving,
                    "before_pt": previous["endpoints_pt"][endpoint], "after_pt": current["endpoints_pt"][endpoint],
                    "endpoint_delta_pt": endpoint_delta, "node_delta_pt": node_delta,
                    "pure_translation": pure_translation, "follows_target": follows})
                if pure_translation and not follows:
                    errors.append(f"Endpoint did not follow its translated or stationary node: {part}/{previous['name']}/{endpoint}")
                elif target_id in moving and not pure_translation:
                    errors.append(f"Selected node also changed size or rotation: {part}/{node_before['name']}")
                if (requested_ids or requested_names) and target_id not in moving and not close(endpoint_delta, [0, 0], tolerance):
                    errors.append(f"Endpoint on unselected target moved: {part}/{previous['name']}/{endpoint}")
            connector_checks.append(record)
    if len(before["slides"]) != len(after["slides"]):
        errors.append("Slide count changed")
    if requested_ids-found_ids or requested_names-found_names:
        errors.append(f"Requested moved objects missing: IDs {sorted(requested_ids-found_ids)}, names {sorted(requested_names-found_names)}")
    for name, value in (expected_text or {}).items():
        matches = [row for slide in after["slides"] for row in slide["objects"] if row["name"] == name]
        if len(matches) != 1 or matches[0]["text"] != value:
            errors.append(f"Expected native text was not preserved uniquely: {name}")
    return {"scope": "Saved DrawingML with nested group transforms; combine with native host receipts and rendered previews.",
            "status": "CHECKS_PASS_REVIEW_REQUIRED" if not errors else "FAIL", "errors": errors,
            "tolerance_pt": tolerance, "moved_nodes": moved, "text_changes": text_changes,
            "connector_checks": connector_checks, "before": before, "after": after}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("before", type=Path)
    parser.add_argument("after", type=Path)
    parser.add_argument("--moved-id", action="append", default=[])
    parser.add_argument("--moved-name", action="append", default=[])
    parser.add_argument("--tolerance-pt", type=float, default=.02)
    parser.add_argument("--expect-text", action="append", default=[], metavar="NAME=VALUE")
    parser.add_argument("--expected-delta-pt", type=float, nargs=2, metavar=("DX", "DY"))
    parser.add_argument("--output", type=Path)
    args = parser.parse_args()
    if not math.isfinite(args.tolerance_pt) or args.tolerance_pt <= 0:
        parser.error("tolerance must be positive and finite")
    text_expectations = {}
    for item in args.expect_text:
        if "=" not in item or not item.split("=", 1)[0]:
            parser.error("--expect-text must be NAME=VALUE")
        name, value = item.split("=", 1)
        if name in text_expectations:
            parser.error("--expect-text names must be unique")
        text_expectations[name] = value
    if args.expected_delta_pt is not None and not all(math.isfinite(x) for x in args.expected_delta_pt):
        parser.error("expected delta must be finite")
    result = compare(inspect(args.before), inspect(args.after), args.moved_id, args.moved_name, args.tolerance_pt,
                     text_expectations, args.expected_delta_pt)
    if args.output:
        args.output.parent.mkdir(parents=True, exist_ok=True)
        args.output.write_text(json.dumps(result, ensure_ascii=False, indent=2)+"\n", encoding="utf-8")
    print(json.dumps({key: value for key, value in result.items() if key not in {"before", "after", "connector_checks"}}, ensure_ascii=True))
    return int(bool(result["errors"]))


if __name__ == "__main__":
    raise SystemExit(main())
