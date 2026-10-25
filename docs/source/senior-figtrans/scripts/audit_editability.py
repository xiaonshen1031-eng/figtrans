#!/usr/bin/env python3
"""Inspect PPTX/VSDX editability evidence without Office or third-party modules.

Copyright (c) 2026 NCEPU-Senior
Author: XiaoN

This is a package-structure audit, not a visual or semantic quality certificate.
Run: python audit_editability.py figure.pptx --require-native --output audit.json
Exit codes: 0 structural pass, 1 failed audit, 2 unreadable input or invalid usage.
"""
from __future__ import annotations

import argparse
from collections import Counter
import json
from pathlib import Path
import posixpath
import re
import sys
from urllib.parse import unquote, urlsplit
from zipfile import BadZipFile, ZipFile
import xml.etree.ElementTree as ET

NS = {
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "c": "http://schemas.openxmlformats.org/drawingml/2006/chart",
    "m": "http://schemas.openxmlformats.org/officeDocument/2006/math",
    "v": "http://schemas.microsoft.com/office/visio/2012/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
}


def local_name(tag):
    return tag.rsplit("}", 1)[-1]


def natural_key(value):
    return [int(s) if s.isdigit() else s for s in re.split(r"(\d+)", value)]


def xml_part(archive, name):
    return ET.fromstring(archive.read(name))


def select_alternate_content(root):
    """Count logical objects once instead of counting Choice plus fallback copies.

    Select the first declared Choice (or Fallback if there is no Choice). This is
    an inventory convention, not proof that a particular renderer supports it.
    """
    ns = "http://schemas.openxmlformats.org/markup-compatibility/2006"
    selected = 0
    for parent in list(root.iter()):
        for child in list(parent):
            if child.tag != "{" + ns + "}AlternateContent":
                continue
            choice = child.find("{" + ns + "}Choice")
            if choice is None:
                choice = child.find("{" + ns + "}Fallback")
            position = list(parent).index(child)
            parent.remove(child)
            if choice is not None:
                for offset, member in enumerate(list(choice)):
                    parent.insert(position + offset, member)
            selected += 1
    return selected


def relationship_source(name):
    if name == "_rels/.rels":
        return ""
    directory, filename = posixpath.split(name)
    if not directory.endswith("/_rels") or not filename.endswith(".rels"):
        return None
    return posixpath.join(directory[:-6], filename[:-5])


def relationships(archive, names):
    """Validate every internal relationship target; do not fetch external links."""
    by_source, failures, external = {}, [], []
    for name in sorted(n for n in names if n.endswith(".rels")):
        source = relationship_source(name)
        if source is None:
            continue
        relation_map = {}
        try:
            root = xml_part(archive, name)
        except ET.ParseError as exc:
            failures.append({"code": "invalid_relationship_xml", "part": name, "detail": str(exc)})
            continue
        for rel in root:
            rid, target = rel.get("Id"), rel.get("Target", "")
            item = {"id": rid, "type": rel.get("Type", ""), "target": target}
            if rid in relation_map:
                failures.append({"code": "duplicate_relationship_id", "part": name, "id": rid})
            if rel.get("TargetMode", "").lower() == "external":
                item["external"] = True
                external.append({"source": source, **item})
            else:
                path = unquote(urlsplit(target.replace("\\", "/")).path)
                resolved = posixpath.normpath(path.lstrip("/") if path.startswith("/") else
                                             posixpath.join(posixpath.dirname(source), path))
                item["resolved_part"] = resolved
                item["exists"] = resolved in names
                if not item["exists"]:
                    failures.append({"code": "missing_relationship_target", "part": name,
                                     "id": rid, "target": target, "resolved_part": resolved})
            relation_map[rid] = item
        by_source[source] = relation_map
    return by_source, failures, external


def ref_relationships(element, relation_map, part, failures):
    refs = []
    for descendant in element.iter():
        for attr, rid in descendant.attrib.items():
            if attr.startswith("{" + NS["r"] + "}"):
                if rid not in relation_map:
                    failures.append({"code": "missing_relationship_id", "part": part,
                                     "id": rid, "element": local_name(descendant.tag)})
                refs.append({"id": rid, **relation_map.get(rid, {})})
    return refs


def shape_name(element):
    nv = element.find(".//p:cNvPr", NS)
    return {"id": nv.get("id"), "name": nv.get("name", ""),
            "description": nv.get("descr", "")} if nv is not None else {}


def drawing_bounds(element):
    xfrm = element.find("p:spPr/a:xfrm", NS)
    if xfrm is None:
        xfrm = element.find("p:xfrm", NS)
    if xfrm is None:
        return None
    off, ext = xfrm.find("a:off", NS), xfrm.find("a:ext", NS)
    if off is None or ext is None:
        return None
    try:
        return {"x": int(off.get("x", 0)), "y": int(off.get("y", 0)),
                "width": int(ext.get("cx", 0)), "height": int(ext.get("cy", 0)),
                "coordinate_unit": "EMU", "rotation_60000deg": int(xfrm.get("rot", 0))}
    except ValueError:
        return None


def coverage(bounds, width, height):
    if not bounds or width <= 0 or height <= 0:
        return None
    x, y, w, h = (bounds[k] for k in ("x", "y", "width", "height"))
    overlap_w = max(0, min(width, x + w) - max(0, x))
    overlap_h = max(0, min(height, y + h) - max(0, y))
    return round(overlap_w * overlap_h / (width * height), 6)


def pptx_inventory(archive, names, by_source, failures):
    presentation = xml_part(archive, "ppt/presentation.xml")
    size = presentation.find("p:sldSz", NS)
    width, height = (int(size.get(k, 0)) for k in ("cx", "cy")) if size is not None else (0, 0)
    pages = []
    for part in sorted((n for n in names if re.fullmatch(r"ppt/slides/slide\d+\.xml", n)), key=natural_key):
        root = xml_part(archive, part)
        relations = by_source.get(part, {})
        ref_relationships(root, relations, part, failures)
        alternate_content = select_alternate_content(root)
        items, counts, texts, full_page, connections = [], Counter(), [], [], []

        def walk(parent, ancestor_groups=()):
            for element in parent:
                kind = local_name(element.tag)
                if kind not in {"sp", "pic", "cxnSp", "grpSp", "graphicFrame"}:
                    continue
                identity = shape_name(element)
                item = {"kind": kind, **identity, "parent_groups": list(ancestor_groups)}
                bounds = drawing_bounds(element)
                if bounds:
                    item["bounds"] = bounds
                if kind == "grpSp":
                    counts["groups"] += 1
                    items.append(item)
                    walk(element, ancestor_groups + (identity.get("id"),))
                    continue
                label = "".join(
                    e.text or "" for e in element.findall(".//a:t", NS))
                if label:
                    item["text"] = label
                    texts.append({**identity, "text": label})
                    counts["text_bearing_objects"] += 1
                omml = element.findall(".//m:oMath", NS)
                if omml:
                    counts["omml_equations"] += len(omml)
                    item["equation_kind"] = "Office Math (OMML), not MathType OLE"
                if kind == "sp":
                    counts["native_shapes"] += 1
                    textbox = element.find("p:nvSpPr/p:cNvSpPr", NS)
                    is_textbox = textbox is not None and textbox.get("txBox") in {"1", "true"}
                    item["is_textbox"] = is_textbox
                    area = coverage(bounds, width, height) if not ancestor_groups else None
                    item["background_candidate"] = bool(area is not None and area >= .85 and not label)
                    if not is_textbox and not item["background_candidate"]:
                        counts["non_background_geometry"] += 1
                elif kind == "pic":
                    counts["pictures"] += 1
                    area = coverage(bounds, width, height) if not ancestor_groups else None
                    item["page_area_coverage"] = area
                    if area is not None and area >= .70:
                        full_page.append({**identity, "page_area_coverage": area})
                    item["editability"] = "picture asset; internal pixels/vector artwork unverified"
                    if ancestor_groups:
                        item["coverage_note"] = "group coordinate transform not evaluated; inspect renderer"
                elif kind == "cxnSp":
                    counts["connectors"] += 1
                    ends = []
                    for endpoint in ("stCxn", "endCxn"):
                        node = element.find(".//a:" + endpoint, NS)
                        if node is not None:
                            ends.append({"endpoint": endpoint, "shape_id": node.get("id"), "site": node.get("idx")})
                    item["attachments"] = ends
                    item["attachment_mode"] = "both_ends" if len(ends) == 2 else "one_end" if ends else "unattached"
                    connections.append({**identity, "attachments": ends})
                    counts["attached_connectors"] += int(len(ends) == 2)
                if kind == "graphicFrame":
                    counts["graphic_frames"] += 1
                tables = element.findall(".//a:tbl", NS)
                charts = element.findall(".//c:chart", NS)
                counts["native_tables"] += len(tables)
                counts["chart_references"] += len(charts)
                if tables:
                    item["table_count"] = len(tables)
                if charts:
                    item["chart_count"] = len(charts)
                    item["chart_editability_note"] = "chart part referenced; data fidelity needs source comparison"
                oles = element.findall(".//p:oleObj", NS)
                if oles:
                    item["ole_objects"] = []
                for ole in oles:
                    prog_id = ole.get("progId", "")
                    declared_math = prog_id.lower().startswith("equation.dsmt")
                    counts["ole_objects"] += 1
                    counts["mathtype_declared_ole"] += int(declared_math)
                    counts["generic_ole"] += int(not declared_math)
                    item["ole_objects"].append({"prog_id": prog_id,
                        "mathtype_declaration": declared_math,
                        "native_mtef_verified": False,
                        "verification": "ProgID alone does not verify OLE payload or openability"})
                items.append(item)

        tree = root.find("p:cSld/p:spTree", NS)
        if tree is not None:
            walk(tree)
        all_ids = {i.get("id") for i in items if i.get("id")}
        for conn in connections:
            for end in conn["attachments"]:
                if end["shape_id"] not in all_ids:
                    failures.append({"code": "missing_connector_shape", "part": part,
                                     "connector": conn.get("name"), **end})
        pages.append({"part": part, "counts": dict(counts), "objects": items,
                      "readable_text": texts, "connections": connections,
                      "full_page_picture_candidates": full_page,
                      "alternate_content_selections": alternate_content})
    return {"canvas": {"width": width, "height": height, "unit": "EMU"}, "pages": pages}


def visio_inventory(archive, names, by_source, failures):
    pages, sizes = [], {}
    if "visio/pages/pages.xml" in names:
        pages_root = xml_part(archive, "visio/pages/pages.xml")
        for page in pages_root.findall("v:Page", NS):
            rel = page.find("v:Rel", NS)
            rid = rel.get("{" + NS["r"] + "}id") if rel is not None else None
            target = by_source.get("visio/pages/pages.xml", {}).get(rid, {}).get("resolved_part")
            sheet = page.find("v:PageSheet", NS)
            cells = {c.get("N"): c.get("V") for c in sheet.findall("v:Cell", NS)} if sheet is not None else {}
            sizes[target] = {"name": page.get("NameU", page.get("Name", "")), "width": cells.get("PageWidth"),
                             "height": cells.get("PageHeight"), "unit": "inch (internal Visio units)"}
    for part in sorted((n for n in names if re.fullmatch(r"visio/pages/page\d+\.xml", n)), key=natural_key):
        root = xml_part(archive, part)
        relations = by_source.get(part, {})
        ref_relationships(root, relations, part, failures)
        counts, objects, texts = Counter(), [], []

        def walk(shapes, ancestor_groups=()):
            for shape in shapes.findall("v:Shape", NS):
                identity = {"id": shape.get("ID"), "name": shape.get("NameU", shape.get("Name", ""))}
                kind = shape.get("Type", "Shape")
                cells = {c.get("N"): {"value": c.get("V"), "formula": c.get("F")} for c in shape.findall("v:Cell", NS)}
                item = {**identity, "kind": kind, "parent_groups": list(ancestor_groups), "cells": cells}
                text_element = shape.find("v:Text", NS)
                label = "".join(text_element.itertext()) if text_element is not None else ""
                if label:
                    texts.append({**identity, "text": label})
                    item["text"] = label
                    counts["text_bearing_objects"] += 1
                foreign = shape.find("v:ForeignData", NS)
                if foreign is not None or kind == "Foreign":
                    counts["foreign_objects"] += 1
                    attrs = dict(foreign.attrib) if foreign is not None else {}
                    item["foreign_data"] = attrs
                    item["editability"] = "foreign payload; native internal editability unverified"
                    if str(attrs.get("ForeignType", "")).lower() == "object":
                        counts["ole_candidates"] += 1
                        item["equation_verification"] = "unknown OLE class; MathType/MTEF requires payload and application inspection"
                    else:
                        counts["pictures"] += 1
                    item["page_coverage_note"] = "inspect rendered page; masters, groups, rotation and inherited cells can alter bounds"
                elif kind == "Group":
                    counts["groups"] += 1
                else:
                    counts["native_shapes"] += 1
                    # Geometry may be inherited from a master. Record this separately rather than inventing it.
                    geom = shape.findall("v:Section[@N='Geometry']", NS)
                    one_d = any(k in cells for k in ("BeginX", "EndX"))
                    if one_d:
                        counts["one_dimensional_shapes"] += 1
                    if geom or shape.get("Master") or shape.get("MasterShape"):
                        counts["non_background_geometry"] += 1
                    item["has_local_geometry"] = bool(geom)
                    item["master"] = shape.get("Master")
                    item["master_shape"] = shape.get("MasterShape")
                objects.append(item)
                children = shape.find("v:Shapes", NS)
                if children is not None:
                    walk(children, ancestor_groups + (identity["id"],))

        shapes = root.find("v:Shapes", NS)
        if shapes is not None:
            walk(shapes)
        ids = {i["id"] for i in objects}
        connections = [dict(c.attrib) for c in root.findall("v:Connects/v:Connect", NS)]
        attached = {}
        for conn in connections:
            from_id, to_id = conn.get("FromSheet"), conn.get("ToSheet")
            for field, value in (("FromSheet", from_id), ("ToSheet", to_id)):
                if value not in ids:
                    failures.append({"code": "missing_connector_shape", "part": part, "field": field, "shape_id": value})
            attached.setdefault(from_id, set()).add(conn.get("FromCell"))
        counts["connection_records"] = len(connections)
        counts["attached_connectors"] = sum({"BeginX", "EndX"}.issubset(cells) for cells in attached.values())
        pages.append({"part": part, "canvas": sizes.get(part, {}), "counts": dict(counts),
                      "objects": objects, "readable_text": texts, "connections": connections,
                      "full_page_picture_candidates": [],
                      "picture_coverage_status": "requires renderer inspection; not inferred from object count"})
    return {"pages": pages}


def audit(path, require_native=False):
    path = Path(path).resolve()
    suffix = path.suffix.lower()
    if suffix not in {".pptx", ".vsdx"}:
        raise ValueError("Input must be a .pptx or .vsdx file")
    result = {"file": str(path), "format": suffix[1:], "audit_scope": "package structure only",
              "requires_visual_review": True, "requires_source_comparison": True,
              "native_mtef_verified": False, "hard_failures": [], "warnings": [],
              "limitations": ["Counts cannot establish complete editability or fidelity.",
                              "Images may be valid atomic assets; their presence alone is not a failure.",
                              "Declared MathType ProgID is not proof of valid native MTEF or successful reopening.",
                              "AlternateContent uses its first Choice (otherwise Fallback) to avoid double counting; target renderer support is unverified.",
                              "Source labels, missing motifs, clipping, routing and typography require renderer comparison."]}
    with ZipFile(path) as archive:
        names_list = archive.namelist()
        names = set(names_list)
        if len(names_list) != len(names):
            result["hard_failures"].append({"code": "duplicate_zip_entries"})
        corrupt = archive.testzip()
        if corrupt:
            result["hard_failures"].append({"code": "crc_failure", "part": corrupt})
        for required in ("[Content_Types].xml", "_rels/.rels", "ppt/presentation.xml" if suffix == ".pptx" else "visio/document.xml"):
            if required not in names:
                result["hard_failures"].append({"code": "missing_required_part", "part": required})
        by_source, failures, external = relationships(archive, names)
        result["hard_failures"].extend(failures)
        result["external_relationships"] = external
        result["embedded_payload_parts"] = sorted(n for n in names if "/embeddings/" in n and not n.endswith("/"))
        # Parse all XML parts so corruption outside the visible page is not hidden by a narrow inventory.
        for name in sorted(n for n in names if n.endswith(".xml")):
            try:
                xml_part(archive, name)
            except ET.ParseError as exc:
                result["hard_failures"].append({"code": "invalid_xml", "part": name, "detail": str(exc)})
        if any(f["code"] in {"missing_required_part", "invalid_xml", "crc_failure"} for f in result["hard_failures"]):
            result["pages"] = []
        else:
            inventory = (pptx_inventory if suffix == ".pptx" else visio_inventory)(archive, names, by_source, result["hard_failures"])
            result.update(inventory)
        totals = Counter()
        for page in result["pages"]:
            totals.update(page["counts"])
            c = page["counts"]
            has_native = any(c.get(k, 0) for k in ("non_background_geometry", "connectors", "one_dimensional_shapes", "native_tables", "chart_references", "omml_equations"))
            page["has_native_reconstruction_evidence"] = has_native
            if not has_native:
                finding = {"code": "no_native_reconstruction_evidence", "part": page["part"]}
                result["warnings"].append(finding)
                if require_native:
                    result["hard_failures"].append(finding.copy())
            if page["full_page_picture_candidates"]:
                result["warnings"].append({"code": "full_page_picture_candidate", "part": page["part"],
                                           "objects": page["full_page_picture_candidates"],
                                           "action": "Inspect whether reference artwork was flattened; valid backgrounds require justification."})
        if not result["pages"]:
            result["hard_failures"].append({"code": "no_pages_found"})
        result["totals"] = dict(totals)
        result["status"] = "FAIL" if result["hard_failures"] else "STRUCTURE_PASS_REVIEW_REQUIRED"
    return result


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("file", type=Path)
    parser.add_argument("--output", type=Path)
    parser.add_argument("--require-native", action="store_true",
                        help="Fail a page with no native diagram geometry, connectors, table/chart, or OMML evidence.")
    args = parser.parse_args(argv)
    try:
        result = audit(args.file, args.require_native)
        code = int(bool(result["hard_failures"]))
    except (OSError, BadZipFile, ET.ParseError, ValueError, KeyError) as exc:
        result = {"file": str(args.file.resolve()), "status": "ERROR", "requires_visual_review": True,
                  "hard_failures": [{"code": "unreadable_package", "detail": str(exc)}]}
        code = 2
    serialized = json.dumps(result, ensure_ascii=False, indent=2)
    if args.output:
        args.output.parent.mkdir(parents=True, exist_ok=True)
        args.output.write_text(serialized + "\n", encoding="utf-8")
    print(serialized)
    return code


if __name__ == "__main__":
    sys.exit(main())
