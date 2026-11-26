"""Focused offline proof/validation regressions; no Office process is started.

Copyright (c) 2026 NCEPU-Senior. Author: XiaoN.
"""
from copy import deepcopy
import importlib.util
import json
import os
from pathlib import Path
import shutil
import subprocess
import tempfile
import unittest
from xml.etree import ElementTree as E
from zipfile import ZipFile


ROOT = Path(__file__).resolve().parents[1]
SPEC = importlib.util.spec_from_file_location("saved_proof", ROOT / "scripts/check_pptx_roundtrip.py")
proof = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(proof)
P, A = proof.NS["p"], proof.NS["a"]


def tag(namespace, name):
    return "{" + namespace + "}" + name


def transform(parent, x, y, w, h):
    node = E.SubElement(parent, tag(A, "xfrm"))
    E.SubElement(node, tag(A, "off"), x=str(round(x*12700)), y=str(round(y*12700)))
    E.SubElement(node, tag(A, "ext"), cx=str(round(w*12700)), cy=str(round(h*12700)))
    return node


def shape(parent, identity, name, bounds, text="", kind="sp", attachments=None):
    node = E.SubElement(parent, tag(P, kind))
    nv = E.SubElement(node, tag(P, "nvCxnSpPr" if kind == "cxnSp" else "nvSpPr"))
    E.SubElement(nv, tag(P, "cNvPr"), id=str(identity), name=name)
    if attachments:
        cn = E.SubElement(nv, tag(P, "cNvCxnSpPr"))
        for endpoint, target, site in attachments:
            E.SubElement(cn, tag(A, endpoint), id=str(target), idx=str(site))
    props = E.SubElement(node, tag(P, "spPr"))
    transform(props, *bounds)
    E.SubElement(props, tag(A, "prstGeom"), prst="line" if kind == "cxnSp" else "rect")
    if kind == "sp":
        body = E.SubElement(node, tag(P, "txBody"))
        paragraph = E.SubElement(body, tag(A, "p"))
        E.SubElement(E.SubElement(paragraph, tag(A, "r")), tag(A, "t")).text = text
    return node


def package(path, slide, width=380, height=240):
    presentation = E.Element(tag(P, "presentation"))
    identities = E.SubElement(presentation, tag(P, "sldIdLst"))
    E.SubElement(identities, tag(P, "sldId"), id="256", **{
        "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id": "rId1"})
    E.SubElement(presentation, tag(P, "sldSz"), cx=str(width*12700), cy=str(height*12700))
    with ZipFile(path, "w") as archive:
        archive.writestr("ppt/presentation.xml", E.tostring(presentation))
        archive.writestr("ppt/_rels/presentation.xml.rels", '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Target="slides/slide1.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide"/></Relationships>')
        archive.writestr("ppt/slides/slide1.xml", E.tostring(slide))


def scene_package(path):
    scene = json.loads((ROOT / "examples/pptx-fidelity-base.json").read_text())
    slide = E.Element(tag(P, "sld")); tree = E.SubElement(E.SubElement(slide, tag(P, "cSld")), tag(P, "spTree"))
    identities = {row["id"]: i+2 for i, row in enumerate(scene["objects"])}
    rows = {row["id"]: row for row in scene["objects"]}
    for row in scene["objects"]:
        if row["type"] == "connector":
            first, last = rows[row["from"]], rows[row["to"]]
            # Native endpoint geometry is not consumed by ValidateOnly.
            shape(tree, identities[row["id"]], row["id"], (104, 105, 116, 0), kind="cxnSp",
                  attachments=[("stCxn", identities[first["id"]], 3), ("endCxn", identities[last["id"]], 1)])
        else:
            shape(tree, identities[row["id"]], row["id"], [row[k] for k in ("x", "y", "w", "h")], row.get("text", ""))
    package(path, slide)


def grouped_package(path, moved=False):
    slide = E.Element(tag(P, "sld")); tree = E.SubElement(E.SubElement(slide, tag(P, "cSld")), tag(P, "spTree"))
    shape(tree, 2, "source", (10, 20, 30, 40), "source")
    shape(tree, 6, "sink", (200, 150, 40, 30), "sink")
    group = E.SubElement(tree, tag(P, "grpSp")); nv = E.SubElement(group, tag(P, "nvGrpSpPr"))
    E.SubElement(nv, tag(P, "cNvPr"), id="10", name="receiver_group")
    props = E.SubElement(group, tag(P, "grpSpPr")); xf = transform(props, 112 if moved else 100, 80, 90, 60)
    E.SubElement(xf, tag(A, "chOff"), x="0", y="0"); E.SubElement(xf, tag(A, "chExt"), cx=str(90*12700), cy=str(60*12700))
    shape(group, 3, "receiver", (20, 20, 40, 20), "edited" if moved else "receiver")
    head = shape(group, 7, "custom_head", (10, 4, 20, 12)); custom = E.SubElement(head.find("p:spPr", proof.NS), tag(A, "custGeom"))
    E.SubElement(custom, tag(A, "cubicBezTo"))
    shape(group, 8, "rear_port", (10, 10, .25, .25))
    shape(tree, 4, "incoming", (40, 40, 82 if moved else 70, 50), kind="cxnSp",
          attachments=[("stCxn", 2, 3), ("endCxn", 8, 1)])
    shape(tree, 5, "outgoing", (172 if moved else 160, 110, 28 if moved else 40, 40), kind="cxnSp",
          attachments=[("stCxn", 3, 3), ("endCxn", 6, 0)])
    package(path, slide)


class SavedGroupProofTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        root = Path(self.temp.name)
        grouped_package(root / "before.pptx"); grouped_package(root / "after.pptx", True)
        self.before, self.after = proof.inspect(root / "before.pptx"), proof.inspect(root / "after.pptx")

    def compare(self, after=None, **kwargs):
        return proof.compare(self.before, after or self.after, moved_names=["receiver_group"], **kwargs)

    def test_group_and_all_descendants_follow_while_other_end_stays(self):
        result = self.compare(expected_text={"receiver": "edited"}, expected_delta=[12, 0])
        self.assertEqual(result["errors"], [])
        endpoint_checks = [check for row in result["connector_checks"] for check in row["endpoint_checks"]]
        self.assertEqual(sum(check["target_selected"] for check in endpoint_checks), 2)
        self.assertTrue(all(check["follows_target"] for check in endpoint_checks))

    def test_arrow_that_does_not_move_with_group_is_rejected(self):
        edited = deepcopy(self.after)
        head = next(row for row in edited["slides"][0]["objects"] if row["name"] == "custom_head")
        head["corners_pt"] = [[x-12, y] for x, y in head["corners_pt"]]
        self.assertTrue(any("Group member" in error for error in self.compare(edited)["errors"]))

    def test_unselected_node_and_endpoint_cannot_drift_together(self):
        edited = deepcopy(self.after)
        rows = {row["name"]: row for row in edited["slides"][0]["objects"]}
        rows["source"]["corners_pt"] = [[x+2, y] for x, y in rows["source"]["corners_pt"]]
        rows["source"]["centre_pt"][0] += 2; rows["incoming"]["endpoints_pt"]["begin"][0] += 2
        errors = self.compare(edited)["errors"]
        self.assertTrue(any("Unselected object moved" in error for error in errors))
        self.assertTrue(any("Endpoint on unselected target" in error for error in errors))

    def test_changed_port_site_and_expected_text_are_rejected(self):
        edited = deepcopy(self.after)
        line = next(row for row in edited["slides"][0]["objects"] if row["name"] == "incoming")
        line["attachments"]["end"]["xml_site_index"] = 0
        self.assertTrue(self.compare(edited)["errors"])
        self.assertTrue(self.compare(expected_text={"receiver": "wrong"})["errors"])

    def test_unsupported_attached_geometry_does_not_pass(self):
        edited = deepcopy(self.after)
        line = next(row for row in edited["slides"][0]["objects"] if row["name"] == "incoming")
        line["endpoints_pt"] = None
        self.assertTrue(any("explicit path interpretation" in error for error in self.compare(edited)["errors"]))

    def test_affine_scale_flip_rotation_for_nested_coordinates(self):
        t = {"left": 100, "top": 80, "width": 90, "height": 60,
             "rotation_deg": 90, "flip_h": True, "flip_v": False}
        self.assertEqual(proof.apply(proof.shape_matrix(t), (0, 0)), [175, 155])
        self.assertEqual(proof.apply(proof.multiply((2, 0, 0, 3, 10, 20), (1, 0, 0, 1, 4, 5)), (2, 3)), [22, 44])


@unittest.skipUnless(os.name == "nt" and shutil.which("powershell.exe"), "Windows PowerShell validation")
class RefinementValidationTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory(); self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name); self.source = self.root / "source.pptx"; scene_package(self.source)
        self.spec = json.loads((ROOT / "examples/pptx-fidelity.json").read_text())

    def validate(self, spec=None):
        config = self.root / "refinement.json"; config.write_text(json.dumps(spec or self.spec), encoding="utf-8")
        return subprocess.run(["powershell.exe", "-NoProfile", "-NonInteractive", "-File", str(ROOT / "scripts/refine_pptx.ps1"),
                               "-SourceDeck", str(self.source), "-SpecPath", str(config), "-OutputDir", str(self.root / "output"), "-ValidateOnly"],
                              capture_output=True, text=True, timeout=30)

    def test_paired_example_validates_without_output_or_office(self):
        result = self.validate(); self.assertEqual(result.returncode, 0, result.stdout+result.stderr)
        self.assertIn("schema_valid", result.stdout); self.assertFalse((self.root / "output").exists())

    def test_unclosed_excessive_or_nonfinite_paths_are_rejected(self):
        variants = []
        spec = deepcopy(self.spec); spec["arrowheads"][0]["path"]["segments"][-1]["to"] = [-19, -7]; variants.append(spec)
        spec = deepcopy(self.spec); spec["max_head_extent_pt"] = 10; variants.append(spec)
        spec = deepcopy(self.spec); spec["arrowheads"][0]["tip"] = [float("inf"), 0]; variants.append(spec)
        for spec in variants:
            with self.subTest(spec=spec): self.assertNotEqual(self.validate(spec).returncode, 0)

    def test_wrong_receiver_duplicate_members_and_unmeasured_formula_rejected(self):
        variants = []
        spec = deepcopy(self.spec); spec["receiver_groups"][0]["members"] = ["source"]; variants.append(spec)
        spec = deepcopy(self.spec); spec["receiver_groups"][0]["members"].append("receiver"); variants.append(spec)
        spec = deepcopy(self.spec); spec["formula_placements"][0]["evidence"] = ""; variants.append(spec)
        spec = deepcopy(self.spec); spec["formula_placements"].append(deepcopy(spec["formula_placements"][0])); variants.append(spec)
        spec = deepcopy(self.spec); del spec["formula_placements"][0]["observed_baseline_pt"]; variants.append(spec)
        for spec in variants:
            with self.subTest(spec=spec): self.assertNotEqual(self.validate(spec).returncode, 0)

    def test_output_collision_is_rejected_without_overwriting(self):
        folder = self.root / "output"; folder.mkdir(); target = folder / "pptx_fidelity.pptx"; target.write_bytes(b"keep me")
        result = self.validate(); self.assertNotEqual(result.returncode, 0); self.assertEqual(target.read_bytes(), b"keep me")

    def test_cardinal_path_rotations_and_horizontal_only_placement(self):
        script = str(ROOT / "scripts/refine_pptx.ps1").replace("'", "''")
        body = r"""
$tokens=$null;$errors=$null
$ast=[Management.Automation.Language.Parser]::ParseFile('__SCRIPT__',[ref]$tokens,[ref]$errors)
if($errors.Count){throw 'Parse failure'}
$ast.FindAll({param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -in @('Finite','Rotate-Point','Baseline-Shift')},$true) | ForEach-Object { Invoke-Expression $_.Extent.Text }
$rotations=@(foreach($direction in @('right','down','left','up')){,@(Rotate-Point @(-20,7) @(100,100) $direction)})
if((Baseline-Shift ([pscustomobject]@{delta_x_pt=-4})) -ne 0){throw 'Horizontal placement shifted baseline'}
if((Baseline-Shift ([pscustomobject]@{observed_baseline_pt=126;target_baseline_pt=128})) -ne 2){throw 'Incorrect baseline offset'}
$rotations | ConvertTo-Json -Compress
""".replace("__SCRIPT__", script)
        result = subprocess.run(["powershell.exe", "-NoProfile", "-NonInteractive", "-Command", body], capture_output=True, text=True, timeout=30)
        self.assertEqual(result.returncode, 0, result.stdout+result.stderr)
        rotations = json.loads(result.stdout)
        # Windows PowerShell's pipeline can retain a typed-array PSObject shell.
        rotations = [row["value"] if isinstance(row, dict) else row for row in rotations]
        self.assertEqual(rotations, [[80, 107], [93, 80], [120, 93], [107, 120]])


if __name__ == "__main__":
    unittest.main()
