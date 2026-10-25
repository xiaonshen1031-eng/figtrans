"""Behavioral package fixtures for the editability auditor.

Copyright (c) 2026 NCEPU-Senior
Author: XiaoN
"""
import importlib.util
from pathlib import Path
import tempfile
import unittest
from zipfile import ZipFile

SCRIPT = Path(__file__).resolve().parents[1] / "scripts" / "audit_editability.py"
SPEC = importlib.util.spec_from_file_location("audit_editability", SCRIPT)
MOD = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(MOD)

P = MOD.NS["p"]
A = MOD.NS["a"]
R = MOD.NS["r"]
V = MOD.NS["v"]
M = MOD.NS["m"]
REL = MOD.NS["rel"]
TYPE = "http://schemas.openxmlformats.org/package/2006/content-types"


def rels(*items):
    return '<Relationships xmlns="' + REL + '">' + "".join(
        '<Relationship Id="%s" Type="%s" Target="%s"/>' % item for item in items) + '</Relationships>'


def package(path, members):
    with ZipFile(path, "w") as z:
        for name, value in members.items():
            z.writestr(name, value)


def ppt(path, objects, extras=None, slide_rels=None):
    members = {
        "[Content_Types].xml": '<Types xmlns="' + TYPE + '"/>',
        "_rels/.rels": rels(("rDoc", "officeDocument", "ppt/presentation.xml")),
        "ppt/presentation.xml": '<p:presentation xmlns:p="' + P + '"><p:sldSz cx="7200000" cy="4800000"/></p:presentation>',
        "ppt/slides/slide1.xml": '<p:sld xmlns:p="' + P + '" xmlns:a="' + A + '" xmlns:r="' + R + '" xmlns:m="' + M + '"><p:cSld><p:spTree>' + objects + '</p:spTree></p:cSld></p:sld>',
    }
    if slide_rels is not None:
        members["ppt/slides/_rels/slide1.xml.rels"] = slide_rels
    members.update(extras or {})
    package(path, members)


def rectangle(sid, name, text=""):
    return ('<p:sp><p:nvSpPr><p:cNvPr id="%s" name="%s"/><p:cNvSpPr/></p:nvSpPr>'
            '<p:spPr><a:xfrm><a:off x="100000" y="100000"/><a:ext cx="1000000" cy="600000"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>'
            '<p:txBody><a:p><a:r><a:t>%s</a:t></a:r></a:p></p:txBody></p:sp>') % (sid, name, text)


class AuditTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.folder = Path(self.temp.name)

    def tearDown(self):
        self.temp.cleanup()

    def test_flat_picture_is_not_native_reconstruction(self):
        path = self.folder / "flat.pptx"
        pic = '<p:pic><p:nvPicPr><p:cNvPr id="2" name="flattened-reference"/></p:nvPicPr><p:blipFill><a:blip r:embed="rImg"/></p:blipFill><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="7200000" cy="4800000"/></a:xfrm></p:spPr></p:pic>'
        ppt(path, pic, {"ppt/media/image1.png": b"test picture"}, rels(("rImg", "image", "../media/image1.png")))
        result = MOD.audit(path, require_native=True)
        self.assertEqual(result["status"], "FAIL")
        self.assertEqual(result["totals"]["pictures"], 1)
        self.assertEqual(result["pages"][0]["full_page_picture_candidates"][0]["page_area_coverage"], 1)
        self.assertIn("no_native_reconstruction_evidence", {f["code"] for f in result["hard_failures"]})

    def test_native_ppt_connector_keeps_attachment_ids(self):
        path = self.folder / "native.pptx"
        connector = '<p:cxnSp><p:nvCxnSpPr><p:cNvPr id="4" name="data-flow"/><p:cNvCxnSpPr><a:stCxn id="2" idx="3"/><a:endCxn id="3" idx="1"/></p:cNvCxnSpPr></p:nvCxnSpPr></p:cxnSp>'
        ppt(path, rectangle("2", "input", "Input") + rectangle("3", "output", "Output") + connector)
        result = MOD.audit(path, require_native=True)
        self.assertFalse(result["hard_failures"])
        self.assertEqual(result["totals"]["attached_connectors"], 1)
        self.assertEqual(result["pages"][0]["connections"][0]["attachments"][1]["shape_id"], "3")
        self.assertEqual([t["text"] for t in result["pages"][0]["readable_text"]], ["Input", "Output"])
        self.assertTrue(result["requires_visual_review"])

    def test_missing_relationship_target_is_failure(self):
        path = self.folder / "broken.pptx"
        ppt(path, rectangle("2", "box"), slide_rels=rels(("rMissing", "image", "../media/absent.png")))
        result = MOD.audit(path)
        failure = next(f for f in result["hard_failures"] if f["code"] == "missing_relationship_target")
        self.assertEqual(failure["resolved_part"], "ppt/media/absent.png")
        self.assertEqual(result["status"], "FAIL")

    def test_missing_relationship_id_is_failure(self):
        path = self.folder / "missing-id.pptx"
        ppt(path, '<p:pic><p:nvPicPr><p:cNvPr id="2" name="icon"/></p:nvPicPr><p:blipFill><a:blip r:embed="rAbsent"/></p:blipFill></p:pic>')
        result = MOD.audit(path)
        self.assertIn("missing_relationship_id", {f["code"] for f in result["hard_failures"]})

    def test_visio_connects_and_text_inventory(self):
        path = self.folder / "native.vsdx"
        xml = '<PageContents xmlns="' + V + '"><Shapes>'
        xml += '<Shape ID="1" NameU="input" Type="Shape"><Cell N="Width" V="2"/><Section N="Geometry" IX="0"><Row T="LineTo" IX="1"/></Section><Text>Input</Text></Shape>'
        xml += '<Shape ID="2" NameU="output" Type="Shape"><Section N="Geometry" IX="0"/><Text>Output</Text></Shape>'
        xml += '<Shape ID="3" NameU="flow" Type="Shape"><Cell N="BeginX" V="1"/><Cell N="EndX" V="3"/><Section N="Geometry" IX="0"/></Shape>'
        xml += '</Shapes><Connects><Connect FromSheet="3" FromCell="BeginX" ToSheet="1" ToCell="PinX"/><Connect FromSheet="3" FromCell="EndX" ToSheet="2" ToCell="PinX"/></Connects></PageContents>'
        package(path, {"[Content_Types].xml": '<Types xmlns="' + TYPE + '"/>',
                       "_rels/.rels": rels(("rDoc", "officeDocument", "visio/document.xml")),
                       "visio/document.xml": '<VisioDocument xmlns="' + V + '"/>',
                       "visio/pages/page1.xml": xml})
        result = MOD.audit(path, require_native=True)
        self.assertFalse(result["hard_failures"])
        self.assertEqual(result["totals"]["attached_connectors"], 1)
        self.assertEqual(result["totals"]["native_shapes"], 3)
        self.assertEqual(result["pages"][0]["readable_text"][0]["text"], "Input")

    def test_generic_ole_not_claimed_as_mathtype(self):
        path = self.folder / "ole.pptx"
        frame = '<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="2" name="spreadsheet"/></p:nvGraphicFramePr><a:graphic><a:graphicData><p:oleObj progId="Excel.Sheet.12" r:id="rOle"/></a:graphicData></a:graphic></p:graphicFrame>'
        ppt(path, frame, {"ppt/embeddings/data.bin": b"not an equation"}, rels(("rOle", "oleObject", "../embeddings/data.bin")))
        result = MOD.audit(path)
        self.assertEqual(result["totals"]["generic_ole"], 1)
        self.assertEqual(result["totals"]["mathtype_declared_ole"], 0)
        self.assertFalse(result["native_mtef_verified"])

    def test_mathtype_progid_is_declaration_not_proof(self):
        path = self.folder / "equation.pptx"
        frame = '<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="2" name="formula"/></p:nvGraphicFramePr><a:graphic><a:graphicData><p:oleObj progId="Equation.DSMT4" r:id="rOle"/></a:graphicData></a:graphic></p:graphicFrame>'
        ppt(path, frame, {"ppt/embeddings/equation.bin": b"invalid payload despite ProgID"}, rels(("rOle", "oleObject", "../embeddings/equation.bin")))
        result = MOD.audit(path)
        self.assertEqual(result["totals"]["mathtype_declared_ole"], 1)
        self.assertFalse(result["native_mtef_verified"])
        self.assertFalse(result["pages"][0]["objects"][0]["ole_objects"][0]["native_mtef_verified"])

    def test_omml_is_recorded_separately(self):
        path = self.folder / "omml.pptx"
        shape = rectangle("2", "math").replace('</p:txBody>', '<m:oMath><m:r><m:t>x</m:t></m:r></m:oMath></p:txBody>')
        ppt(path, shape)
        result = MOD.audit(path)
        self.assertEqual(result["totals"]["omml_equations"], 1)
        self.assertEqual(result["totals"].get("mathtype_declared_ole", 0), 0)

    def test_alternate_ole_fallback_is_not_double_counted(self):
        path = self.folder / "fallback.pptx"
        ole = '<p:oleObj progId="Equation.DSMT4" r:id="rOle"/>'
        alternate = '<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice Requires="p">' + ole + '</mc:Choice><mc:Fallback>' + ole + '</mc:Fallback></mc:AlternateContent>'
        frame = '<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="2" name="formula"/></p:nvGraphicFramePr><a:graphic><a:graphicData>' + alternate + '</a:graphicData></a:graphic></p:graphicFrame>'
        ppt(path, frame, {"ppt/embeddings/equation.bin": b"test"}, rels(("rOle", "oleObject", "../embeddings/equation.bin")))
        result = MOD.audit(path)
        self.assertEqual(result["totals"]["ole_objects"], 1)
        self.assertEqual(result["totals"]["mathtype_declared_ole"], 1)
        self.assertEqual(result["pages"][0]["alternate_content_selections"], 1)

    def test_invalid_xml_part_is_hard_failure(self):
        path = self.folder / "malformed.pptx"
        ppt(path, rectangle("2", "box"), {"docProps/core.xml": "<broken>"})
        result = MOD.audit(path)
        self.assertIn("invalid_xml", {f["code"] for f in result["hard_failures"]})

    def test_atomic_picture_does_not_fail_native_diagram(self):
        path = self.folder / "mixed.pptx"
        pic = '<p:pic><p:nvPicPr><p:cNvPr id="3" name="atomic-icon"/></p:nvPicPr><p:blipFill><a:blip r:embed="rImg"/></p:blipFill><p:spPr><a:xfrm><a:off x="100000" y="100000"/><a:ext cx="300000" cy="300000"/></a:xfrm></p:spPr></p:pic>'
        ppt(path, rectangle("2", "module", "Module") + pic, {"ppt/media/icon.png": b"test"}, rels(("rImg", "image", "../media/icon.png")))
        result = MOD.audit(path, require_native=True)
        self.assertFalse(result["hard_failures"])
        self.assertFalse(result["pages"][0]["full_page_picture_candidates"])


if __name__ == "__main__":
    unittest.main()
