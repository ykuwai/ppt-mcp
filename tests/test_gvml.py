"""Tests for ``gvml``, the pure half of the clipboard route to PowerPoint.

Nothing here needs PowerPoint, appscript or macOS. ``gvml`` imports only the
standard library, so this file runs on the Windows CI as well, and it is the
only place the XML that gets pasted is checked by a machine: PowerPoint says
nothing about a package it cannot use, it just leaves the slide as it was.

The fixtures under ``tests/fixtures/gvml`` are packages PowerPoint for Mac
itself put on the pasteboard, recorded with ``scripts/gvml_record.py``. The
rectangle was made by PowerPoint with ``make new shape``; the group, the
freeform and the chart were pasted from packages this module built and copied
back, so their bytes are PowerPoint's re-serialisation of what it accepted.
The golden tests ask that what we write is an ordered subtree of what
PowerPoint writes: same tags in the same order, same attribute values where
we set one. That is what catches a template whose elements have drifted out
of schema order, which is the failure PowerPoint hides best. On the machine
these were recorded on, a ``graphicFrame`` with ``xfrm`` before ``graphic``
was dropped by ``paste object`` without an error; the test named after it
holds that finding.
"""

import io
import os
import sys
import xml.etree.ElementTree as ET
import zipfile

import pytest

sys.path.insert(0, "src")

from gvml import Package, PackageError, build, read, validate  # noqa: E402
from gvml import canvas, charts, freeform, shapes  # noqa: E402
from gvml.package import (  # noqa: E402
    CHART_PART,
    CONTENT_TYPES_PART,
    CT_CHART,
    DRAWING_PART,
    NS_A,
    REL_CHART,
    Graft,
    Relationship,
)

FIXTURES = os.path.join(os.path.dirname(__file__), "fixtures", "gvml")
E = canvas.EMU_PER_PT
A = f"{{{NS_A}}}"


def fixture(name) -> Package:
    with open(os.path.join(FIXTURES, f"{name}.gvml.zip"), "rb") as fh:
        return read(fh.read())


def fixture_canvas(name):
    return canvas.parse(fixture(name).drawing())


# Attributes whose values legitimately differ between our package and
# PowerPoint's copy of it: ids are reassigned, the locale is the machine's,
# and a pasted shape's offset is wherever PowerPoint dropped it.
_IGNORED = {("cNvPr", "id"), ("*", "lang"), ("*", "altLang"),
            ("off", "x"), ("off", "y"), ("chOff", "x"), ("chOff", "y")}


def _local(tag):
    return tag.rsplit("}", 1)[-1]


def ordered_subtree(mine, theirs, path=""):
    """Return None when ``mine`` is an ordered subtree of ``theirs``, else why."""
    here = f"{path}/{_local(mine.tag)}"
    if mine.tag != theirs.tag:
        return f"{here}: tag differs, theirs is {_local(theirs.tag)}"
    for key, value in mine.attrib.items():
        local_key = _local(key)
        if (_local(mine.tag), local_key) in _IGNORED or ("*", local_key) in _IGNORED:
            continue
        if theirs.get(key) != value:
            return f"{here}@{local_key}: ours {value!r}, theirs {theirs.get(key)!r}"
    position = 0
    their_children = list(theirs)
    for child in mine:
        while position < len(their_children) and their_children[position].tag != child.tag:
            position += 1
        if position >= len(their_children):
            return f"{here}: {_local(child.tag)} is not where PowerPoint puts it, or is absent"
        problem = ordered_subtree(child, their_children[position], here)
        if problem:
            return problem
        position += 1
    return None


def rect_spec():
    return shapes.shape_xml(2, "HandRect", 100 * E, 100 * E, 200 * E, 120 * E)


def group_spec():
    kids = [
        shapes.shape_xml(3, "KidA", 100 * E, 100 * E, 80 * E, 60 * E),
        shapes.shape_xml(4, "KidB", 200 * E, 150 * E, 80 * E, 60 * E),
    ]
    return shapes.group_xml(2, "HandGroup", 100 * E, 100 * E, 180 * E, 110 * E, kids)


FREEFORM_NODES = [
    {"seg_int": 0, "et_int": 0, "x1": 200, "y1": 100},
    {"seg_int": 1, "et_int": 0, "x1": 250, "y1": 200},
    {"seg_int": 1, "et_int": 1, "x1": 220, "y1": 260, "x2": 160, "y2": 260, "x3": 120, "y3": 200},
]


def freeform_spec():
    geometry, x, y, cx, cy = freeform.build_geometry(100, 150, FREEFORM_NODES, True)
    return shapes.shape_xml(2, "HandFree", x, y, cx, cy, geometry), (x, y, cx, cy)


# ---------------------------------------------------------------------------
# Golden fixtures
# ---------------------------------------------------------------------------
class TestWhatWeWriteIsWhatPowerPointWrites:
    """Every template, as an ordered subtree of PowerPoint's own output."""

    def test_the_rectangle_made_by_powerpoint_itself(self):
        ours = ET.fromstring(canvas.wrap(rect_spec(), 100 * E, 100 * E, 200 * E, 120 * E))
        theirs = ET.parse(io.BytesIO(fixture("rect").drawing())).getroot()
        assert ordered_subtree(ours, theirs) is None

    def test_the_group_and_its_two_members(self):
        ours = ET.fromstring(canvas.wrap(group_spec(), 100 * E, 100 * E, 180 * E, 110 * E))
        theirs = ET.parse(io.BytesIO(fixture("group").drawing())).getroot()
        assert ordered_subtree(ours, theirs) is None

    def test_the_freeform_and_every_point_in_its_path(self):
        body, (x, y, cx, cy) = freeform_spec()
        ours = ET.fromstring(canvas.wrap(body, x, y, cx, cy))
        theirs = ET.parse(io.BytesIO(fixture("freeform").drawing())).getroot()
        assert ordered_subtree(ours, theirs) is None
        # The path is relative to the shape, so the points survive the move.
        theirs_pts = [(p.get("x"), p.get("y")) for p in theirs.iter(f"{A}pt")]
        ours_pts = [(p.get("x"), p.get("y")) for p in ours.iter(f"{A}pt")]
        assert ours_pts == theirs_pts

    def test_the_chart_frame_in_the_order_that_pastes(self):
        """nvGraphicFramePr, graphic, xfrm. The other order is silently dropped."""
        raw = charts.chart_package("HandChart", 50 * E, 50 * E, 500 * E, 350 * E, charts.chart_xml(51))
        ours = ET.parse(io.BytesIO(Package.from_bytes(raw).drawing())).getroot()
        theirs = ET.parse(io.BytesIO(fixture("chart").drawing())).getroot()
        assert ordered_subtree(ours, theirs) is None

    def test_a_frame_with_xfrm_before_graphic_would_be_caught_here(self):
        """The H2 finding: measured on this machine, that order pastes nothing."""
        frame = shapes.chart_frame_xml(2, "HandChart", 50 * E, 50 * E, 500 * E, 350 * E)
        xfrm_start = frame.index("<a:xfrm>")
        xfrm = frame[xfrm_start:frame.index("</a:xfrm>") + len("</a:xfrm>")]
        swapped = frame.replace(xfrm, "").replace("<a:graphic>", xfrm + "<a:graphic>")
        assert swapped != frame
        ours = ET.fromstring(canvas.wrap(swapped, 50 * E, 50 * E, 500 * E, 350 * E))
        theirs = ET.parse(io.BytesIO(fixture("chart").drawing())).getroot()
        problem = ordered_subtree(ours, theirs)
        assert problem is not None
        assert "xfrm" in problem or "graphic" in problem

    def test_the_chart_xml_itself(self):
        ours = ET.fromstring(charts.chart_xml(51))
        theirs = ET.parse(io.BytesIO(fixture("chart").chart())).getroot()
        assert ordered_subtree(ours, theirs) is None

    @pytest.mark.parametrize("type_int", charts.supported_types())
    def test_every_chart_kind_keeps_the_chart_level_order(self, type_int):
        """The kinds without a fixture still share c:chart's element order."""
        ours = ET.fromstring(charts.chart_xml(type_int))
        theirs = ET.parse(io.BytesIO(fixture("chart").chart())).getroot()
        c = "{http://schemas.openxmlformats.org/drawingml/2006/chart}"
        our_chart = ours.find(f"{c}chart")
        their_chart = theirs.find(f"{c}chart")
        our_order = [_local(e.tag) for e in our_chart if _local(e.tag) != "view3D"]
        their_order = [_local(e.tag) for e in their_chart]
        position = 0
        for tag in our_order:
            while position < len(their_order) and their_order[position] != tag:
                position += 1
            assert position < len(their_order), f"{tag} out of order for {type_int}"
            position += 1


# ---------------------------------------------------------------------------
# Round trips
# ---------------------------------------------------------------------------
class TestRoundTrips:
    def test_a_group_reads_back_its_members_in_points(self):
        group = canvas.children(canvas.parse(
            canvas.wrap(group_spec(), 100 * E, 100 * E, 180 * E, 110 * E).encode()
        ))[0]
        assert shapes.group_items(group) == [
            {"name": "KidA", "type": 1, "type_name": "AutoShape",
             "left": 100.0, "top": 100.0, "width": 80.0, "height": 60.0},
            {"name": "KidB", "type": 1, "type_name": "AutoShape",
             "left": 200.0, "top": 150.0, "width": 80.0, "height": 60.0},
        ]

    def test_a_resized_group_scales_its_members(self):
        """PowerPoint keeps chOff/chExt and changes off/ext when a group grows."""
        group = canvas.children(canvas.parse(
            canvas.wrap(group_spec(), 100 * E, 100 * E, 180 * E, 110 * E).encode()
        ))[0]
        ext = group.find(f"{A}grpSpPr/{A}xfrm/{A}ext")
        ext.set("cx", str(360 * E))
        items = shapes.group_items(group)
        assert items[1]["left"] == 300.0
        assert items[1]["width"] == 160.0
        assert items[1]["top"] == 150.0

    def test_a_freeform_reads_back_with_windows_numbering(self):
        body, (x, y, cx, cy) = freeform_spec()
        sp = canvas.children(canvas.parse(canvas.wrap(body, x, y, cx, cy).encode()))[0]
        nodes = freeform.read_nodes(sp)
        assert [n["segment_type"] for n in nodes] == [
            "line", "curve", "inaccessible", "inaccessible", "curve",
            "inaccessible", "inaccessible", "line", "inaccessible",
        ]
        assert [(n["x"], n["y"]) for n in nodes if n["segment_type"] != "inaccessible"] == [
            (100.0, 150.0), (200.0, 100.0), (250.0, 200.0), (120.0, 200.0),
        ]
        # The corner curve's control points are exactly where they were asked.
        assert (nodes[5]["x"], nodes[5]["y"]) == (220.0, 260.0)
        assert (nodes[6]["x"], nodes[6]["y"]) == (160.0, 260.0)
        # The closing node repeats the start, as Windows adds one too.
        assert (nodes[8]["x"], nodes[8]["y"]) == (100.0, 150.0)
        assert nodes[8]["note"] == freeform.INACCESSIBLE_NOTE

    @pytest.mark.parametrize("type_int", charts.supported_types())
    def test_a_chart_reads_back_the_data_it_was_given(self, type_int):
        categories, series = charts.default_data(type_int)
        data = charts.read_chart(charts.chart_xml(type_int, categories, series).encode())
        assert [s["name"] for s in data["series"]] == [s["name"] for s in series]
        assert [s["values"] for s in data["series"]] == [s["values"] for s in series]
        assert data["categories"] == [str(c) for c in categories]

    def test_a_chart_with_custom_data(self):
        xml = charts.chart_xml(57, ["東京", "大阪"], [{"name": "売上", "values": [1.5, 2]}])
        data = charts.read_chart(xml.encode())
        assert data == {"categories": ["東京", "大阪"],
                        "series": [{"name": "売上", "values": [1.5, 2.0]}]}


# ---------------------------------------------------------------------------
# Package integrity
# ---------------------------------------------------------------------------
class TestValidateRefusesWhatPowerPointWouldSilentlyDrop:
    @pytest.mark.parametrize("name", ["rect", "group", "freeform", "chart"])
    def test_every_fixture_passes(self, name):
        validate(fixture(name).to_bytes())

    def test_bytes_that_are_not_a_zip(self):
        with pytest.raises(PackageError, match="not a zip"):
            validate(b"<a:graphic/>")

    def test_a_drawing_that_is_not_well_formed(self):
        with pytest.raises(PackageError, match="not well formed"):
            validate(build("<a:graphic><a:graphicData>"))

    def test_an_empty_canvas(self):
        with pytest.raises(PackageError, match="no shape at all"):
            validate(build(canvas.wrap("", 0, 0, 10, 10)))

    def test_a_relationship_pointing_nowhere(self):
        raw = build(
            canvas.wrap(shapes.chart_frame_xml(2, "C", 0, 0, 10, 10), 0, 0, 10, 10),
            drawing_rels=[Relationship("rId1", REL_CHART, CHART_PART)],
        )
        with pytest.raises(PackageError, match="not in the package"):
            validate(raw)

    def test_a_part_without_a_content_type(self):
        package = Package.from_bytes(build(canvas.wrap(rect_spec(), 0, 0, 10, 10)))
        package.parts["clipboard/media/image1.png"] = b"\x89PNG"
        with pytest.raises(PackageError, match="no content type"):
            package.validate()

    def test_a_missing_content_types_part(self):
        package = Package.from_bytes(build(canvas.wrap(rect_spec(), 0, 0, 10, 10)))
        del package.parts[CONTENT_TYPES_PART]
        with pytest.raises(PackageError, match="missing"):
            package.validate()

    def test_a_missing_root_relationship(self):
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w") as zf:
            zf.writestr(CONTENT_TYPES_PART, Package.from_bytes(
                build(canvas.wrap(rect_spec(), 0, 0, 10, 10))).parts[CONTENT_TYPES_PART])
            zf.writestr(DRAWING_PART, canvas.wrap(rect_spec(), 0, 0, 10, 10))
        with pytest.raises(PackageError, match="no drawing relationship"):
            validate(buf.getvalue())

    def test_the_drawing_is_reached_through_the_root_relationship(self):
        package = fixture("chart")
        assert package.drawing_part() == DRAWING_PART
        assert package.chart_part() == CHART_PART
        assert package.content_type(CHART_PART) == CT_CHART
        assert fixture("group").chart_part() is None


# ---------------------------------------------------------------------------
# Reading PowerPoint's own output
# ---------------------------------------------------------------------------
class TestReadingWhatPowerPointWrote:
    def test_the_chart_fixture_reads_as_windows_reports_it(self):
        data = charts.read_chart(fixture("chart").chart())
        assert data["categories"] == ["Category 1", "Category 2", "Category 3", "Category 4"]
        assert data["series"][0] == {"name": "Series 1", "values": [4.3, 2.5, 3.5, 4.5]}
        assert data["series"][2]["values"] == [2.0, 2.0, 3.0, 5.0]
        assert charts.kind_of(fixture("chart").chart()) == "barChart"

    def test_a_gap_in_a_cache_reads_as_none(self):
        xml = charts.chart_xml(51, ["a", "b", "c"], [{"name": "s", "values": [1, None, 3]}])
        assert charts.read_chart(xml.encode())["series"][0]["values"] == [1.0, None, 3.0]

    def test_a_scatter_keeps_its_x_values_as_the_categories(self):
        xml = charts.chart_xml(-4169, [0.5, 1.5], [{"name": "y", "values": [3, 4]}])
        assert charts.read_chart(xml.encode())["categories"] == ["0.5", "1.5"]

    def test_the_group_fixture_lists_both_members(self):
        group = canvas.children(fixture_canvas("group"))[0]
        assert [i["name"] for i in shapes.group_items(group)] == ["KidA", "KidB"]
        assert [i["width"] for i in shapes.group_items(group)] == [80.0, 80.0]

    def test_the_freeform_fixture_reads_nine_nodes_in_the_recorded_shape(self):
        """PowerPoint moved the shape when it pasted; the path is the same."""
        sp = canvas.children(fixture_canvas("freeform"))[0]
        nodes = freeform.read_nodes(sp)
        assert len(nodes) == 9
        origin = (nodes[0]["x"], nodes[0]["y"])
        relative = [(round(n["x"] - origin[0], 2), round(n["y"] - origin[1], 2)) for n in nodes]
        assert relative[1] == (100.0, -50.0)
        assert relative[4] == (150.0, 50.0)
        assert relative[7] == (20.0, 50.0)
        assert relative[8] == (0.0, 0.0)
        assert shapes.type_of(sp) == shapes.MSO_FREEFORM

    def test_a_freeform_resized_after_drawing_is_scaled_not_refused(self):
        """Measured: doubling the width doubles ext cx and leaves path w alone."""
        sp = canvas.children(fixture_canvas("freeform"))[0]
        before = freeform.read_nodes(sp)
        ext = sp.find(f"{A}spPr/{A}xfrm/{A}ext")
        ext.set("cx", str(int(ext.get("cx")) * 2))
        after = freeform.read_nodes(sp)
        assert after[1]["x"] - after[0]["x"] == pytest.approx(2 * (before[1]["x"] - before[0]["x"]), abs=0.02)
        assert after[1]["y"] == before[1]["y"]

    def test_the_rect_fixture_is_an_autoshape_and_the_chart_a_chart(self):
        assert shapes.describe(canvas.children(fixture_canvas("rect"))[0]).type == shapes.MSO_AUTO_SHAPE
        assert shapes.describe(canvas.children(fixture_canvas("chart"))[0]).type == shapes.MSO_CHART
        assert shapes.describe(canvas.children(fixture_canvas("group"))[0]).type == shapes.MSO_GROUP
        assert shapes.name_of(canvas.children(fixture_canvas("chart"))[0]) == "HandChart"


class TestEditingTypeIsReadOffTheHandles:
    """OOXML stores no editing type; the handles say what PowerPoint would show."""

    @staticmethod
    def _sp_with(points):
        """A path in points: two curves meeting at (1000, 1000)."""
        def p(x, y):
            return f'<a:pt x="{x * E}" y="{y * E}"/>'

        body = (
            f'<a:moveTo>{p(0, 0)}</a:moveTo>'
            f'<a:cubicBezTo>{p(*points[0])}{p(*points[1])}{p(1000, 1000)}</a:cubicBezTo>'
            f'<a:cubicBezTo>{p(*points[2])}{p(1500, 2000)}{p(2000, 2000)}</a:cubicBezTo>'
        )
        geometry = (
            '<a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="0" t="0" r="r" b="b"/>'
            f'<a:pathLst><a:path w="{2000 * E}" h="{2000 * E}">{body}</a:path></a:pathLst></a:custGeom>'
        )
        return canvas.children(canvas.parse(
            canvas.wrap(shapes.shape_xml(2, "F", 0, 0, 2000 * E, 2000 * E, geometry), 0, 0, 2000 * E, 2000 * E).encode()
        ))[0]

    def test_opposite_equal_handles_are_symmetric(self):
        nodes = freeform.read_nodes(self._sp_with([(300, 0), (800, 800), (1200, 1200)]))
        assert nodes[3]["editing_type"] == "symmetric"

    def test_opposite_unequal_handles_are_smooth(self):
        nodes = freeform.read_nodes(self._sp_with([(300, 0), (900, 900), (1400, 1400)]))
        assert nodes[3]["editing_type"] == "smooth"

    def test_handles_at_an_angle_are_a_corner(self):
        nodes = freeform.read_nodes(self._sp_with([(300, 0), (800, 800), (1000, 1500)]))
        assert nodes[3]["editing_type"] == "corner"


# ---------------------------------------------------------------------------
# Grafting parts from several packages into one
# ---------------------------------------------------------------------------
class TestGraftingCarriesWhatAMemberRefersTo:
    @staticmethod
    def _picture_package(name, png):
        pic = (
            f'<a:pic><a:nvPicPr><a:cNvPr id="2" name="{name}"/><a:cNvPicPr/></a:nvPicPr>'
            '<a:blipFill><a:blip r:embed="rId1"/><a:stretch><a:fillRect/></a:stretch></a:blipFill>'
            '<a:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="100"/></a:xfrm>'
            '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></a:spPr></a:pic>'
        )
        return read(build(
            canvas.wrap(pic, 0, 0, 100, 100),
            parts={"clipboard/media/image1.png": png},
            drawing_rels=[Relationship("rId1", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image", "clipboard/media/image1.png")],
            defaults={"png": "image/png"},
        ))

    def test_two_pictures_named_image1_both_arrive(self):
        graft = Graft.empty()
        kids = []
        for name, png in (("P1", b"one"), ("P2", b"two")):
            package = self._picture_package(name, png)
            element = canvas.children(canvas.parse(package.drawing()))[0]
            kids.append(canvas.serialize(graft.take(package, element)))
        raw = build(
            canvas.wrap(shapes.group_xml(2, "G", 0, 0, 100, 100, kids), 0, 0, 100, 100),
            parts=graft.parts, drawing_rels=graft.drawing_rels,
            overrides=graft.overrides, defaults=graft.defaults, nested_rels=graft.nested_rels,
        )
        package = read(raw)
        media = sorted(n for n in package.parts if n.startswith("clipboard/media/"))
        assert len(media) == 2
        assert sorted(package.parts[m] for m in media) == [b"one", b"two"]
        rel_ids = [r.id for r in package.relationships(DRAWING_PART)]
        assert rel_ids == ["rId1", "rId2"]
        embeds = [b.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed")
                  for b in canvas.parse(package.drawing()).iter(f"{A}blip")]
        assert embeds == ["rId1", "rId2"]
        assert package.content_type(media[0]) == "image/png"

    def test_a_chart_member_carries_its_chart_part(self):
        package = fixture("chart")
        graft = Graft.empty()
        element = graft.take(package, canvas.children(canvas.parse(package.drawing()))[0])
        raw = build(
            canvas.wrap(shapes.group_xml(2, "G", 0, 0, 100, 100, [canvas.serialize(element)]), 0, 0, 100, 100),
            parts=graft.parts, drawing_rels=graft.drawing_rels,
            overrides=graft.overrides, defaults=graft.defaults, nested_rels=graft.nested_rels,
        )
        merged = read(raw)
        assert merged.chart_part() is not None
        assert charts.read_chart(merged.chart())["categories"][0] == "Category 1"
        # The theme part is deliberately left behind.
        assert not any("theme" in n for n in merged.parts)


class TestUnitsAndNames:
    def test_twelve_thousand_seven_hundred_emu_is_one_point(self):
        assert canvas.emu(1) == 12700
        assert canvas.pt(12700) == 1.0
        assert canvas.pt(12700 * 398.33333) == 398.33

    def test_a_chart_type_without_a_template_is_named(self):
        with pytest.raises(charts.ChartTypeError, match="9999"):
            charts.chart_xml(9999)

    def test_a_series_of_the_wrong_length_is_refused(self):
        with pytest.raises(ValueError, match="2 values for 3 categories"):
            charts.chart_xml(51, ["a", "b", "c"], [{"name": "s", "values": [1, 2]}])

    def test_names_are_escaped_in_the_xml(self):
        xml = shapes.shape_xml(2, 'A "<B>" & C', 0, 0, 1, 1)
        assert shapes.name_of(ET.fromstring(f'<r xmlns:a="{NS_A}">{xml}</r>')[0]) == 'A "<B>" & C'

    def test_a_degenerate_path_still_has_an_extent(self):
        _, _, _, cx, cy = freeform.build_geometry(0, 0, [{"seg_int": 0, "et_int": 0, "x1": 100, "y1": 0}], False)
        assert cx == 100 * E
        assert cy == 1
