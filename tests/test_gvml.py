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
    REL_THEME,
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


# ---------------------------------------------------------------------------
# Editing a chart PowerPoint wrote
# ---------------------------------------------------------------------------
# The five chart editors change one thing in the fixture's chart1.xml and
# leave the rest. What is pinned: the change reads back, the untouched parts
# are byte-for-byte the fixture's, and every element lands in schema order,
# which is the one mistake PowerPoint hides (it drops the paste silently).

C = f"{{{charts.NS_C}}}"


def chart_bytes():
    return fixture("chart").chart()


def combo_bytes():
    """The recorded chart with its third series lifted into a second plot
    group over a secondary value axis. Built in ``conftest`` because the
    macOS tests need the same chart."""
    from conftest import combo_chart_xml

    return combo_chart_xml(chart_bytes())


def _children(element):
    return [_local(c.tag) for c in element]


def _in_order(names, order):
    """True when ``names`` appear in the sequence ``order`` gives them."""
    ranks = [order.index(n) for n in names if n in order]
    return ranks == sorted(ranks)


class TestChartDataIsRewrittenInPlace:
    def test_new_categories_and_series_read_back(self):
        out = charts.set_data(chart_bytes(), ["Q1", "Q2"], [
            {"name": "Sales", "values": [120, 180]},
            {"name": "Cost", "values": [80, 90]},
        ])
        assert charts.read_chart(out) == {
            "categories": ["Q1", "Q2"],
            "series": [
                {"name": "Sales", "values": [120.0, 180.0]},
                {"name": "Cost", "values": [80.0, 90.0]},
            ],
        }

    def test_more_series_than_before_clone_the_last_without_its_formatting(self):
        coloured = charts.set_series(chart_bytes(), 3, color="#FF0000")
        out = charts.set_data(coloured, ["a"], [
            {"name": str(i), "values": [i]} for i in range(5)
        ])
        assert charts.series_count(out) == 5
        assert charts.read_series(out, 3)["color"] == "#FF0000"
        assert charts.read_series(out, 4)["color"] is None
        assert charts.read_series(out, 5)["color"] is None

    def test_fewer_series_than_before_drop_the_rest(self):
        out = charts.set_data(chart_bytes(), ["a"], [{"name": "only", "values": [1]}])
        assert charts.series_count(out) == 1

    def test_every_series_keeps_schema_order(self):
        out = charts.set_data(chart_bytes(), ["Q1"], [{"name": "S", "values": [1]}] * 2)
        for ser in ET.fromstring(out).iter(f"{C}ser"):
            assert _in_order(_children(ser), charts._SER_ORDER), _children(ser)

    def test_the_series_go_where_the_old_ones_were(self):
        out = charts.set_data(chart_bytes(), ["Q1"], [{"name": "S", "values": [1]}])
        plot = ET.fromstring(out).find(f"{C}chart/{C}plotArea/{C}barChart")
        assert _children(plot) == ["barDir", "grouping", "varyColors", "ser", "dLbls", "gapWidth", "axId", "axId"]

    def test_the_rest_of_the_chart_is_untouched(self):
        out = charts.set_data(chart_bytes(), ["Q1"], [{"name": "S", "values": [1]}])
        before = ET.fromstring(chart_bytes())
        after = ET.fromstring(out)
        for name in ("legend", "plotVisOnly", "dispBlanksAs"):
            assert ET.tostring(after.find(f"{C}chart/{C}{name}")) == ET.tostring(before.find(f"{C}chart/{C}{name}"))
        assert ET.tostring(after.find(f"{C}chart/{C}plotArea/{C}valAx")) == ET.tostring(before.find(f"{C}chart/{C}plotArea/{C}valAx"))

    def test_a_workbook_pointer_is_dropped_and_named(self):
        with_workbook = chart_bytes().replace(
            b"</c:chartSpace>",
            b'<c:externalData r:id="rId9"><c:autoUpdate val="0"/></c:externalData></c:chartSpace>',
        )
        assert charts.external_data_id(with_workbook) == "rId9"
        out = charts.set_data(with_workbook, ["Q1"], [{"name": "S", "values": [1]}])
        assert charts.external_data_id(out) is None

    def test_a_series_of_the_wrong_length_is_refused(self):
        with pytest.raises(ValueError, match="2 values for 1 categories"):
            charts.set_data(chart_bytes(), ["Q1"], [{"name": "S", "values": [1, 2]}])

    def test_the_prefixes_powerpoint_used_survive_the_round_trip(self):
        """``mc:Choice Requires="c14"`` names a prefix; renaming it breaks it."""
        out = charts.set_data(chart_bytes(), ["Q1"], [{"name": "S", "values": [1]}])
        assert b'<mc:Choice Requires="c14">' in out
        assert b"<c14:style" in out
        assert b"ns0:" not in out and b"ns1:" not in out

    def test_scatter_data_goes_into_x_and_y_values(self):
        scatter = charts.set_type(chart_bytes(), -4169)
        out = charts.set_data(scatter, [1.5, 2.5], [{"name": "Y", "values": [3, 4]}])
        ser = next(ET.fromstring(out).iter(f"{C}ser"))
        assert "cat" not in _children(ser) and "xVal" in _children(ser)
        assert charts.read_chart(out)["categories"] == ["1.5", "2.5"]


class TestChartTypeIsSwappedUnderTheData:
    @pytest.mark.parametrize("type_int,kind", [
        (4, "lineChart"), (57, "barChart"), (1, "areaChart"), (-4151, "radarChart"),
    ])
    def test_a_kind_with_the_same_axes_keeps_them(self, type_int, kind):
        out = charts.set_type(chart_bytes(), type_int)
        assert charts.kind_of(out) == kind
        assert charts.read_chart(out) == charts.read_chart(chart_bytes())
        root = ET.fromstring(out)
        plot_area = root.find(f"{C}chart/{C}plotArea")
        # The fixture's own axes, ids and all, still there and still referenced.
        assert _children(plot_area) == ["layout", kind, "catAx", "valAx"]
        plot_ids = [a.get("val") for a in plot_area.find(f"{C}{kind}").findall(f"{C}axId")]
        axis_ids = [a.find(f"{C}axId").get("val") for a in plot_area if _local(a.tag) in ("catAx", "valAx")]
        assert plot_ids == axis_ids

    def test_a_pie_has_no_axes_left(self):
        out = charts.set_type(chart_bytes(), 5)
        plot_area = ET.fromstring(out).find(f"{C}chart/{C}plotArea")
        assert _children(plot_area) == ["layout", "pieChart"]

    def test_a_scatter_gets_two_value_axes_and_numeric_x(self):
        out = charts.set_type(chart_bytes(), -4169)
        plot_area = ET.fromstring(out).find(f"{C}chart/{C}plotArea")
        assert _children(plot_area) == ["layout", "scatterChart", "valAx", "valAx"]
        assert charts.read_chart(out)["categories"] == ["1.0", "2.0", "3.0", "4.0"]

    def test_three_d_adds_a_view_and_two_d_removes_it(self):
        three_d = charts.set_type(chart_bytes(), 54)
        chart = ET.fromstring(three_d).find(f"{C}chart")
        assert "view3D" in _children(chart)
        assert _in_order(_children(chart), charts._CHART_ORDER)
        flat = charts.set_type(three_d, 51)
        assert "view3D" not in _children(ET.fromstring(flat).find(f"{C}chart"))

    def test_the_legend_and_title_survive_a_change_of_kind(self):
        titled = charts.format_chart(chart_bytes(), title="Hello", legend_position="top")
        out = charts.set_type(titled, 4)
        assert charts.chart_summary(out) == {
            "has_title": True, "title": "Hello", "has_legend": True, "legend_position": "t",
        }

    def test_a_kind_without_a_template_is_named(self):
        with pytest.raises(charts.ChartTypeError, match="-4100"):
            charts.set_type(chart_bytes(), -4100)


class TestAComboChartKeepsEveryGroupsSeries:
    """Two plot groups, the second over a secondary axis. The new kind has
    one plot, so every group's series goes into it; the secondary axis is
    what leaves, not the data plotted against it."""

    def test_the_fixture_is_a_combo_chart_with_a_secondary_axis(self):
        combo = combo_bytes()
        plot_area = ET.fromstring(combo).find(f"{C}chart/{C}plotArea")
        assert _children(plot_area) == [
            "layout", "barChart", "lineChart", "catAx", "valAx", "valAx", "catAx",
        ]
        assert charts.read_axis(combo, "secondary_value") is not None
        assert charts.series_count(combo) == 3

    @pytest.mark.parametrize("type_int,kind", [(4, "lineChart"), (57, "barChart")])
    def test_every_series_arrives_in_the_new_plot_in_the_same_order(self, type_int, kind):
        combo = combo_bytes()
        out = charts.set_type(combo, type_int)
        assert charts.kind_of(out) == kind
        # The order read_chart reports is the order the caller saw, and the
        # order ppt_change_chart_type's read-back check compares against.
        assert charts.read_chart(out) == charts.read_chart(combo)
        plot_area = ET.fromstring(out).find(f"{C}chart/{C}plotArea")
        assert _children(plot_area) == ["layout", kind, "catAx", "valAx"]
        assert len(plot_area.find(f"{C}{kind}").findall(f"{C}ser")) == 3
        # One plot, one pair of axes, so the secondary pair goes.
        assert charts.read_axis(out, "secondary_value") is None
        assert [s.get("val") for s in plot_area.find(f"{C}{kind}").findall(f"{C}axId")] == ["10", "20"]

    def test_the_series_are_renumbered_from_nothing(self):
        out = charts.set_type(combo_bytes(), 4)
        sers = list(ET.fromstring(out).iter(f"{C}ser"))
        assert [s.find(f"{C}idx").get("val") for s in sers] == ["0", "1", "2"]
        assert [s.find(f"{C}order").get("val") for s in sers] == ["0", "1", "2"]

    def test_a_pie_takes_every_group_too_and_loses_the_axes(self):
        out = charts.set_type(combo_bytes(), 5)
        assert [s["name"] for s in charts.read_chart(out)["series"]] == [
            "Series 1", "Series 2", "Series 3",
        ]
        assert _children(ET.fromstring(out).find(f"{C}chart/{C}plotArea")) == ["layout", "pieChart"]

    def test_a_longer_series_in_the_second_group_keeps_its_points(self):
        root = charts.load(combo_bytes())
        cache = root.find(f"{C}chart/{C}plotArea/{C}lineChart/{C}ser/{C}val/{C}numRef/{C}numCache")
        cache.find(f"{C}ptCount").set("val", "5")
        cache.append(ET.fromstring(f'<c:pt xmlns:c="{charts.NS_C}" idx="4"><c:v>9</c:v></c:pt>'))
        out = charts.set_type(charts.dump(root), 4)
        got = charts.read_chart(out)
        assert got["series"][2]["values"] == [2.0, 2.0, 3.0, 5.0, 9.0]
        # The categories stretch to the longest series rather than cutting it.
        assert got["categories"] == ["Category 1", "Category 2", "Category 3", "Category 4", ""]

    def test_a_group_with_categories_of_its_own_is_refused_by_name(self):
        """One plot carries one set of category labels. Relabelling a series
        or leaving it behind would both lose what the caller had, so the
        series is named and nothing is written."""
        root = charts.load(combo_bytes())
        line_cats = root.find(f"{C}chart/{C}plotArea/{C}lineChart/{C}ser/{C}cat")
        for i, pt in enumerate(line_cats.iter(f"{C}pt")):
            pt.find(f"{C}v").text = f"Week {i + 1}"
        with pytest.raises(PackageError, match="'Series 3' against categories of its own"):
            charts.set_type(charts.dump(root), 4)


class TestChartTitleAndLegend:
    def test_a_title_goes_first_in_the_chart(self):
        out = charts.format_chart(chart_bytes(), title="売上")
        chart = ET.fromstring(out).find(f"{C}chart")
        assert _children(chart)[:2] == ["title", "autoTitleDeleted"]
        assert charts.chart_summary(out)["title"] == "売上"

    def test_a_second_title_replaces_the_first(self):
        out = charts.format_chart(charts.format_chart(chart_bytes(), title="A"), title="B")
        chart = ET.fromstring(out).find(f"{C}chart")
        assert _children(chart).count("title") == 1
        assert charts.chart_summary(out)["title"] == "B"

    def test_the_legend_can_go_and_come_back_on_the_right(self):
        gone = charts.format_chart(chart_bytes(), has_legend=False)
        assert charts.chart_summary(gone)["has_legend"] is False
        back = charts.format_chart(gone, has_legend=True)
        assert charts.chart_summary(back)["legend_position"] == "r"
        chart = ET.fromstring(back).find(f"{C}chart")
        assert _in_order(_children(chart), charts._CHART_ORDER)

    @pytest.mark.parametrize("name,code", list(charts.LEGEND_POS.items()))
    def test_every_preset_position_is_written(self, name, code):
        out = charts.format_chart(chart_bytes(), legend_position=name)
        assert charts.chart_summary(out)["legend_position"] == code

    def test_a_position_on_a_chart_with_no_legend_is_the_windows_error(self):
        gone = charts.format_chart(chart_bytes(), has_legend=False)
        with pytest.raises(ValueError, match="Cannot set legend position when chart has no legend"):
            charts.format_chart(gone, legend_position="top")

    def test_has_legend_and_a_position_in_one_call_apply_in_that_order(self):
        gone = charts.format_chart(chart_bytes(), has_legend=False)
        out = charts.format_chart(gone, has_legend=True, legend_position="bottom")
        assert charts.chart_summary(out)["legend_position"] == "b"


class TestChartAxes:
    def test_the_value_axis_takes_every_mapped_field(self):
        out, applied = charts.format_axis(
            chart_bytes(), "value", title="円", min_scale=0, max_scale=200,
            major_unit=50, minor_unit=10, major_tick_mark="cross",
            minor_tick_mark="inside", reverse_order=True, log_scale=True,
            log_base=2, number_format="#,##0",
        )
        assert applied == [
            "title", "min_scale", "max_scale", "major_unit", "minor_unit",
            "major_tick_mark", "minor_tick_mark", "reverse_order", "log_scale",
            "log_base", "number_format",
        ]
        assert charts.read_axis(out, "value") == {
            "title": "円", "min_scale": 0.0, "max_scale": 200.0, "major_unit": 50.0,
            "minor_unit": 10.0, "tick_label_spacing": None, "tick_mark_spacing": None,
            "major_tick_mark": "cross", "minor_tick_mark": "inside",
            "reverse_order": True, "log_scale": True, "log_base": 2.0,
            "number_format": "#,##0",
        }

    def test_the_axis_and_its_scaling_stay_in_schema_order(self):
        out, _ = charts.format_axis(
            chart_bytes(), "value", title="t", min_scale=1, max_scale=9,
            major_unit=2, minor_unit=1, log_scale=True, reverse_order=False,
            number_format="0",
        )
        axis = ET.fromstring(out).find(f"{C}chart/{C}plotArea/{C}valAx")
        assert _in_order(_children(axis), charts._AXIS_ORDER), _children(axis)
        assert _children(axis.find(f"{C}scaling")) == ["logBase", "orientation", "max", "min"]

    def test_the_category_axis_takes_the_spacing_fields(self):
        out, applied = charts.format_axis(
            chart_bytes(), "category", tick_label_spacing=3, tick_mark_spacing=2,
        )
        assert applied == ["tick_label_spacing", "tick_mark_spacing"]
        got = charts.read_axis(out, "category")
        assert (got["tick_label_spacing"], got["tick_mark_spacing"]) == (3, 2)
        axis = ET.fromstring(out).find(f"{C}chart/{C}plotArea/{C}catAx")
        assert _in_order(_children(axis), charts._AXIS_ORDER), _children(axis)

    def test_turning_the_log_scale_off_removes_the_base(self):
        on, _ = charts.format_axis(chart_bytes(), "value", log_scale=True)
        off, _ = charts.format_axis(on, "value", log_scale=False)
        assert charts.read_axis(off, "value")["log_scale"] is False

    def test_a_pie_has_no_axis_and_says_so_the_windows_way(self):
        pie = charts.set_type(chart_bytes(), 5)
        with pytest.raises(ValueError, match="Axis 'value' is not available on this chart type"):
            charts.format_axis(pie, "value", title="x")

    def test_a_scatter_answers_category_with_its_x_axis(self):
        scatter = charts.set_type(chart_bytes(), -4169)
        out, _ = charts.format_axis(scatter, "category", title="x")
        first = [a for a in ET.fromstring(out).find(f"{C}chart/{C}plotArea") if _local(a.tag) == "valAx"][0]
        assert charts._title_text(first.find(f"{C}title")) == "x"

    def test_no_secondary_or_series_axis_on_the_fixture(self):
        assert charts.read_axis(chart_bytes(), "secondary_value") is None
        assert charts.read_axis(chart_bytes(), "series") is None


class TestOneChartSeries:
    def test_a_bar_series_takes_a_fill_labels_and_a_weight(self):
        out = charts.set_series(chart_bytes(), 2, color="#ff0000", show_data_labels=True, line_weight=1.5)
        assert charts.read_series(out, 2) == {"color": "#FF0000", "line_weight": 1.5, "show_data_labels": True}
        assert charts.read_series(out, 1) == {"color": None, "line_weight": None, "show_data_labels": None}
        ser = list(ET.fromstring(out).iter(f"{C}ser"))[1]
        assert _children(ser) == ["idx", "order", "tx", "spPr", "invertIfNegative", "dLbls", "cat", "val"]
        assert _in_order(_children(ser.find(f"{C}dLbls")), charts._DLBLS_ORDER)

    def test_a_line_series_takes_its_colour_on_the_line(self):
        """PowerPoint keeps only the line's colour on a line series; a fill
        written beside it is dropped on the paste. Measured, so pinned."""
        out = charts.set_series(charts.set_type(chart_bytes(), 4), 1, color="#00FF00")
        sp = next(ET.fromstring(out).iter(f"{C}ser")).find(f"{C}spPr")
        assert _children(sp) == ["ln"]
        assert charts.read_series(out, 1)["color"] == "#00FF00"

    def test_labels_can_be_turned_off_again(self):
        on = charts.set_series(chart_bytes(), 1, show_data_labels=True)
        off = charts.set_series(on, 1, show_data_labels=False)
        assert charts.read_series(off, 1)["show_data_labels"] is False

    def test_an_index_past_the_end_is_the_windows_error(self):
        with pytest.raises(ValueError, match="series_index 4 out of range \\(chart has 3 series\\)"):
            charts.set_series(chart_bytes(), 4, color="#000000")

    def test_a_colour_that_is_not_hex_is_refused(self):
        with pytest.raises(ValueError, match="#RRGGBB"):
            charts.set_series(chart_bytes(), 1, color="red")


class TestPackagePartsCanBeLetGoOf:
    def test_removing_a_relationship_drops_its_part(self):
        package = fixture("chart")
        drawing = package.drawing_part()
        theme = [r for r in package.relationships(drawing) if r.type == REL_THEME][0]
        removed = package.remove_relationship(drawing, theme.id)
        assert removed == "clipboard/theme/theme1.xml"
        assert removed not in package.parts
        assert [r.type for r in package.relationships(drawing)] == [REL_CHART]
        package.validate()

    def test_an_unknown_relationship_changes_nothing(self):
        package = fixture("chart")
        before = dict(package.parts)
        assert package.remove_relationship(package.drawing_part(), "rId99") is None
        assert package.parts == before


# ---------------------------------------------------------------------------
# Editing a freeform path PowerPoint wrote
# ---------------------------------------------------------------------------
# The four node editors work on a Path read out of the fixture's custGeom,
# in slide points with Windows numbering, and write it back into the same
# a:sp. The fixture is the closed four segment shape recorded from
# PowerPoint: line, curve, curve, line back to the start, nine nodes.

def fixture_path():
    sp = canvas.children(fixture_canvas("freeform"))[0]
    return sp, freeform.read_path(sp)


class TestAPathReadsAndWritesAsTheSameNodes:
    def test_the_fixture_reads_as_one_closed_path_of_four_segments(self):
        _, path = fixture_path()
        assert path.closed is True
        assert [s[0] for s in path.segments] == ["line", "curve", "curve", "line"]
        assert freeform.node_count(path) == 9
        assert freeform.nodes_of([path]) == freeform.read_nodes(fixture_path()[0])

    def test_writing_a_path_back_resizes_the_shape_to_it(self):
        sp, path = fixture_path()
        moved = freeform.set_node_position(path, 2, 800.0, 100.0)
        x, y, cx, cy = freeform.write_path(sp, moved)
        assert shapes.xfrm_of(sp) == (x, y, cx, cy)
        # The box holds every point, the moved node's handle included.
        nodes = freeform.nodes_of([moved])
        right = max(n["x"] for n in nodes)
        assert (x, cx) == (canvas.emu(398.33), canvas.emu(right) - canvas.emu(398.33))
        assert freeform.read_nodes(sp) == freeform.nodes_of([moved])
        # The path's own box equals the extent, so 1 EMU is 1/12700 pt.
        path_el = sp.find(f"{A}spPr/{A}custGeom/{A}pathLst/{A}path")
        assert (int(path_el.get("w")), int(path_el.get("h"))) == (cx, cy)

    def test_the_path_attributes_and_the_rest_of_the_shape_survive(self):
        sp, path = fixture_path()
        path_el = sp.find(f"{A}spPr/{A}custGeom/{A}pathLst/{A}path")
        path_el.set("fill", "none")
        path = freeform.read_path(sp)
        style_before = ET.tostring(sp.find(f"{A}style"))
        freeform.write_path(sp, freeform.delete_node(path, 2))
        assert sp.find(f"{A}spPr/{A}custGeom/{A}pathLst/{A}path").get("fill") == "none"
        assert ET.tostring(sp.find(f"{A}style")) == style_before
        assert sp.find(f"{A}spPr/{A}custGeom/{A}avLst") is not None

    def test_two_subpaths_are_refused_rather_than_flattened(self):
        sp, _ = fixture_path()
        path_lst = sp.find(f"{A}spPr/{A}custGeom/{A}pathLst")
        path_lst.append(ET.fromstring(
            f'<a:path xmlns:a="{NS_A}" w="10" h="10"><a:moveTo><a:pt x="0" y="0"/></a:moveTo>'
            '<a:lnTo><a:pt x="10" y="10"/></a:lnTo></a:path>'
        ))
        with pytest.raises(PackageError, match="holds 2 paths"):
            freeform.read_path(sp)
        # Reading alone still numbers them continuously, as Windows does.
        assert len(freeform.read_nodes(sp)) == 11


class TestNodeEdits:
    def test_moving_an_anchor_takes_its_handles_along(self):
        _, path = fixture_path()
        before = freeform.nodes_of([path])
        moved = freeform.set_node_position(path, 5, 600.0, 300.0)
        after = freeform.nodes_of([moved])
        assert (after[4]["x"], after[4]["y"]) == (600.0, 300.0)
        dx, dy = 600.0 - before[4]["x"], 300.0 - before[4]["y"]
        # Node 4 is the handle arriving at node 5, node 6 the one leaving it.
        assert (after[3]["x"], after[3]["y"]) == (round(before[3]["x"] + dx, 2), round(before[3]["y"] + dy, 2))
        assert (after[5]["x"], after[5]["y"]) == (round(before[5]["x"] + dx, 2), round(before[5]["y"] + dy, 2))
        assert freeform.node_count(moved) == 9

    def test_moving_a_control_point_moves_only_itself(self):
        _, path = fixture_path()
        before = freeform.nodes_of([path])
        after = freeform.nodes_of([freeform.set_node_position(path, 3, 1.0, 2.0)])
        assert (after[2]["x"], after[2]["y"]) == (1.0, 2.0)
        assert [n for i, n in enumerate(after) if i != 2] == [n for i, n in enumerate(before) if i != 2]

    def test_inserting_a_line_after_a_vertex_adds_one_node(self):
        _, path = fixture_path()
        out = freeform.insert_node(path, 2, 0, 0, 10.0, 20.0)
        nodes = freeform.nodes_of([out])
        assert len(nodes) == 10
        assert (nodes[2]["x"], nodes[2]["y"]) == (10.0, 20.0)
        # The curve that followed node 2 now starts at the new point.
        assert nodes[2]["segment_type"] == "curve"

    def test_inserting_a_curve_adds_three_nodes(self):
        _, path = fixture_path()
        auto = freeform.insert_node(path, 1, 1, 0, 10.0, 20.0)
        assert freeform.node_count(auto) == 12
        corner = freeform.insert_node(path, 1, 1, 1, 1.0, 1.0, 2.0, 2.0, 3.0, 3.0)
        # After node 1, the start, so it is the first segment now.
        assert corner.segments[0] == ("curve", (1.0, 1.0), (2.0, 2.0), (3.0, 3.0))
        assert corner.segments[1] == path.segments[0]

    def test_inserting_after_the_last_node_appends(self):
        _, path = fixture_path()
        out = freeform.insert_node(path, 9, 0, 0, 5.0, 5.0)
        assert out.segments[-1] == ("line", (5.0, 5.0))

    def test_inserting_after_a_control_point_is_refused(self):
        _, path = fixture_path()
        with pytest.raises(ValueError, match="after_index 3 is a curve's control point"):
            freeform.insert_node(path, 3, 0, 0, 5.0, 5.0)

    def test_deleting_a_vertex_removes_the_segment_after_it(self):
        _, path = fixture_path()
        out = freeform.delete_node(path, 2)
        assert [s[0] for s in out.segments] == ["line", "curve", "line"]
        assert out.segments[0][1] == path.segments[1][3]
        assert freeform.node_count(out) == 6

    def test_deleting_a_control_point_deletes_the_whole_curve(self):
        """The tool promises the other control point and the end point go
        too, so three nodes leave and no straight edge is put in their place."""
        _, path = fixture_path()
        out = freeform.delete_node(path, 3)
        assert [s[0] for s in out.segments] == ["line", "curve", "line"]
        assert out.segments[1] == path.segments[2]
        assert freeform.node_count(out) == 6
        # Either control point stands for the same curve.
        assert freeform.delete_node(path, 4).segments == out.segments

    def test_a_path_of_one_curve_cannot_lose_its_control_point(self):
        one = freeform.Path((0.0, 0.0), [("curve", (1.0, 1.0), (2.0, 2.0), (3.0, 3.0))])
        with pytest.raises(ValueError, match="at least two nodes"):
            freeform.delete_node(one, 2)

    def test_deleting_the_first_node_starts_the_path_at_the_next(self):
        _, path = fixture_path()
        out = freeform.delete_node(path, 1)
        assert out.start == path.segments[0][1]
        assert len(out.segments) == 3

    def test_deleting_the_last_node_removes_its_segment(self):
        _, path = fixture_path()
        out = freeform.delete_node(path, 9)
        assert len(out.segments) == 3
        assert freeform.node_count(out) == 8

    def test_a_path_cannot_be_deleted_down_to_one_node(self):
        two = freeform.Path((0.0, 0.0), [("line", (1.0, 1.0))])
        with pytest.raises(ValueError, match="at least two nodes"):
            freeform.delete_node(two, 2)

    def test_a_line_becomes_a_curve_with_handles_on_the_chord(self):
        _, path = fixture_path()
        out = freeform.set_segment_type(path, 1, 1)
        seg = out.segments[0]
        start, end = path.start, path.segments[0][1]
        assert seg[0] == "curve"
        assert seg[1] == pytest.approx((start[0] + (end[0] - start[0]) / 3, start[1] + (end[1] - start[1]) / 3))
        assert seg[2] == pytest.approx((start[0] + 2 * (end[0] - start[0]) / 3, start[1] + 2 * (end[1] - start[1]) / 3))
        assert freeform.node_count(out) == 11

    def test_a_curve_becomes_a_line_and_loses_two_nodes(self):
        _, path = fixture_path()
        out = freeform.set_segment_type(path, 2, 0)
        assert out.segments[1] == ("line", path.segments[1][3])
        assert freeform.node_count(out) == 7

    def test_a_control_point_stands_for_its_own_curve(self):
        _, path = fixture_path()
        assert freeform.set_segment_type(path, 3, 0).segments == freeform.set_segment_type(path, 2, 0).segments

    def test_the_last_node_has_no_segment_to_set(self):
        _, path = fixture_path()
        with pytest.raises(ValueError, match="node_index 9 is the last node"):
            freeform.set_segment_type(path, 9, 1)

    @pytest.mark.parametrize("call", [
        lambda p: freeform.set_node_position(p, 10, 0.0, 0.0),
        lambda p: freeform.delete_node(p, 0),
        lambda p: freeform.set_segment_type(p, 10, 0),
    ])
    def test_an_index_out_of_range_is_the_windows_error(self, call):
        _, path = fixture_path()
        with pytest.raises(ValueError, match="out of range \\(shape has 9 nodes\\)"):
            call(path)
