"""Shared fixtures and path setup for the test suite."""

import os
import sys

# Allow tests to import from src/ without installing the package.
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "src"))


def combo_chart_xml(chart_xml_bytes):
    """A combo chart built from a plain one, for the tests that need two plots.

    The third series is lifted out of the bar group into a line group over a
    secondary value axis, which is the shape PowerPoint writes for a combo
    chart and the one the recorded fixtures do not hold. It is built here
    rather than recorded because ``tests/fixtures/gvml/chart.gvml.zip`` is
    pinned as having no secondary axis, and because both the pure tests and
    the macOS ones need it.
    """
    import xml.etree.ElementTree as ET

    from gvml import charts

    c = f"{{{charts.NS_C}}}"
    root = charts.load(chart_xml_bytes)
    plot_area = root.find(f"{c}chart/{c}plotArea")
    bar = plot_area.find(f"{c}barChart")
    third = bar.findall(f"{c}ser")[2]
    bar.remove(third)
    line = ET.fromstring(
        f'<c:lineChart xmlns:c="{charts.NS_C}"><c:grouping val="standard"/>'
        '<c:varyColors val="0"/><c:marker val="1"/>'
        '<c:axId val="30"/><c:axId val="40"/></c:lineChart>'
    )
    line.insert(2, third)
    plot_area.insert(list(plot_area).index(bar) + 1, line)
    # The secondary pair: a value axis on the right, and a category axis of
    # its own, hidden, that it crosses at the maximum.
    plot_area.append(ET.fromstring(
        f'<c:valAx xmlns:c="{charts.NS_C}"><c:axId val="40"/>'
        '<c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/>'
        '<c:axPos val="r"/><c:numFmt formatCode="General" sourceLinked="1"/>'
        '<c:majorTickMark val="out"/><c:minorTickMark val="none"/>'
        '<c:tickLblPos val="nextTo"/><c:crossAx val="30"/><c:crosses val="max"/>'
        '<c:crossBetween val="between"/></c:valAx>'
    ))
    plot_area.append(ET.fromstring(
        f'<c:catAx xmlns:c="{charts.NS_C}"><c:axId val="30"/>'
        '<c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="1"/>'
        '<c:axPos val="b"/><c:majorTickMark val="out"/><c:minorTickMark val="none"/>'
        '<c:tickLblPos val="nextTo"/><c:crossAx val="40"/><c:crosses val="autoZero"/>'
        '<c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/>'
        '<c:noMultiLvlLbl val="0"/></c:catAx>'
    ))
    return charts.dump(root)
