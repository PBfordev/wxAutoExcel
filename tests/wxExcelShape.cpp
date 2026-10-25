/////////////////////////////////////////////////////////////////////////////
// Purpose:     Tests for wxExcelShapes, wxExcelShape, and wxExcelShapeRange
// Author:      Claude Opus 5.5 (Anthropic), under the direction of PB
// Copyright:   (c) 2026 PB <pbfordev@gmail.com>
// License:     MIT license
/////////////////////////////////////////////////////////////////////////////

#include <catch2/catch_test_macros.hpp>

#include "excel_test_fixture.h"

#if WXAUTOEXCEL_USE_SHAPES

using namespace wxAutoExcel;

namespace
{

wxExcelShapes GetShapes(ExcelTestWorkbook& book)
{
    wxExcelShapes shapes = book.GetWorksheet().GetShapes();
    REQUIRE(shapes);
    return shapes;
}

long GetCount(wxExcelShapes& shapes)
{
    const long count = shapes.GetCount();
    REQUIRE(shapes);
    return count;
}

wxExcelShape AddNamedRectangle(wxExcelShapes& shapes, const wxString& name,
                               double left, double top)
{
    wxExcelShape shape = shapes.AddShape(msoShapeRectangle, left, top, 100, 50);
    REQUIRE(shape);
    shape.SetName(name);
    REQUIRE(shape);
    return shape;
}

wxString GetText(wxExcelShape& shape)
{
    wxExcelTextRange2 textRange = shape.GetTextFrame2().GetTextRange();
    REQUIRE(textRange);
    const wxString text = textRange.GetText();
    REQUIRE(textRange);
    return text;
}

void SetText(wxExcelShape& shape, const wxString& text)
{
    wxExcelTextRange2 textRange = shape.GetTextFrame2().GetTextRange();
    REQUIRE(textRange);
    textRange.SetText(text);
    REQUIRE(textRange);
}

} // unnamed namespace

TEST_CASE("wxExcelShapes adds shapes of different types",
          "[excel][shapes]")
{
    ExcelTestWorkbook book;
    REQUIRE(book.IsReady());

    wxExcelShapes shapes = GetShapes(book);
    CHECK(GetCount(shapes) == 0);

    wxExcelShape rectangle = shapes.AddShape(msoShapeRectangle, 10, 20, 100, 50);
    REQUIRE(rectangle);
    CHECK(rectangle.GetType() == msoAutoShape);
    REQUIRE(rectangle);
    CHECK(rectangle.GetAutoShapeType() == msoShapeRectangle);
    REQUIRE(rectangle);

    wxExcelShape line = shapes.AddLine(10, 100, 200, 150);
    REQUIRE(line);
    CHECK(line.GetType() == msoLine);
    REQUIRE(line);

    wxExcelShape textbox = shapes.AddTextbox(msoTextOrientationHorizontal, 250, 20, 150, 40);
    REQUIRE(textbox);
    CHECK(textbox.GetType() == msoTextBox);
    REQUIRE(textbox);
    SetText(textbox, "Some text");
    CHECK(GetText(textbox) == "Some text");

    CHECK(GetCount(shapes) == 3);

    rectangle.SetAutoShapeType(msoShapeOval);
    REQUIRE(rectangle);
    CHECK(rectangle.GetAutoShapeType() == msoShapeOval);
    REQUIRE(rectangle);

    // the shapes with similar names have different types
    wxExcelShape rounded = shapes.AddShape(msoShapeRound2DiagRectangle, 10, 200, 60, 40);
    REQUIRE(rounded);
    CHECK(rounded.GetAutoShapeType() == msoShapeRound2DiagRectangle);
    REQUIRE(rounded);
    wxExcelShape snipped = shapes.AddShape(msoShapeSnip2DiagRectangle, 100, 200, 60, 40);
    REQUIRE(snipped);
    CHECK(snipped.GetAutoShapeType() == msoShapeSnip2DiagRectangle);
    REQUIRE(snipped);
    CHECK(rounded.GetAutoShapeType() != snipped.GetAutoShapeType());
}

TEST_CASE("wxExcelShape is positioned, sized, named, and found by its name",
          "[excel][shapes]")
{
    ExcelTestWorkbook book;
    REQUIRE(book.IsReady());

    wxExcelShapes shapes = GetShapes(book);
    wxExcelShape shape = shapes.AddShape(msoShapeRectangle, 10, 20, 100, 50);
    REQUIRE(shape);

    CHECK(shape.GetLeft() == 10.0);
    CHECK(shape.GetTop() == 20.0);
    CHECK(shape.GetWidth() == 100.0);
    CHECK(shape.GetHeight() == 50.0);
    REQUIRE(shape);

    shape.SetLeft(200.0);
    shape.SetTop(150.0);
    shape.SetWidth(80.0);
    shape.SetHeight(40.0);
    REQUIRE(shape);
    CHECK(shape.GetLeft() == 200.0);
    CHECK(shape.GetTop() == 150.0);
    CHECK(shape.GetWidth() == 80.0);
    CHECK(shape.GetHeight() == 40.0);
    REQUIRE(shape);

    shape.IncrementLeft(-190.0);
    shape.IncrementTop(-145.0);
    REQUIRE(shape);
    CHECK(shape.GetLeft() == 10.0);
    CHECK(shape.GetTop() == 5.0);
    REQUIRE(shape);

    wxExcelRange topLeft = shape.GetTopLeftCell();
    REQUIRE(topLeft);
    CHECK(topLeft.GetAddress() == "$A$1");
    REQUIRE(topLeft);

    shape.SetName("Box");
    REQUIRE(shape);
    CHECK(shape.GetName() == "Box");
    REQUIRE(shape);

    wxExcelShape byName = shapes.Item("Box");
    REQUIRE(byName);
    CHECK(byName.GetLeft() == 10.0);
    REQUIRE(byName);

    shape.SetRotation(45.0);
    REQUIRE(shape);
    CHECK(shape.GetRotation() == 45.0);
    REQUIRE(shape);

    shape.Flip(msoFlipHorizontal);
    REQUIRE(shape);
    CHECK(shape.GetHorizontalFlip() == msoTrue);
    REQUIRE(shape);
}

TEST_CASE("wxExcelShape is duplicated, reordered, and deleted",
          "[excel][shapes]")
{
    ExcelTestWorkbook book;
    REQUIRE(book.IsReady());

    wxExcelShapes shapes = GetShapes(book);
    wxExcelShape first = AddNamedRectangle(shapes, "First", 10, 10);
    wxExcelShape second = AddNamedRectangle(shapes, "Second", 40, 40);
    SetText(first, "text");

    CHECK(first.GetZOrderPosition() == 1);
    CHECK(second.GetZOrderPosition() == 2);
    REQUIRE(first);
    REQUIRE(second);

    first.ZOrder(msoBringToFront);
    REQUIRE(first);
    CHECK(first.GetZOrderPosition() == 2);
    CHECK(second.GetZOrderPosition() == 1);
    REQUIRE(first);
    REQUIRE(second);

    wxExcelShape copy = first.Duplicate();
    REQUIRE(copy);
    CHECK(GetCount(shapes) == 3);
    CHECK(GetText(copy) == "text");

    second.Delete();
    REQUIRE(second);
    CHECK(GetCount(shapes) == 2);

    {
        wxLogNull noLog;
        wxExcelShape deleted = shapes.Item("Second");
        CHECK_FALSE(deleted);
    }
}

TEST_CASE("wxExcelShapeRange groups and ungroups shapes",
          "[excel][shapes][shaperange]")
{
    ExcelTestWorkbook book;
    REQUIRE(book.IsReady());

    wxExcelShapes shapes = GetShapes(book);
    AddNamedRectangle(shapes, "Left", 10, 10);
    AddNamedRectangle(shapes, "Right", 150, 10);
    AddNamedRectangle(shapes, "Alone", 300, 10);

    wxVector<wxString> names;
    names.push_back("Left");
    names.push_back("Right");
    wxExcelShapeRange range = shapes.GetRange(names);
    REQUIRE(range);
    CHECK(range.GetCount() == 2);
    REQUIRE(range);

    wxExcelShape group = range.Group();
    REQUIRE(group);
    CHECK(group.GetType() == msoGroup);
    REQUIRE(group);
    CHECK(GetCount(shapes) == 2);

    wxExcelGroupShapes items = group.GetGroupItems();
    REQUIRE(items);
    CHECK(items.GetCount() == 2);
    REQUIRE(items);

    // the group spans its shapes
    CHECK(group.GetLeft() == 10.0);
    CHECK(group.GetWidth() == 240.0);
    REQUIRE(group);

    wxExcelShapeRange ungrouped = group.Ungroup();
    REQUIRE(ungrouped);
    CHECK(ungrouped.GetCount() == 2);
    REQUIRE(ungrouped);
    CHECK(GetCount(shapes) == 3);

    wxVector<long> indices;
    indices.push_back(1);
    indices.push_back(3);
    wxExcelShapeRange byIndex = shapes.GetRange(indices);
    REQUIRE(byIndex);
    CHECK(byIndex.GetCount() == 2);
    REQUIRE(byIndex);

    byIndex.Delete();
    REQUIRE(byIndex);
    CHECK(GetCount(shapes) == 1);
}

#if WXAUTOEXCEL_USE_CHARTS

TEST_CASE("wxExcelShapes adds a chart",
          "[excel][shapes][chart]")
{
    ExcelTestWorkbook book;
    REQUIRE(book.IsReady());

    wxExcelRange data = book.GetRange("A1:A3");
    data.SetValue(wxVariant(1.0));
    REQUIRE(data);

    wxExcelShapes shapes = GetShapes(book);
    wxExcelShape shape = shapes.AddChart2();
    REQUIRE(shape);
    CHECK(shape.GetHasChart() == msoTrue);
    REQUIRE(shape);

    wxExcelChart chart = shape.GetChart();
    REQUIRE(chart);
    chart.SetSourceData(data);
    REQUIRE(chart);

    // the chart is a chart object too
    wxExcelChartObjects chartObjects = book.GetWorksheet().ChartObjects();
    REQUIRE(chartObjects);
    CHECK(chartObjects.GetCount() == 1);
    REQUIRE(chartObjects);
}

#endif // WXAUTOEXCEL_USE_CHARTS

#endif // WXAUTOEXCEL_USE_SHAPES
