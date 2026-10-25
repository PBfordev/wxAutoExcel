/////////////////////////////////////////////////////////////////////////////
// Purpose:     Tests for adding shapes of other types than AutoShapes, their
//              formats (callouts, connectors, form controls, nodes, text
//              effects, pictures, shadows, glows, soft edges, reflections,
//              text frames, columns, paragraphs, bullets, and tab stops),
//              and aligning, distributing, scaling, and rotating shapes
// Author:      Claude Opus 5.5 (Anthropic), under the direction of PB
// Copyright:   (c) 2026 PB <pbfordev@gmail.com>
// License:     MIT license
/////////////////////////////////////////////////////////////////////////////

#include <catch2/catch_approx.hpp>
#include <catch2/catch_test_macros.hpp>

#include "excel_test_fixture.h"
#include "temp_files.h"

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

wxExcelShape AddRectangle(wxExcelShapes& shapes, double left, double top,
                          double width = 100, double height = 50)
{
    wxExcelShape shape = shapes.AddShape(msoShapeRectangle, left, top, width, height);
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

long GetNodeCount(wxExcelShape& shape)
{
    wxExcelShapeNodes nodes = shape.GetNodes();
    REQUIRE(nodes);
    const long count = nodes.GetCount();
    REQUIRE(nodes);
    return count;
}

} // unnamed namespace

TEST_CASE("wxExcelShapes adds text boxes, labels, callouts, and text effects",
          "[excel][shapes][text]")
{
    ExcelTestWorkbook book;
    REQUIRE(book.IsReady());

    wxExcelShapes shapes = GetShapes(book);

    wxExcelShape textBox = shapes.AddTextbox(msoTextOrientationHorizontal, 10, 10, 100, 30);
    REQUIRE(textBox);
    CHECK(textBox.GetType() == msoTextBox);
    SetText(textBox, "Text box");
    CHECK(GetText(textBox) == "Text box");

    wxExcelShape label = shapes.AddLabel(msoTextOrientationHorizontal, 10, 50, 100, 30);
    REQUIRE(label);
    SetText(label, "Label");
    CHECK(GetText(label) == "Label");

    // Excel reports msoCalloutOne for msoCalloutTwo, so use a callout with
    // a line of two segments, whose type it reports as given
    wxExcelShape callout = shapes.AddCallout(msoCalloutThree, 150, 10, 100, 50);
    REQUIRE(callout);
    CHECK(callout.GetType() == msoCallout);
    SetText(callout, "Callout");
    CHECK(GetText(callout) == "Callout");

    wxExcelCalloutFormat calloutFormat = callout.GetCallout();
    REQUIRE(calloutFormat);
    CHECK(calloutFormat.GetType() == msoCalloutThree);
    calloutFormat.SetAccent(msoTrue);
    calloutFormat.SetBorder(msoTrue);
    calloutFormat.SetAngle(msoCalloutAngle45);
    calloutFormat.PresetDrop(msoCalloutDropCenter);
    calloutFormat.SetGap(5);
    REQUIRE(calloutFormat);
    CHECK(calloutFormat.GetAccent() == msoTrue);
    CHECK(calloutFormat.GetBorder() == msoTrue);
    CHECK(calloutFormat.GetAngle() == msoCalloutAngle45);
    CHECK(calloutFormat.GetDropType() == msoCalloutDropCenter);
    CHECK(calloutFormat.GetGap() == Catch::Approx(5.0));
    REQUIRE(calloutFormat);

    wxExcelShape textEffect = shapes.AddTextEffect(msoTextEffect1, "Effect", "Arial", 24,
                                                   msoTrue, msoFalse, 10, 100);
    REQUIRE(textEffect);
    wxExcelTextEffectFormat textEffectFormat = textEffect.GetTextEffect();
    REQUIRE(textEffectFormat);
    CHECK(textEffectFormat.GetText() == "Effect");
    CHECK(textEffectFormat.GetFontName() == "Arial");
    CHECK(textEffectFormat.GetFontSize() == 24.0);
    CHECK(textEffectFormat.GetFontBold() == msoTrue);
    CHECK(textEffectFormat.GetFontItalic() == msoFalse);
    REQUIRE(textEffectFormat);

    textEffectFormat.SetText("Changed");
    textEffectFormat.SetFontItalic(msoTrue);
    textEffectFormat.SetFontSize(30);
    REQUIRE(textEffectFormat);
    CHECK(textEffectFormat.GetText() == "Changed");
    CHECK(textEffectFormat.GetFontItalic() == msoTrue);
    CHECK(textEffectFormat.GetFontSize() == 30.0);
}

TEST_CASE("wxExcelConnectorFormat connects and disconnects shapes",
          "[excel][shapes][connector]")
{
    ExcelTestWorkbook book;
    REQUIRE(book.IsReady());

    wxExcelShapes shapes = GetShapes(book);
    wxExcelShape first = AddRectangle(shapes, 10, 10);
    first.SetName("First");
    wxExcelShape second = AddRectangle(shapes, 200, 100);
    second.SetName("Second");
    REQUIRE(first);
    REQUIRE(second);

    wxExcelShape connector = shapes.AddConnector(msoConnectorElbow, 0, 0, 10, 10);
    REQUIRE(connector);
    CHECK(connector.GetConnector() == msoTrue);

    wxExcelConnectorFormat format = connector.GetConnectorFormat();
    REQUIRE(format);
    CHECK(format.GetType() == msoConnectorElbow);
    CHECK(format.GetBeginConnected() == msoFalse);
    REQUIRE(format);

    format.BeginConnect(first, 1);
    format.EndConnect(second, 3);
    REQUIRE(format);
    CHECK(format.GetBeginConnected() == msoTrue);
    CHECK(format.GetEndConnected() == msoTrue);
    CHECK(format.GetBeginConnectedShape().GetName() == "First");
    CHECK(format.GetEndConnectedShape().GetName() == "Second");
    CHECK(format.GetBeginConnectionSite() == 1);
    CHECK(format.GetEndConnectionSite() == 3);
    REQUIRE(format);

    connector.RerouteConnections();
    REQUIRE(connector);
    CHECK(format.GetBeginConnected() == msoTrue);

    format.BeginDisconnect();
    format.EndDisconnect();
    REQUIRE(format);
    CHECK(format.GetBeginConnected() == msoFalse);
    CHECK(format.GetEndConnected() == msoFalse);
}

TEST_CASE("wxExcelControlFormat changes form controls",
          "[excel][shapes][controls]")
{
    ExcelTestWorkbook book;
    REQUIRE(book.IsReady());

    wxExcelShapes shapes = GetShapes(book);

    SECTION("check box linked to a cell")
    {
        wxExcelShape checkBox = shapes.AddFormControl(xlCheckBox, 10, 10, 100, 20);
        REQUIRE(checkBox);
        CHECK(checkBox.GetType() == msoFormControl);
        CHECK(checkBox.GetFormControlType() == xlCheckBox);

        wxExcelControlFormat control = checkBox.GetControlFormat();
        REQUIRE(control);
        control.SetLinkedCell("$A$1");
        REQUIRE(control);
        CHECK(control.GetLinkedCell() == "$A$1");

        control.SetValue(xlOn);
        REQUIRE(control);
        CHECK(control.GetValue() == xlOn);
        CHECK(book.GetRange("A1").GetValue().GetBool());

        control.SetValue(xlOff);
        REQUIRE(control);
        CHECK_FALSE(book.GetRange("A1").GetValue().GetBool());
    }

    SECTION("list box with items")
    {
        wxExcelShape listBox = shapes.AddFormControl(xlListBox, 10, 10, 100, 60);
        REQUIRE(listBox);
        wxExcelControlFormat control = listBox.GetControlFormat();
        REQUIRE(control);

        control.AddItem("One", NULL);
        control.AddItem("Two", NULL);
        long index = 1;
        control.AddItem("Zero", &index);
        REQUIRE(control);
        CHECK(control.GetListCount() == 3);
        CHECK(control.List(1) == "Zero");
        CHECK(control.List(3) == "Two");
        REQUIRE(control);

        control.SetListIndex(2);
        REQUIRE(control);
        CHECK(control.GetListIndex() == 2);

        control.RemoveItem(1);
        REQUIRE(control);
        CHECK(control.GetListCount() == 2);
        CHECK(control.List(1) == "One");

        wxArrayString items;
        items.push_back("A");
        items.push_back("B");
        control.SetList(items);
        REQUIRE(control);
        wxArrayString list = control.List();
        REQUIRE(control);
        REQUIRE(list.size() == 2);
        CHECK(list[1] == "B");

        control.RemoveAllItems();
        REQUIRE(control);
        CHECK(control.GetListCount() == 0);
    }

    SECTION("list box filled from a range")
    {
        book.GetRange("A1").SetValue("Red");
        book.GetRange("A2").SetValue("Green");
        book.GetRange("A3").SetValue("Blue");

        wxExcelShape listBox = shapes.AddFormControl(xlListBox, 100, 10, 100, 60);
        REQUIRE(listBox);
        wxExcelControlFormat control = listBox.GetControlFormat();
        REQUIRE(control);
        control.SetListFillRange("$A$1:$A$3");
        REQUIRE(control);
        CHECK(control.GetListFillRange() == "$A$1:$A$3");
        CHECK(control.GetListCount() == 3);
        CHECK(control.List(3) == "Blue");
        REQUIRE(control);
    }

    SECTION("scroll bar")
    {
        wxExcelShape scrollBar = shapes.AddFormControl(xlScrollBar, 10, 10, 20, 100);
        REQUIRE(scrollBar);
        wxExcelControlFormat control = scrollBar.GetControlFormat();
        REQUIRE(control);
        control.SetMin(10);
        control.SetMax(50);
        control.SetSmallChange(2);
        control.SetLargeChange(10);
        control.SetValue(20);
        REQUIRE(control);
        CHECK(control.GetMin() == 10);
        CHECK(control.GetMax() == 50);
        CHECK(control.GetSmallChange() == 2);
        CHECK(control.GetLargeChange() == 10);
        CHECK(control.GetValue() == 20);
        REQUIRE(control);
    }
}

TEST_CASE("wxExcelShapeNodes changes the nodes of a polyline and a freeform",
          "[excel][shapes][nodes]")
{
    ExcelTestWorkbook book;
    REQUIRE(book.IsReady());

    wxExcelShapes shapes = GetShapes(book);

    SECTION("polyline")
    {
        wxVector<wxPoint2DDouble> points;
        points.push_back(wxPoint2DDouble(10, 10));
        points.push_back(wxPoint2DDouble(100, 50));
        points.push_back(wxPoint2DDouble(10, 100));
        wxExcelShape polyline = shapes.AddPolyline(points);
        REQUIRE(polyline);
        CHECK(polyline.GetType() == msoFreeform);
        CHECK(GetNodeCount(polyline) == 3);

        wxExcelShapeNodes nodes = polyline.GetNodes();
        REQUIRE(nodes);
        wxExcelShapeNode second = nodes.Item(2);
        REQUIRE(second);
        CHECK(second.GetSegmentType() == msoSegmentLine);
        wxVector<wxPoint2DDouble> secondPoints = second.GetPoints();
        REQUIRE(second);
        REQUIRE(secondPoints.size() == 1);
        CHECK(secondPoints[0].m_x == Catch::Approx(100.0));
        CHECK(secondPoints[0].m_y == Catch::Approx(50.0));

        nodes.Insert(2, msoSegmentLine, msoEditingAuto, 60, 20);
        REQUIRE(nodes);
        CHECK(GetNodeCount(polyline) == 4);

        // the node is inserted after the given one
        nodes.SetPosition(3, 120, 60);
        REQUIRE(nodes);
        secondPoints = nodes.Item(3).GetPoints();
        REQUIRE(secondPoints.size() == 1);
        CHECK(secondPoints[0].m_x == Catch::Approx(120.0));

        nodes.Delete(2);
        REQUIRE(nodes);
        CHECK(GetNodeCount(polyline) == 3);
    }

    SECTION("freeform")
    {
        wxExcelFreeformBuilder builder = shapes.BuildFreeform(msoEditingAuto, 10, 10);
        REQUIRE(builder);
        builder.AddNodes(msoSegmentLine, msoEditingAuto, 100, 10);
        builder.AddNodes(msoSegmentLine, msoEditingAuto, 100, 100);
        builder.AddNodes(msoSegmentLine, msoEditingAuto, 10, 100);
        REQUIRE(builder);

        wxExcelShape freeform = builder.ConvertToShape();
        REQUIRE(freeform);
        CHECK(freeform.GetType() == msoFreeform);
        CHECK(GetNodeCount(freeform) == 4);
        CHECK(freeform.GetWidth() == Catch::Approx(90.0));
        REQUIRE(freeform);
    }
}

TEST_CASE("wxExcelPictureFormat changes a picture",
          "[excel][shapes][picture]")
{
    ExcelTestWorkbook book;
    REQUIRE(book.IsReady());

    TempPicture file;
    wxExcelShape picture = GetShapes(book).AddPicture(file.GetPath(), msoFalse, msoTrue, 10, 10, 40, 40);
    REQUIRE(picture);
    CHECK(picture.GetType() == msoPicture);
    REQUIRE(picture);

    wxExcelPictureFormat format = picture.GetPictureFormat();
    REQUIRE(format);
    CHECK(format.GetBrightness() == Catch::Approx(0.5));
    REQUIRE(format);

    format.SetBrightness(0.7);
    format.SetContrast(0.6);
    REQUIRE(format);
    CHECK(format.GetBrightness() == Catch::Approx(0.7));
    CHECK(format.GetContrast() == Catch::Approx(0.6));

    format.IncrementBrightness(-0.1);
    format.IncrementContrast(0.1);
    REQUIRE(format);
    CHECK(format.GetBrightness() == Catch::Approx(0.6));
    CHECK(format.GetContrast() == Catch::Approx(0.7));

    format.SetColorType(msoPictureGrayscale);
    REQUIRE(format);
    CHECK(format.GetColorType() == msoPictureGrayscale);
}

TEST_CASE("wxExcelShape effects change shadow, glow, soft edges, and reflection",
          "[excel][shapes][effects]")
{
    ExcelTestWorkbook book;
    REQUIRE(book.IsReady());

    wxExcelShapes shapes = GetShapes(book);
    wxExcelShape shape = AddRectangle(shapes, 10, 10);

    wxExcelShadowFormat shadow = shape.GetShadow();
    REQUIRE(shadow);
    shadow.SetVisible(msoTrue);
    shadow.SetOffsetX(5);
    shadow.SetOffsetY(-3);
    shadow.SetBlur(4);
    shadow.SetTransparency(0.5);
    REQUIRE(shadow);
    CHECK(shadow.GetVisible() == msoTrue);
    CHECK(shadow.GetOffsetX() == Catch::Approx(5.0));
    CHECK(shadow.GetOffsetY() == Catch::Approx(-3.0));
    CHECK(shadow.GetBlur() == Catch::Approx(4.0));
    CHECK(shadow.GetTransparency() == Catch::Approx(0.5));
    REQUIRE(shadow);

    wxExcelGlowFormat glow = shape.GetGlow();
    REQUIRE(glow);
    glow.SetRadius(8);
    REQUIRE(glow);
    wxExcelColorFormat glowColor = glow.GetColor();
    REQUIRE(glowColor);
    glowColor.SetRGB(*wxRED);
    REQUIRE(glowColor);
    CHECK(glow.GetRadius() == Catch::Approx(8.0));
    CHECK(glow.GetColor().GetRGB() == *wxRED);
    REQUIRE(glow);

    wxExcelSoftEdgeFormat softEdge = shape.GetSoftEdge();
    REQUIRE(softEdge);
    softEdge.SetType(msoSoftEdgeType2);
    REQUIRE(softEdge);
    CHECK(softEdge.GetType() == msoSoftEdgeType2);

    wxExcelReflectionFormat reflection = shape.GetReflection();
    REQUIRE(reflection);
    reflection.SetType(msoReflectionType3);
    REQUIRE(reflection);
    CHECK(reflection.GetType() == msoReflectionType3);
}

TEST_CASE("wxExcelTextFrame changes the text layout of a shape",
          "[excel][shapes][text]")
{
    ExcelTestWorkbook book;
    REQUIRE(book.IsReady());

    wxExcelShapes shapes = GetShapes(book);
    wxExcelShape shape = AddRectangle(shapes, 10, 10, 150, 80);
    SetText(shape, "Some text");

    wxExcelTextFrame frame = shape.GetTextFrame();
    REQUIRE(frame);
    CHECK(frame.Characters().GetText() == "Some text");
    long start = 6, length = 4;
    CHECK(frame.Characters(&start, &length).GetText() == "text");
    REQUIRE(frame);

    frame.SetMarginLeft(10);
    frame.SetMarginTop(5);
    frame.SetHorizontalAlignment(xlHAlignRight);
    frame.SetVerticalAlignment(xlVAlignTop);
    REQUIRE(frame);
    CHECK(frame.GetMarginLeft() == Catch::Approx(10.0));
    CHECK(frame.GetMarginTop() == Catch::Approx(5.0));
    CHECK(frame.GetHorizontalAlignment() == xlHAlignRight);
    CHECK(frame.GetVerticalAlignment() == xlVAlignTop);
    REQUIRE(frame);

    frame.SetAutoSize(true);
    REQUIRE(frame);
    CHECK(frame.GetAutoSize());

    wxExcelTextColumn2 column = shape.GetTextFrame2().GetColumn();
    REQUIRE(column);
    column.SetNumber(2);
    column.SetSpacing(12);
    REQUIRE(column);
    CHECK(column.GetNumber() == 2);
    CHECK(column.GetSpacing() == Catch::Approx(12.0));
}

TEST_CASE("wxExcelParagraphFormat2 changes alignment, bullets, and tab stops",
          "[excel][shapes][text]")
{
    ExcelTestWorkbook book;
    REQUIRE(book.IsReady());

    wxExcelShapes shapes = GetShapes(book);
    wxExcelShape shape = AddRectangle(shapes, 10, 10, 150, 80);
    SetText(shape, "First\rSecond");

    wxExcelTextRange2 text = shape.GetTextFrame2().GetTextRange();
    REQUIRE(text);
    wxExcelParagraphFormat2 paragraph = text.GetParagraphFormat();
    REQUIRE(paragraph);

    paragraph.SetAlignment(msoAlignCenter);
    paragraph.SetLeftIndent(10);
    paragraph.SetSpaceAfter(6);
    REQUIRE(paragraph);
    CHECK(paragraph.GetAlignment() == msoAlignCenter);
    CHECK(paragraph.GetLeftIndent() == Catch::Approx(10.0));
    CHECK(paragraph.GetSpaceAfter() == Catch::Approx(6.0));
    REQUIRE(paragraph);

    wxExcelBulletFormat2 bullet = paragraph.GetBullet();
    REQUIRE(bullet);
    bullet.SetVisible(msoTrue);
    bullet.SetType(msoBulletNumbered);
    bullet.SetStartValue(3);
    REQUIRE(bullet);
    CHECK(bullet.GetVisible() == msoTrue);
    CHECK(bullet.GetType() == msoBulletNumbered);
    CHECK(bullet.GetStartValue() == 3);
    REQUIRE(bullet);

    bullet.SetType(msoBulletUnnumbered);
    bullet.SetCharacter(8226);
    bullet.SetRelativeSize(1.5);
    REQUIRE(bullet);
    CHECK(bullet.GetCharacter() == 8226);
    CHECK(bullet.GetRelativeSize() == Catch::Approx(1.5));
    REQUIRE(bullet);

    wxExcelTabStops2 tabStops = paragraph.GetTabStops();
    REQUIRE(tabStops);
    CHECK(tabStops.GetCount() == 0);
    REQUIRE(tabStops.Add(msoTabStopLeft, 20));
    REQUIRE(tabStops.Add(msoTabStopRight, 60));
    CHECK(tabStops.GetCount() == 2);
    REQUIRE(tabStops);

    wxExcelTabStop2 right = tabStops.Item(2);
    REQUIRE(right);
    CHECK(right.GetType() == msoTabStopRight);
    CHECK(right.GetPosition() == Catch::Approx(60.0));
    right.SetPosition(80);
    REQUIRE(right);
    CHECK(tabStops.Item(2).GetPosition() == Catch::Approx(80.0));

    tabStops.Item(1).Clear();
    REQUIRE(tabStops);
    CHECK(tabStops.GetCount() == 1);
}

TEST_CASE("wxExcelShapeRange aligns and distributes shapes",
          "[excel][shapes][range]")
{
    ExcelTestWorkbook book;
    REQUIRE(book.IsReady());

    wxExcelShapes shapes = GetShapes(book);
    wxExcelShape first = AddRectangle(shapes, 10, 10, 50, 50);
    wxExcelShape second = AddRectangle(shapes, 40, 100, 50, 50);
    wxExcelShape third = AddRectangle(shapes, 300, 200, 50, 50);

    wxVector<long> indices;
    indices.push_back(1);
    indices.push_back(2);
    indices.push_back(3);
    wxExcelShapeRange range = shapes.GetRange(indices);
    REQUIRE(range);

    range.Align(msoAlignLefts);
    REQUIRE(range);
    CHECK(first.GetLeft() == Catch::Approx(10.0));
    CHECK(second.GetLeft() == Catch::Approx(10.0));
    CHECK(third.GetLeft() == Catch::Approx(10.0));

    // the space between the shapes becomes the same, the outer ones stay
    range.Distribute(msoDistributeVertically);
    REQUIRE(range);
    CHECK(first.GetTop() == Catch::Approx(10.0));
    CHECK(third.GetTop() == Catch::Approx(200.0));
    CHECK(second.GetTop() == Catch::Approx(105.0));
}

TEST_CASE("wxExcelShape scales and rotates",
          "[excel][shapes]")
{
    ExcelTestWorkbook book;
    REQUIRE(book.IsReady());

    wxExcelShapes shapes = GetShapes(book);
    wxExcelShape shape = AddRectangle(shapes, 10, 10, 100, 50);

    MsoScaleFrom fromTopLeft = msoScaleFromTopLeft;
    shape.ScaleWidth(2, msoFalse, &fromTopLeft);
    shape.ScaleHeight(0.5, msoFalse, &fromTopLeft);
    REQUIRE(shape);
    CHECK(shape.GetWidth() == Catch::Approx(200.0));
    CHECK(shape.GetHeight() == Catch::Approx(25.0));
    CHECK(shape.GetLeft() == Catch::Approx(10.0));
    REQUIRE(shape);

    shape.IncrementRotation(30);
    REQUIRE(shape);
    CHECK(shape.GetRotation() == Catch::Approx(30.0));
    shape.IncrementRotation(-45);
    REQUIRE(shape);
    CHECK(shape.GetRotation() == Catch::Approx(345.0));

    wxExcelThreeDFormat threeD = shape.GetThreeD();
    REQUIRE(threeD);
    threeD.IncrementRotationX(20);
    threeD.IncrementRotationY(10);
    threeD.IncrementRotationZ(5);
    REQUIRE(threeD);
    CHECK(threeD.GetRotationX() == Catch::Approx(20.0));
    CHECK(threeD.GetRotationY() == Catch::Approx(10.0));
    CHECK(threeD.GetRotationZ() == Catch::Approx(5.0));
    REQUIRE(threeD);

    threeD.ResetRotation();
    REQUIRE(threeD);
    CHECK(threeD.GetRotationX() == Catch::Approx(0.0));
    CHECK(threeD.GetRotationY() == Catch::Approx(0.0));
}

#endif // WXAUTOEXCEL_USE_SHAPES
