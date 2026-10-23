/////////////////////////////////////////////////////////////////////////////
// Purpose:     Tests for the setters of read-write properties which had
//              only a getter, setting values different from the current ones
// Author:      Claude Opus 5.5 (Anthropic), under the direction of PB
// Copyright:   (c) 2026 PB <pbfordev@gmail.com>
// License:     MIT license
/////////////////////////////////////////////////////////////////////////////

#include <catch2/catch_test_macros.hpp>

#include "excel_test_fixture.h"

using namespace wxAutoExcel;

TEST_CASE("wxExcelApplication sets the cursor movement",
          "[excel][setters][application]")
{
    wxExcelApplication& application = wxAutoExcelTests::GetApplication();
    REQUIRE(application);

    const long original = application.GetCursorMovement();
    REQUIRE(application);
    const long other = original == xlLogicalCursor ? xlVisualCursor : xlLogicalCursor;

    application.SetCursorMovement(other);
    REQUIRE(application);
    CHECK(application.GetCursorMovement() == other);

    application.SetCursorMovement(original);
    REQUIRE(application);
}

TEST_CASE("wxExcelRange sets the reading order and shows the detail of an outline",
          "[excel][setters][range]")
{
    ExcelTestWorkbook book;
    REQUIRE(book.IsReady());

    wxExcelRange cell = book.GetRange("A1");
    cell.SetReadingOrder(xlRTL);
    REQUIRE(cell);
    CHECK(cell.GetReadingOrder() == xlRTL);
    REQUIRE(cell);

    // rows 2 to 4 are the detail of the summary row 5
    wxExcelRange detail = book.GetRange("A2:A4").GetEntireRow();
    REQUIRE(detail);
    REQUIRE(detail.Group());

    wxExcelRange summary = book.GetRange("A5").GetEntireRow();
    REQUIRE(summary);
    summary.SetShowDetail(false);
    REQUIRE(summary);
    CHECK_FALSE(summary.GetShowDetail());
    REQUIRE(summary);
    CHECK(book.GetRange("A3").GetEntireRow().GetHidden());

    summary.SetShowDetail(true);
    REQUIRE(summary);
    CHECK(summary.GetShowDetail());
    REQUIRE(summary);
    CHECK_FALSE(book.GetRange("A3").GetEntireRow().GetHidden());

    REQUIRE(detail.Ungroup());
}

TEST_CASE("wxExcelWorksheet sets right to left display and what can be selected",
          "[excel][setters][worksheet]")
{
    ExcelTestWorkbook book;
    REQUIRE(book.IsReady());

    wxExcelWorksheet& worksheet = book.GetWorksheet();

    worksheet.SetDisplayRightToLeft(true);
    REQUIRE(worksheet);
    CHECK(worksheet.GetDisplayRightToLeft());
    REQUIRE(worksheet);
    worksheet.SetDisplayRightToLeft(false);
    REQUIRE(worksheet);
    CHECK_FALSE(worksheet.GetDisplayRightToLeft());
    REQUIRE(worksheet);

    worksheet.SetEnableSelection(xlNoSelection);
    REQUIRE(worksheet);
    CHECK(worksheet.GetEnableSelection() == xlNoSelection);
    REQUIRE(worksheet);
}

TEST_CASE("wxExcelSort sets the sort method",
          "[excel][setters][sort]")
{
    ExcelTestWorkbook book;
    REQUIRE(book.IsReady());

    wxExcelSort sort = book.GetWorksheet().GetSort();
    REQUIRE(sort);
    sort.SetSortMethod(xlStroke);
    REQUIRE(sort);
    CHECK(sort.GetSortMethod() == xlStroke);
    REQUIRE(sort);
}

#if WXAUTOEXCEL_USE_CHARTS

TEST_CASE("wxExcelPlotArea sets its inside position and size",
          "[excel][setters][chart]")
{
    ExcelTestWorkbook book;
    REQUIRE(book.IsReady());

    wxExcelRange data = book.GetRange("A1:A4");
    data.SetValue(wxVariant(1.0));
    REQUIRE(data);

    wxExcelChartObject chartObject = book.GetWorksheet().ChartObjects().Add(100, 10, 400, 300);
    REQUIRE(chartObject);
    wxExcelChart chart = chartObject.GetChart();
    REQUIRE(chart);
    chart.SetSourceData(data);
    chart.SetChartType(xlColumnClustered);
    REQUIRE(chart);

    // the size first, so that the plot area fits in the chart when it moves
    wxExcelPlotArea plotArea = chart.GetPlotArea();
    REQUIRE(plotArea);
    // getting the size makes Excel lay out the new chart, setting it fails before
    CHECK(plotArea.GetInsideWidth() > 150.0);
    REQUIRE(plotArea);
    plotArea.SetInsideWidth(150.0);
    plotArea.SetInsideHeight(120.0);
    plotArea.SetInsideLeft(50.0);
    plotArea.SetInsideTop(40.0);
    REQUIRE(plotArea);
    CHECK(plotArea.GetInsideWidth() == 150.0);
    CHECK(plotArea.GetInsideHeight() == 120.0);
    CHECK(plotArea.GetInsideLeft() == 50.0);
    CHECK(plotArea.GetInsideTop() == 40.0);
    REQUIRE(plotArea);
}

TEST_CASE("wxExcelChartGroup sets the hole size of a doughnut chart",
          "[excel][setters][chart]")
{
    ExcelTestWorkbook book;
    REQUIRE(book.IsReady());

    wxExcelRange data = book.GetRange("A1:A4");
    data.SetValue(wxVariant(1.0));
    REQUIRE(data);

    wxExcelChartObject chartObject = book.GetWorksheet().ChartObjects().Add(100, 10, 400, 300);
    REQUIRE(chartObject);
    wxExcelChart chart = chartObject.GetChart();
    REQUIRE(chart);
    chart.SetSourceData(data);
    chart.SetChartType(xlDoughnut);
    REQUIRE(chart);

    wxExcelChartGroup chartGroup = chart.ChartGroups(1);
    REQUIRE(chartGroup);
    chartGroup.SetDoughnutHoleSize(25);
    REQUIRE(chartGroup);
    CHECK(chartGroup.GetDoughnutHoleSize() == 25);
    REQUIRE(chartGroup);
}

TEST_CASE("wxExcelSparkHorizontalAxis sets the plot order",
          "[excel][setters][sparkline]")
{
    ExcelTestWorkbook book;
    REQUIRE(book.IsReady());

    wxExcelRange data = book.GetRange("A1:D1");
    data.SetValue(wxVariant(2.0));
    REQUIRE(data);

    wxExcelSparklineGroups groups = book.GetRange("E1").GetSparklineGroups();
    REQUIRE(groups);
    wxExcelSparklineGroup group = groups.Add(xlSparkLine, "A1:D1");
    REQUIRE(group);

    wxExcelSparkHorizontalAxis axis = group.GetAxes().GetHorizontal();
    REQUIRE(axis);
    axis.SetRightToLeftPlotOrder(true);
    REQUIRE(axis);
    CHECK(axis.GetRightToLeftPlotOrder());
    REQUIRE(axis);
}

#endif // WXAUTOEXCEL_USE_CHARTS

#if WXAUTOEXCEL_USE_SHAPES

TEST_CASE("wxExcelShapeRange sets its position, size, rotation, and visibility",
          "[excel][setters][shaperange]")
{
    ExcelTestWorkbook book;
    REQUIRE(book.IsReady());

    wxExcelShapes shapes = book.GetWorksheet().GetShapes();
    REQUIRE(shapes);
    wxExcelShape shape = shapes.AddShape(msoShapeRectangle, 10, 20, 100, 50);
    REQUIRE(shape);
    const wxString name = shape.GetName();
    REQUIRE(shape);

    wxExcelShapeRange range = shapes.GetRange(name);
    REQUIRE(range);

    range.SetTop(30.0);
    range.SetWidth(80.0);
    range.SetHeight(40.0);
    range.SetRotation(90.0);
    REQUIRE(range);
    CHECK(range.GetTop() == 30.0);
    CHECK(range.GetWidth() == 80.0);
    CHECK(range.GetHeight() == 40.0);
    CHECK(range.GetRotation() == 90.0);
    REQUIRE(range);

    range.SetAlternativeText("A box");
    range.SetBlackWhiteMode(msoBlackWhiteGrayScale);
    REQUIRE(range);
    CHECK(range.GetAlternativeText() == "A box");
    CHECK(range.GetBlackWhiteMode() == msoBlackWhiteGrayScale);
    REQUIRE(range);

    range.SetVisible(msoFalse);
    REQUIRE(range);
    CHECK(range.GetVisible() == msoFalse);
    CHECK(shape.GetVisible() == msoFalse);
    REQUIRE(range);
}

TEST_CASE("wxExcelShape sets its adjustments, background style, and fill",
          "[excel][setters][shape]")
{
    ExcelTestWorkbook book;
    REQUIRE(book.IsReady());

    wxExcelShapes shapes = book.GetWorksheet().GetShapes();
    REQUIRE(shapes);
    wxExcelShape shape = shapes.AddShape(msoShapeRoundedRectangle, 10, 20, 100, 50);
    REQUIRE(shape);

    wxExcelAdjustments adjustments = shape.GetAdjustments();
    REQUIRE(adjustments);
    adjustments.SetItem(1, 0.25);
    REQUIRE(adjustments);
    CHECK(adjustments.GetItem(1) == 0.25);
    REQUIRE(adjustments);

    shape.SetBackgroundStyle(msoBackgroundStyle3);
    REQUIRE(shape);
    CHECK(shape.GetBackgroundStyle() == msoBackgroundStyle3);
    REQUIRE(shape);

    wxExcelFillFormat fill = shape.GetFill();
    REQUIRE(fill);
    wxExcelColorFormat foreColor = fill.GetForeColor();
    REQUIRE(foreColor);
    foreColor.SetBrightness(0.5);
    REQUIRE(foreColor);
    CHECK(foreColor.GetBrightness() == 0.5);
    REQUIRE(foreColor);

    fill.PresetTextured(msoTextureCanvas);
    REQUIRE(fill);
    fill.SetTextureAlignment(msoTextureCenter);
    REQUIRE(fill);
    CHECK(fill.GetTextureAlignment() == msoTextureCenter);
    REQUIRE(fill);
}

TEST_CASE("wxExcelTextFrame2 and wxExcelParagraphFormat2 set the size and alignment",
          "[excel][setters][shape][text]")
{
    ExcelTestWorkbook book;
    REQUIRE(book.IsReady());

    wxExcelShapes shapes = book.GetWorksheet().GetShapes();
    REQUIRE(shapes);
    wxExcelShape shape = shapes.AddShape(msoShapeRectangle, 10, 20, 100, 50);
    REQUIRE(shape);

    wxExcelTextFrame2 textFrame = shape.GetTextFrame2();
    REQUIRE(textFrame);
    wxExcelTextRange2 textRange = textFrame.GetTextRange();
    REQUIRE(textRange);
    textRange.SetText("Some text");
    REQUIRE(textRange);

    textFrame.SetAutoSize(msoAutoSizeShapeToFitText);
    REQUIRE(textFrame);
    CHECK(textFrame.GetAutoSize() == msoAutoSizeShapeToFitText);
    REQUIRE(textFrame);

    wxExcelParagraphFormat2 paragraphFormat = textRange.GetParagraphFormat();
    REQUIRE(paragraphFormat);
    paragraphFormat.SetAlignment(msoAlignCenter);
    REQUIRE(paragraphFormat);
    CHECK(paragraphFormat.GetAlignment() == msoAlignCenter);
    REQUIRE(paragraphFormat);
}

#endif // WXAUTOEXCEL_USE_SHAPES
