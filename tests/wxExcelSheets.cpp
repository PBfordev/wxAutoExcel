/////////////////////////////////////////////////////////////////////////////
// Purpose:     Tests for copying worksheets and for wxExcelSheets, which
//              contains worksheets and chart sheets
// Author:      Claude Opus 5.5 (Anthropic), under the direction of PB
// Copyright:   (c) 2026 PB <pbfordev@gmail.com>
// License:     MIT license
/////////////////////////////////////////////////////////////////////////////

#include <catch2/catch_test_macros.hpp>

#include "excel_test_fixture.h"

using namespace wxAutoExcel;

namespace
{

long GetCount(wxExcelSheets& sheets)
{
    const long count = sheets.GetCount();
    REQUIRE(sheets);
    return count;
}

long GetCount(wxExcelWorksheets& worksheets)
{
    const long count = worksheets.GetCount();
    REQUIRE(worksheets);
    return count;
}

wxString GetName(wxExcelSheet sheet)
{
    REQUIRE(sheet);
    const wxString name = sheet.GetName();
    REQUIRE(sheet);
    return name;
}

wxExcelSheets GetSheets(ExcelTestWorkbook& book)
{
    wxExcelSheets sheets = book.GetWorkbook().GetSheets();
    REQUIRE(sheets);
    return sheets;
}

long GetWorkbookCount()
{
    wxExcelWorkbooks workbooks = wxAutoExcelTests::GetApplication().GetWorkbooks();
    REQUIRE(workbooks);
    const long count = workbooks.GetCount();
    REQUIRE(workbooks);
    return count;
}

} // unnamed namespace

TEST_CASE("wxExcelWorksheet copies itself with its contents within the workbook",
          "[excel][worksheet][copy]")
{
    ExcelTestWorkbook book;
    REQUIRE(book.IsReady());

    wxExcelWorksheet& original = book.GetWorksheet();
    original.SetName("Original");
    REQUIRE(original);

    wxExcelRange cell = book.GetRange("B2");
    cell.SetValue("copied value");
    REQUIRE(cell);

    wxExcelWorksheets worksheets = book.GetWorksheets();
    const long initialCount = GetCount(worksheets);

    CHECK(original.CopyAfterOrBefore(original, true));
    CHECK(GetCount(worksheets) == initialCount + 1);

    wxExcelWorksheet copy = worksheets[original.GetIndex() + 1];
    REQUIRE(copy);
    CHECK(copy.GetName() == "Original (2)");
    REQUIRE(copy);

    wxExcelRange copiedCell = copy.GetRange("B2");
    REQUIRE(copiedCell);
    const wxVariant value = copiedCell.GetValue();
    REQUIRE(copiedCell);
    REQUIRE(value.IsType("string"));
    CHECK(value.GetString() == "copied value");

    // the copy is independent of the original
    copiedCell.SetValue("changed");
    REQUIRE(copiedCell);
    const wxVariant originalValue = cell.GetValue();
    REQUIRE(cell);
    CHECK(originalValue.GetString() == "copied value");

    CHECK(original.CopyAfterOrBefore(original, false));
    CHECK(GetCount(worksheets) == initialCount + 2);
    CHECK(GetName(worksheets[1L]) == "Original (3)");
}

TEST_CASE("wxExcelWorksheet copied and moved without a position goes to a new workbook",
          "[excel][worksheet][copy]")
{
    ExcelTestWorkbook book;
    REQUIRE(book.IsReady());

    wxExcelApplication& application = wxAutoExcelTests::GetApplication();
    const long initialWorkbookCount = GetWorkbookCount();

    wxExcelWorksheets worksheets = book.GetWorksheets();
    wxExcelWorksheet moved = worksheets.Add();
    REQUIRE(moved);
    moved.SetName("Moved");
    REQUIRE(moved);
    const long initialCount = GetCount(worksheets);

    wxExcelWorksheet& worksheet = book.GetWorksheet();
    worksheet.SetName("Copied");
    REQUIRE(worksheet);
    CHECK(worksheet.Copy());
    CHECK(GetWorkbookCount() == initialWorkbookCount + 1);
    CHECK(GetCount(worksheets) == initialCount);

    wxExcelWorkbook copyBook = application.GetActiveWorkbook();
    REQUIRE(copyBook);
    CHECK(GetName(copyBook.GetSheets()[1L]) == "Copied");
    copyBook.Close(false);
    REQUIRE(copyBook);
    CHECK(GetWorkbookCount() == initialWorkbookCount);

    CHECK(moved.Move());
    CHECK(GetWorkbookCount() == initialWorkbookCount + 1);
    CHECK(GetCount(worksheets) == initialCount - 1);

    wxExcelWorkbook moveBook = application.GetActiveWorkbook();
    REQUIRE(moveBook);
    CHECK(GetName(moveBook.GetSheets()[1L]) == "Moved");
    moveBook.Close(false);
    REQUIRE(moveBook);
}

TEST_CASE("wxExcelSheets adds, moves, and deletes worksheets",
          "[excel][sheets]")
{
    ExcelTestWorkbook book;
    REQUIRE(book.IsReady());

    wxExcelSheets sheets = GetSheets(book);
    const long initialCount = GetCount(sheets);
    wxExcelWorksheets worksheets = book.GetWorksheets();
    CHECK(GetCount(worksheets) == initialCount);

    wxExcelSheet last = sheets[initialCount];
    REQUIRE(last);
    wxExcelSheet added = sheets.AddAfterOrBefore(last, true, 2);
    REQUIRE(added);
    CHECK(GetCount(sheets) == initialCount + 2);
    CHECK(added.IsWorksheet());
    REQUIRE(added);
    CHECK(added.GetType() == xlWorksheet);
    REQUIRE(added);

    wxExcelWorksheet asWorksheet = added.ToWorksheet();
    REQUIRE(asWorksheet);
    asWorksheet.SetName("Added");
    REQUIRE(asWorksheet);

    wxExcelSheet byName = sheets.GetItem("Added");
    REQUIRE(byName);
    CHECK(GetName(byName) == "Added");

    wxExcelSheet first = sheets[1L];
    REQUIRE(first);
    CHECK(byName.MoveAfterOrBefore(first, false));
    CHECK(GetName(sheets[1L]) == "Added");

    CHECK(byName.CopyAfterOrBefore(sheets[GetCount(sheets)], true));
    CHECK(GetCount(sheets) == initialCount + 3);
    CHECK(GetName(sheets[GetCount(sheets)]) == "Added (2)");

    CHECK(byName.Delete());
    CHECK(GetCount(sheets) == initialCount + 2);
    CHECK(GetCount(worksheets) == initialCount + 2);
}

#if WXAUTOEXCEL_USE_CHARTS

TEST_CASE("wxExcelSheets contains chart sheets, which wxExcelWorksheets does not",
          "[excel][sheets][chart]")
{
    ExcelTestWorkbook book;
    REQUIRE(book.IsReady());

    wxExcelRange data = book.GetRange("A1:B3");
    data.SetValue(wxVariant(1.0));
    REQUIRE(data);

    wxExcelSheets sheets = GetSheets(book);
    wxExcelWorksheets worksheets = book.GetWorksheets();
    const long initialCount = GetCount(sheets);

    wxExcelSheet chartSheet = sheets.Add(1, xlChart);
    REQUIRE(chartSheet);
    CHECK(GetCount(sheets) == initialCount + 1);
    CHECK(GetCount(worksheets) == initialCount);

    CHECK(chartSheet.IsChart());
    REQUIRE(chartSheet);
    CHECK_FALSE(chartSheet.IsWorksheet());
    REQUIRE(chartSheet);

    wxExcelChart chart = chartSheet.ToChart();
    REQUIRE(chart);
    chart.SetSourceData(data);
    REQUIRE(chart);

    wxExcelCharts charts = book.GetWorkbook().GetCharts();
    REQUIRE(charts);
    CHECK(charts.GetCount() == 1);
    REQUIRE(charts);

    // a chart sheet is not a worksheet
    CHECK_FALSE(chartSheet.ToWorksheet());

    CHECK(chartSheet.Delete());
    CHECK(GetCount(sheets) == initialCount);
    CHECK(charts.GetCount() == 0);
    REQUIRE(charts);
}

#endif // WXAUTOEXCEL_USE_CHARTS
