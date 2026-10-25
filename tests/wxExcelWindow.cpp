/////////////////////////////////////////////////////////////////////////////
// Purpose:     Tests for wxExcelWindows, wxExcelWindow, wxExcelPanes,
//              wxExcelPane, wxExcelSheetViews, and wxExcelWorksheetView
// Author:      Claude Opus 5.5 (Anthropic), under the direction of PB
// Copyright:   (c) 2026 PB <pbfordev@gmail.com>
// License:     MIT license
/////////////////////////////////////////////////////////////////////////////

#include <catch2/catch_test_macros.hpp>

#include "excel_test_fixture.h"

using namespace wxAutoExcel;

namespace
{

// Returns the window of the workbook, after activating the workbook and its
// first worksheet, so that the window shows the worksheet.
wxExcelWindow GetWindow(ExcelTestWorkbook& book)
{
    REQUIRE(book.GetWorkbook().Activate());
    REQUIRE(book.GetWorksheet().Activate());

    wxExcelWindows windows = book.GetWorkbook().GetWindows();
    REQUIRE(windows);
    wxExcelWindow window = windows.GetItem(1);
    REQUIRE(window);
    return window;
}

long GetPaneCount(wxExcelWindow& window)
{
    wxExcelPanes panes = window.GetPanes();
    REQUIRE(panes);
    const long count = panes.GetCount();
    REQUIRE(panes);
    return count;
}

// Returns the address of the top-left cell of the range, e.g. "C10".
wxString GetTopLeftCell(wxExcelRange range)
{
    REQUIRE(range);
    const wxString address = range.GetRange("A1").GetAddress(false, false);
    REQUIRE(range);
    return address;
}

} // unnamed namespace

TEST_CASE("wxExcelWindow shows the active sheet and cell of its workbook",
          "[excel][window]")
{
    ExcelTestWorkbook book;
    REQUIRE(book.IsReady());

    wxExcelWindow window = GetWindow(book);

    CHECK(window.GetType() == xlWorkbook);
    CHECK(window.GetIndex() == 1);
    CHECK(window.GetCaption() == book.GetWorkbook().GetName());
    REQUIRE(window);

    wxExcelSheet sheet = window.GetActiveSheet();
    REQUIRE(sheet);
    CHECK(sheet.GetName() == book.GetWorksheet().GetName());

    REQUIRE(book.GetRange("C3").Activate());
    wxExcelRange activeCell = window.GetActiveCell();
    REQUIRE(activeCell);
    CHECK(activeCell.GetAddress(false, false) == "C3");

    CHECK(book.GetWorkbook().GetWindows().GetCount() == 1);
    wxExcelWindow newWindow = window.NewWindow();
    REQUIRE(newWindow);
    CHECK(book.GetWorkbook().GetWindows().GetCount() == 2);
    CHECK(newWindow.GetIndex() == 1);
}

TEST_CASE("wxExcelWindow changes zoom, view, and displayed elements",
          "[excel][window]")
{
    ExcelTestWorkbook book;
    REQUIRE(book.IsReady());

    wxExcelWindow window = GetWindow(book);

    window.SetZoom(150);
    REQUIRE(window);
    CHECK(window.GetZoom() == 150);

    CHECK(window.GetDisplayGridlines());
    window.SetDisplayGridlines(false);
    REQUIRE(window);
    CHECK_FALSE(window.GetDisplayGridlines());

    CHECK(window.GetDisplayHeadings());
    window.SetDisplayHeadings(false);
    REQUIRE(window);
    CHECK_FALSE(window.GetDisplayHeadings());

    window.SetDisplayFormulas(true);
    REQUIRE(window);
    CHECK(window.GetDisplayFormulas());

    window.SetDisplayZeros(false);
    REQUIRE(window);
    CHECK_FALSE(window.GetDisplayZeros());

    CHECK(window.GetView() == xlNormalView);
    window.SetView(xlPageLayoutView);
    REQUIRE(window);
    CHECK(window.GetView() == xlPageLayoutView);
    window.SetView(xlNormalView);
    REQUIRE(window);
    CHECK(window.GetView() == xlNormalView);
}

TEST_CASE("wxExcelWindow splits and freezes its panes",
          "[excel][window][panes]")
{
    ExcelTestWorkbook book;
    REQUIRE(book.IsReady());

    wxExcelWindow window = GetWindow(book);

    CHECK_FALSE(window.GetSplit());
    CHECK(GetPaneCount(window) == 1);

    window.SetSplitRow(2);
    window.SetSplitColumn(1);
    REQUIRE(window);
    CHECK(window.GetSplit());
    CHECK(window.GetSplitRow() == 2);
    CHECK(window.GetSplitColumn() == 1);
    CHECK(GetPaneCount(window) == 4);

    // each pane starts at its own cell
    wxExcelPanes panes = window.GetPanes();
    REQUIRE(panes);
    wxExcelPane bottomRight = panes.GetItem(4);
    REQUIRE(bottomRight);
    CHECK(bottomRight.GetIndex() == 4);
    CHECK(GetTopLeftCell(bottomRight.GetVisibleRange()) == "B3");
    CHECK(GetTopLeftCell(panes.GetItem(1).GetVisibleRange()) == "A1");

    CHECK_FALSE(window.GetFreezePanes());
    window.SetFreezePanes(true);
    REQUIRE(window);
    CHECK(window.GetFreezePanes());
    CHECK(GetPaneCount(window) == 4);

    window.SetFreezePanes(false);
    window.SetSplit(false);
    REQUIRE(window);
    CHECK_FALSE(window.GetFreezePanes());
    CHECK_FALSE(window.GetSplit());
    CHECK(GetPaneCount(window) == 1);
}

TEST_CASE("wxExcelWindow and wxExcelPane scroll",
          "[excel][window][panes]")
{
    ExcelTestWorkbook book;
    REQUIRE(book.IsReady());

    wxExcelWindow window = GetWindow(book);

    window.SetScrollRow(10);
    window.SetScrollColumn(3);
    REQUIRE(window);
    CHECK(window.GetScrollRow() == 10);
    CHECK(window.GetScrollColumn() == 3);
    CHECK(GetTopLeftCell(window.GetVisibleRange()) == "C10");

    long down = 2;
    REQUIRE(window.SmallScroll(&down));
    CHECK(window.GetScrollRow() == 12);

    // the only pane scrolls with the window
    wxExcelPane pane = window.GetActivePane();
    REQUIRE(pane);
    CHECK(pane.GetScrollRow() == 12);
    CHECK(pane.GetScrollColumn() == 3);

    long toLeft = 1;
    REQUIRE(pane.SmallScroll(NULL, NULL, NULL, &toLeft));
    CHECK(pane.GetScrollColumn() == 2);
    CHECK(window.GetScrollColumn() == 2);

    pane.SetScrollRow(1);
    pane.SetScrollColumn(1);
    REQUIRE(pane);
    CHECK(GetTopLeftCell(pane.GetVisibleRange()) == "A1");
}

TEST_CASE("wxExcelWorksheetView changes how the window shows its worksheet",
          "[excel][window][sheetviews]")
{
    ExcelTestWorkbook book;
    REQUIRE(book.IsReady());

    wxExcelWindow window = GetWindow(book);

    wxExcelSheetViews sheetViews = window.GetSheetViews();
    REQUIRE(sheetViews);
    CHECK(sheetViews.GetCount() == 1);
    REQUIRE(sheetViews);

    wxExcelSheetView sheetView = sheetViews.GetItem(1);
    REQUIRE(sheetView);
    CHECK(sheetView.IsView(wxExcelSheetView::WorksheetView));
    CHECK_FALSE(sheetView.IsView(wxExcelSheetView::ChartView));

    wxExcelWorksheetView worksheetView = sheetView.ToWorksheetView();
    REQUIRE(worksheetView);

    wxExcelWorksheet sheet = worksheetView.GetSheet();
    REQUIRE(sheet);
    CHECK(sheet.GetName() == book.GetWorksheet().GetName());

    worksheetView.SetDisplayGridlines(false);
    REQUIRE(worksheetView);
    CHECK_FALSE(worksheetView.GetDisplayGridlines());
    CHECK_FALSE(window.GetDisplayGridlines());

    worksheetView.SetDisplayZeros(false);
    REQUIRE(worksheetView);
    CHECK_FALSE(window.GetDisplayZeros());
}
