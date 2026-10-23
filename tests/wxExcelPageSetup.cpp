/////////////////////////////////////////////////////////////////////////////
// Purpose:     Tests for wxExcelPageSetup
// Author:      Claude Opus 5.5 (Anthropic), under the direction of PB
// Copyright:   (c) 2026 PB <pbfordev@gmail.com>
// License:     MIT license
/////////////////////////////////////////////////////////////////////////////

#include <catch2/catch_test_macros.hpp>

#include "excel_test_fixture.h"

using namespace wxAutoExcel;

namespace
{

wxExcelPageSetup GetPageSetup(ExcelTestWorkbook& book)
{
    wxExcelPageSetup pageSetup = book.GetWorksheet().GetPageSetup();
    REQUIRE(pageSetup);
    return pageSetup;
}

} // unnamed namespace

TEST_CASE("wxExcelPageSetup sets the orientation, margins, and centering",
          "[excel][pagesetup]")
{
    ExcelTestWorkbook book;
    REQUIRE(book.IsReady());

    wxExcelPageSetup pageSetup = GetPageSetup(book);

    pageSetup.SetOrientation(xlLandscape);
    REQUIRE(pageSetup);
    CHECK(pageSetup.GetOrientation() == xlLandscape);
    REQUIRE(pageSetup);

    // the margins are in points, 36 points is half an inch
    pageSetup.SetLeftMargin(36.0);
    pageSetup.SetRightMargin(36.0);
    pageSetup.SetTopMargin(72.0);
    pageSetup.SetBottomMargin(54.0);
    REQUIRE(pageSetup);
    CHECK(pageSetup.GetLeftMargin() == 36.0);
    CHECK(pageSetup.GetRightMargin() == 36.0);
    CHECK(pageSetup.GetTopMargin() == 72.0);
    CHECK(pageSetup.GetBottomMargin() == 54.0);
    REQUIRE(pageSetup);

    pageSetup.SetCenterHorizontally(true);
    REQUIRE(pageSetup);
    CHECK(pageSetup.GetCenterHorizontally());
    REQUIRE(pageSetup);

    pageSetup.SetOrientation(xlPortrait);
    REQUIRE(pageSetup);
    CHECK(pageSetup.GetOrientation() == xlPortrait);
    REQUIRE(pageSetup);
}

TEST_CASE("wxExcelPageSetup sets the print area and the print titles",
          "[excel][pagesetup]")
{
    ExcelTestWorkbook book;
    REQUIRE(book.IsReady());

    wxExcelPageSetup pageSetup = GetPageSetup(book);
    CHECK(pageSetup.GetPrintArea().empty());
    REQUIRE(pageSetup);

    pageSetup.SetPrintArea("$A$1:$D$20");
    REQUIRE(pageSetup);
    CHECK(pageSetup.GetPrintArea() == "$A$1:$D$20");
    REQUIRE(pageSetup);

    // the print area is stored as the worksheet name Print_Area
    wxExcelNames names = book.GetWorksheet().GetNames();
    REQUIRE(names);
    CHECK(names.GetCount() == 1);
    REQUIRE(names);

    pageSetup.SetPrintTitleRows("$1:$2");
    REQUIRE(pageSetup);
    CHECK(pageSetup.GetPrintTitleRows() == "$1:$2");
    REQUIRE(pageSetup);

    pageSetup.SetPrintTitleColumns("$A:$A");
    REQUIRE(pageSetup);
    CHECK(pageSetup.GetPrintTitleColumns() == "$A:$A");
    REQUIRE(pageSetup);

    pageSetup.SetPrintArea("");
    REQUIRE(pageSetup);
    CHECK(pageSetup.GetPrintArea().empty());
    REQUIRE(pageSetup);
}

TEST_CASE("wxExcelPageSetup sets the headers, footers, and printing options",
          "[excel][pagesetup]")
{
    ExcelTestWorkbook book;
    REQUIRE(book.IsReady());

    wxExcelPageSetup pageSetup = GetPageSetup(book);

    pageSetup.SetLeftHeader("Report");
    pageSetup.SetCenterHeader("&A");
    pageSetup.SetRightFooter("Page &P of &N");
    REQUIRE(pageSetup);
    CHECK(pageSetup.GetLeftHeader() == "Report");
    CHECK(pageSetup.GetCenterHeader() == "&A");
    CHECK(pageSetup.GetRightFooter() == "Page &P of &N");
    REQUIRE(pageSetup);

    pageSetup.SetPrintGridlines(true);
    pageSetup.SetPrintHeadings(true);
    pageSetup.SetBlackAndWhite(true);
    REQUIRE(pageSetup);
    CHECK(pageSetup.GetPrintGridlines());
    CHECK(pageSetup.GetPrintHeadings());
    CHECK(pageSetup.GetBlackAndWhite());
    REQUIRE(pageSetup);

    pageSetup.SetOrder(xlOverThenDown);
    pageSetup.SetFirstPageNumber(5);
    REQUIRE(pageSetup);
    CHECK(pageSetup.GetOrder() == xlOverThenDown);
    CHECK(pageSetup.GetFirstPageNumber() == 5);
    REQUIRE(pageSetup);
}

TEST_CASE("wxExcelPageSetup sets the scaling",
          "[excel][pagesetup]")
{
    ExcelTestWorkbook book;
    REQUIRE(book.IsReady());

    wxExcelPageSetup pageSetup = GetPageSetup(book);

    pageSetup.SetZoom(75);
    REQUIRE(pageSetup);
    CHECK(pageSetup.GetZoom() == 75);
    REQUIRE(pageSetup);

    pageSetup.SetFitToPagesWide(1);
    pageSetup.SetFitToPagesTall(3);
    REQUIRE(pageSetup);
    CHECK(pageSetup.GetFitToPagesWide() == 1);
    CHECK(pageSetup.GetFitToPagesTall() == 3);
    REQUIRE(pageSetup);

    // a zoom out of the 10 to 400 range is rejected
    {
        wxLogNull noLog;
        pageSetup.SetZoom(500);
        CHECK_FALSE(pageSetup);
    }
    CHECK(pageSetup.GetZoom() == 75);
    REQUIRE(pageSetup);
}
