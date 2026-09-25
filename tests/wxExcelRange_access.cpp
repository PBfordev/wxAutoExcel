/////////////////////////////////////////////////////////////////////////////
// Purpose:     Tests for obtaining wxExcelRange objects
// Author:      OpenAI Codex, under the direction of PB
// Generated:   AI-generated using GPT-5.6 Sol with High reasoning effort
// Copyright:   (c) 2026 PB <pbfordev@gmail.com>
// License:     MIT license
/////////////////////////////////////////////////////////////////////////////

#include <catch2/catch_test_macros.hpp>

#include "excel_test_fixture.h"

using namespace wxAutoExcel;

namespace
{

wxExcelWorksheet GetWorksheet(ExcelTestWorkbook& excel)
{
    wxExcelWorksheets worksheets = excel.GetWorksheets();
    wxExcelWorksheet worksheet = worksheets[1];
    REQUIRE(worksheet);
    return worksheet;
}

void CheckAddress(wxExcelRange range, const wxString& expected)
{
    REQUIRE(range);
    const wxString address = range.GetAddress();
    REQUIRE(range);
    CHECK(address == expected);
}

} // unnamed namespace

TEST_CASE("wxExcelRangeOwner GetRange overloads select expected cells",
          "[excel][range][access][get-range]")
{
    ExcelTestWorkbook excel;
    REQUIRE(excel.IsReady());

    wxExcelWorksheet worksheet = GetWorksheet(excel);

    CheckAddress(worksheet.GetRange("B2"), "$B$2");
    CheckAddress(worksheet.GetRange("B2", "D4"), "$B$2:$D$4");

    wxExcelRange first = worksheet.GetRange("B2");
    REQUIRE(first);
    wxExcelRange last = worksheet.GetRange("D4");
    REQUIRE(last);

    CheckAddress(worksheet.GetRange(first, last), "$B$2:$D$4");
    CheckAddress(worksheet.GetRange(first, "D4"), "$B$2:$D$4");
}

TEST_CASE("wxExcelRange GetRange addresses are relative to their owner",
          "[excel][range][access][get-range][relative]")
{
    ExcelTestWorkbook excel;
    REQUIRE(excel.IsReady());

    wxExcelWorksheet worksheet = GetWorksheet(excel);
    wxExcelRange owner = worksheet.GetRange("F6:J10");
    REQUIRE(owner);

    CheckAddress(owner.GetRange("A1"), "$F$6");
    CheckAddress(owner.GetRange("B2", "D4"), "$G$7:$I$9");

    // Range.Range interprets Range arguments relative to the owner too, so
    // use worksheet A1/C3 as the equivalent of the relative string addresses.
    wxExcelRange first = worksheet.GetRange("A1");
    REQUIRE(first);
    wxExcelRange last = worksheet.GetRange("C3");
    REQUIRE(last);

    CheckAddress(owner.GetRange(first, last), "$F$6:$H$8");
    CheckAddress(owner.GetRange(first, "C3"), "$F$6:$H$8");
}

TEST_CASE("wxExcelRangeOwner GetCells overloads select expected cells",
          "[excel][range][access][get-cells]")
{
    ExcelTestWorkbook excel;
    REQUIRE(excel.IsReady());

    wxExcelWorksheet worksheet = GetWorksheet(excel);
    wxExcelRange owner = worksheet.GetRange("B2:D4");
    REQUIRE(owner);

    CheckAddress(owner.GetCells(), "$B$2:$D$4");

    long row = 2;
    long column = 3;
    CheckAddress(owner.GetCells(&row, &column), "$D$3");

    long itemIndex = 2;
    CheckAddress(owner.GetCells(&itemIndex, NULL), "$C$2");

    long columnOnly = 3;
    CheckAddress(owner.GetCells(NULL, &columnOnly), "$D$2");

    CheckAddress(worksheet.GetCells(4, "C"), "$C$4");
}

TEST_CASE("wxExcelRangeOwner GetRows overloads select expected rows",
          "[excel][range][access][get-rows]")
{
    ExcelTestWorkbook excel;
    REQUIRE(excel.IsReady());

    wxExcelWorksheet worksheet = GetWorksheet(excel);
    wxExcelRange owner = worksheet.GetRange("B2:D5");
    REQUIRE(owner);

    CheckAddress(owner.GetRows(), "$B$2:$D$5");
    CheckAddress(owner.GetRows(2), "$B$3:$D$3");
    CheckAddress(owner.GetRows("1:2"), "$B$2:$D$3");

    CheckAddress(worksheet.GetRows(3), "$3:$3");
    CheckAddress(worksheet.GetRows("3:5"), "$3:$5");
}

TEST_CASE("wxExcelRangeOwner GetColumns overloads select expected columns",
          "[excel][range][access][get-columns]")
{
    ExcelTestWorkbook excel;
    REQUIRE(excel.IsReady());

    wxExcelWorksheet worksheet = GetWorksheet(excel);
    wxExcelRange owner = worksheet.GetRange("B2:E4");
    REQUIRE(owner);

    CheckAddress(owner.GetColumns(), "$B$2:$E$4");
    CheckAddress(owner.GetColumns(2), "$C$2:$C$4");
    CheckAddress(owner.GetColumns("A:B"), "$B$2:$C$4");

    CheckAddress(worksheet.GetColumns(3), "$C:$C");
    CheckAddress(worksheet.GetColumns("C:E"), "$C:$E");
}

TEST_CASE("wxExcelApplication range access uses the active worksheet",
          "[excel][range][access][application]")
{
    ExcelTestWorkbook excel;
    REQUIRE(excel.IsReady());

    wxExcelWorksheet worksheet = GetWorksheet(excel);
    REQUIRE(worksheet.Activate());

    wxExcelApplication& application = wxAutoExcelTests::GetApplication();
    long row = 4;
    long column = 3;
    CheckAddress(application.GetCells(&row, &column), "$C$4");
    CheckAddress(application.GetRange("B2", "D4"), "$B$2:$D$4");
    CheckAddress(application.GetRows(3), "$3:$3");
    CheckAddress(application.GetColumns(3), "$C:$C");
}
