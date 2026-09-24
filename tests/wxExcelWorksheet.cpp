/////////////////////////////////////////////////////////////////////////////
// Purpose:     Tests for wxExcelWorksheet collection operations
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

long GetCount(wxExcelWorksheets& worksheets)
{
    const long count = worksheets.GetCount();
    REQUIRE(worksheets);
    return count;
}

wxExcelWorksheet GetSheet(wxExcelWorksheets& worksheets, long index)
{
    wxExcelWorksheet worksheet = worksheets[index];
    REQUIRE(worksheet);
    return worksheet;
}

void CheckSheetAt(wxExcelWorksheets& worksheets, long index,
                  const wxString& expectedName)
{
    wxExcelWorksheet worksheet = GetSheet(worksheets, index);
    const wxString name = worksheet.GetName();
    REQUIRE(worksheet);
    CHECK(name == expectedName);
}

wxExcelWorksheet AddNamed(wxExcelWorksheets& worksheets,
                          wxExcelWorksheet relativeTo, bool after,
                          const wxString& name)
{
    wxExcelWorksheet worksheet =
        worksheets.AddAfterOrBefore(relativeTo, after);
    REQUIRE(worksheet);

    worksheet.SetName(name);
    REQUIRE(worksheet);
    return worksheet;
}

long GetIndex(wxExcelWorksheet& worksheet)
{
    const long index = worksheet.GetIndex();
    REQUIRE(worksheet);
    return index;
}

} // unnamed namespace

TEST_CASE("wxExcelWorksheets adds sheets at requested positions",
          "[excel][worksheet][add]")
{
    ExcelTestWorkbook excel;
    REQUIRE(excel.IsReady());

    wxExcelWorksheets worksheets = excel.GetWorksheets();
    const long initialCount = GetCount(worksheets);
    REQUIRE(initialCount >= 1);

    wxExcelWorksheet originalFirst = GetSheet(worksheets, 1);
    wxExcelWorksheet beforeFirst = AddNamed(
        worksheets, originalFirst, false, "wxAE Before First");
    wxExcelWorksheet afterFirst = AddNamed(
        worksheets, originalFirst, true, "wxAE After First");

    wxExcelWorksheet currentLast = GetSheet(worksheets, GetCount(worksheets));
    wxExcelWorksheet atEnd = AddNamed(
        worksheets, currentLast, true, "wxAE At End");

    CHECK(GetCount(worksheets) == initialCount + 3);
    CHECK(GetIndex(beforeFirst) == 1);
    CHECK(GetIndex(originalFirst) == 2);
    CHECK(GetIndex(afterFirst) == 3);
    CHECK(GetIndex(atEnd) == GetCount(worksheets));

    CheckSheetAt(worksheets, 1, "wxAE Before First");
    CheckSheetAt(worksheets, 3, "wxAE After First");
    CheckSheetAt(worksheets, GetCount(worksheets), "wxAE At End");
}

TEST_CASE("wxExcelWorksheet moves to the beginning, middle, and end",
          "[excel][worksheet][move]")
{
    ExcelTestWorkbook excel;
    REQUIRE(excel.IsReady());

    wxExcelWorksheets worksheets = excel.GetWorksheets();
    const long originalCount = GetCount(worksheets);
    REQUIRE(originalCount >= 1);

    wxExcelWorksheet originalLast = GetSheet(worksheets, originalCount);
    wxExcelWorksheet sheetA = AddNamed(
        worksheets, originalLast, true, "wxAE Move A");
    wxExcelWorksheet sheetB = AddNamed(
        worksheets, sheetA, true, "wxAE Move B");
    wxExcelWorksheet moving = AddNamed(
        worksheets, sheetB, true, "wxAE Moving");

    wxExcelWorksheet first = GetSheet(worksheets, 1);
    CHECK(moving.MoveAfterOrBefore(first, false));
    CHECK(GetIndex(moving) == 1);
    CheckSheetAt(worksheets, 1, "wxAE Moving");

    CHECK(moving.MoveAfterOrBefore(sheetA, true));
    CHECK(GetIndex(moving) == GetIndex(sheetA) + 1);
    CHECK(GetIndex(sheetB) == GetIndex(moving) + 1);
    CheckSheetAt(worksheets, GetIndex(moving), "wxAE Moving");

    wxExcelWorksheet last = GetSheet(worksheets, GetCount(worksheets));
    CHECK(moving.MoveAfterOrBefore(last, true));
    CHECK(GetIndex(moving) == GetCount(worksheets));
    CheckSheetAt(worksheets, GetCount(worksheets), "wxAE Moving");
}

TEST_CASE("wxExcelWorksheet can be renamed and retrieved by its new name",
          "[excel][worksheet][name]")
{
    ExcelTestWorkbook excel;
    REQUIRE(excel.IsReady());

    wxExcelWorksheets worksheets = excel.GetWorksheets();
    wxExcelWorksheet last = GetSheet(worksheets, GetCount(worksheets));
    wxExcelWorksheet worksheet = AddNamed(
        worksheets, last, true, "wxAE Original Name");
    const long index = GetIndex(worksheet);

    worksheet.SetName("wxAE Renamed Sheet");
    REQUIRE(worksheet);

    const wxString name = worksheet.GetName();
    REQUIRE(worksheet);
    CHECK(name == "wxAE Renamed Sheet");
    CheckSheetAt(worksheets, index, "wxAE Renamed Sheet");

    wxExcelWorksheet byName =
        worksheets.GetItem(wxString("wxAE Renamed Sheet"));
    REQUIRE(byName);
    CHECK(GetIndex(byName) == index);
}

TEST_CASE("wxExcelWorksheet removes sheets from different positions",
          "[excel][worksheet][delete]")
{
    ExcelTestWorkbook excel;
    REQUIRE(excel.IsReady());

    wxExcelWorksheets worksheets = excel.GetWorksheets();
    const long initialCount = GetCount(worksheets);
    REQUIRE(initialCount >= 1);

    wxExcelWorksheet firstAnchor = GetSheet(worksheets, 1);
    wxExcelWorksheet removeFirst = AddNamed(
        worksheets, firstAnchor, false, "wxAE Delete First");

    wxExcelWorksheet currentLast = GetSheet(worksheets, GetCount(worksheets));
    wxExcelWorksheet removeMiddle = AddNamed(
        worksheets, currentLast, true, "wxAE Delete Middle");
    wxExcelWorksheet removeLast = AddNamed(
        worksheets, removeMiddle, true, "wxAE Delete Last");

    CHECK(GetCount(worksheets) == initialCount + 3);

    CHECK(removeMiddle.Delete());
    CHECK(GetCount(worksheets) == initialCount + 2);
    CheckSheetAt(worksheets, GetCount(worksheets), "wxAE Delete Last");

    CHECK(removeFirst.Delete());
    CHECK(GetCount(worksheets) == initialCount + 1);

    CHECK(removeLast.Delete());
    CHECK(GetCount(worksheets) == initialCount);
    CheckSheetAt(worksheets, 1, firstAnchor.GetName());
    REQUIRE(firstAnchor);
}
