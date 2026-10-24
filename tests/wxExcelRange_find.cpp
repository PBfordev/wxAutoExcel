/////////////////////////////////////////////////////////////////////////////
// Purpose:     Tests for finding and replacing in wxExcelRange
// Author:      Claude Opus 5.5 (Anthropic), under the direction of PB
// Copyright:   (c) 2026 PB <pbfordev@gmail.com>
// License:     MIT license
/////////////////////////////////////////////////////////////////////////////

#include <catch2/catch_test_macros.hpp>

#include "excel_test_fixture.h"

using namespace wxAutoExcel;

namespace
{

// Fills B2:D4 with
//   apple   Banana  cherry
//   10      apple   APPLE
//   pineapple 20    apple pie
wxExcelRange FillTable(ExcelTestWorkbook& book)
{
    wxExcelRangeData data(3, 3);

    data(0, 0) = "apple";
    data(0, 1) = "Banana";
    data(0, 2) = "cherry";
    data(1, 0) = 10.0;
    data(1, 1) = "apple";
    data(1, 2) = "APPLE";
    data(2, 0) = "pineapple";
    data(2, 1) = 20.0;
    data(2, 2) = "apple pie";

    wxExcelRange range = book.GetRange("B2:D4");
    REQUIRE(range.SetValues_(data));
    return range;
}

wxString GetAddress(wxExcelRange& range)
{
    REQUIRE(range.IsOk_());
    const wxString address = range.GetAddress(false, false);
    REQUIRE(range);
    return address;
}

wxString JoinAddresses(wxExcelRangeVector& ranges)
{
    wxString result;

    for ( size_t i = 0; i < ranges.size(); ++i )
    {
        if ( i > 0 )
            result += " ";
        result += GetAddress(ranges[i]);
    }
    return result;
}

wxString GetText(ExcelTestWorkbook& book, const wxString& address)
{
    wxExcelRange range = book.GetRange(address);
    const wxString text = range.GetText();
    REQUIRE(range);
    return text;
}

} // unnamed namespace

TEST_CASE("wxExcelRange::Find finds the first matching cell",
          "[excel][range][find]")
{
    ExcelTestWorkbook book;
    REQUIRE(book.IsReady());
    wxExcelRange range = FillTable(book);

    XlLookAt whole = xlWhole;
    XlLookAt part = xlPart;
    XlFindLookIn values = xlValues;

    SECTION("the upper-left cell is searched last")
    {
        wxExcelRange found = range.Find("apple", NULL, &values, &whole,
                                        NULL, NULL, false);
        REQUIRE(range);
        CHECK(GetAddress(found) == "C3");
    }

    SECTION("the search starts after the given cell")
    {
        wxExcelRange after = book.GetRange("C3");
        wxExcelRange found = range.Find("apple", &after, &values, &whole,
                                        NULL, NULL, false);
        REQUIRE(range);
        CHECK(GetAddress(found) == "D3");
    }

    SECTION("the case is matched")
    {
        wxExcelRange found = range.Find("APPLE", NULL, &values, &whole,
                                        NULL, NULL, true);
        REQUIRE(range);
        CHECK(GetAddress(found) == "D3");
    }

    SECTION("a part is matched")
    {
        wxExcelRange found = range.Find("pie", NULL, &values, &part,
                                        NULL, NULL, false);
        REQUIRE(range);
        CHECK(GetAddress(found) == "D4");
    }

    SECTION("a number is found")
    {
        wxExcelRange found = range.Find(20.0, NULL, &values, &whole);
        REQUIRE(range);
        CHECK(GetAddress(found) == "C4");
    }

    SECTION("searching by columns")
    {
        XlSearchOrder byColumns = xlByColumns;
        wxExcelRange found = range.Find("apple", NULL, &values, &part,
                                        &byColumns, NULL, false);
        REQUIRE(range);
        CHECK(GetAddress(found) == "B4");
    }

    SECTION("searching backwards")
    {
        XlSearchDirection previous = xlPrevious;
        wxExcelRange found = range.Find("apple", NULL, &values, &part,
                                        NULL, &previous, false);
        REQUIRE(range);
        CHECK(GetAddress(found) == "D4");
    }

    SECTION("nothing is found without an error")
    {
        wxExcelRange found = range.Find("durian", NULL, &values, &whole);
        CHECK(range);
        CHECK_FALSE(found.IsOk_());
    }
}

TEST_CASE("wxExcelRange::FindNext and FindPrevious continue the search",
          "[excel][range][find]")
{
    ExcelTestWorkbook book;
    REQUIRE(book.IsReady());
    wxExcelRange range = FillTable(book);

    XlLookAt whole = xlWhole;
    XlFindLookIn values = xlValues;

    wxExcelRange first = range.Find("apple", NULL, &values, &whole,
                                    NULL, NULL, false);
    REQUIRE(range);
    REQUIRE(GetAddress(first) == "C3");

    wxExcelRange next = range.FindNext(&first);
    REQUIRE(range);
    CHECK(GetAddress(next) == "D3");

    // the search wraps around to the upper-left cell
    wxExcelRange wrapped = range.FindNext(&next);
    REQUIRE(range);
    CHECK(GetAddress(wrapped) == "B2");

    wxExcelRange again = range.FindNext(&wrapped);
    REQUIRE(range);
    CHECK(GetAddress(again) == "C3");

    wxExcelRange previous = range.FindPrevious(&again);
    REQUIRE(range);
    CHECK(GetAddress(previous) == "B2");
}

TEST_CASE("wxExcelRange::FindAll_ finds all the matching cells",
          "[excel][range][find]")
{
    ExcelTestWorkbook book;
    REQUIRE(book.IsReady());
    wxExcelRange range = FillTable(book);

    XlLookAt whole = xlWhole;
    XlLookAt part = xlPart;
    XlFindLookIn values = xlValues;
    wxExcelRangeVector found;

    SECTION("by rows, starting with the upper-left cell")
    {
        REQUIRE(range.FindAll_("apple", found, &values, &whole, NULL, false));
        CHECK(JoinAddresses(found) == "B2 C3 D3");
    }

    SECTION("by columns")
    {
        XlSearchOrder byColumns = xlByColumns;
        REQUIRE(range.FindAll_("apple", found, &values, &part, &byColumns, false));
        CHECK(JoinAddresses(found) == "B2 B4 C3 D3 D4");
    }

    SECTION("matching the case")
    {
        REQUIRE(range.FindAll_("APPLE", found, &values, &part, NULL, true));
        CHECK(JoinAddresses(found) == "D3");
    }

    SECTION("only the bottom-right cell matches")
    {
        REQUIRE(range.FindAll_("apple pie", found, &values, &whole, NULL, false));
        CHECK(JoinAddresses(found) == "D4");
    }

    SECTION("nothing is found")
    {
        found.push_back(range);
        REQUIRE(range.FindAll_("durian", found, &values, &whole));
        CHECK(found.empty());
    }

    SECTION("in a range with several areas")
    {
        wxExcelRangeVector parts;
        parts.push_back(book.GetRange("D2:D4"));
        parts.push_back(book.GetRange("B2:B3"));
        wxExcelRange areas = wxAutoExcelTests::GetApplication().Union(parts);
        REQUIRE(areas);
        REQUIRE(areas.FindAll_("apple", found, &values, &part, NULL, false));
        CHECK(JoinAddresses(found) == "D3 D4 B2");
    }

    SECTION("in all the cells of the worksheet")
    {
        book.GetRange("XFD1048576").SetValue("apple");

        wxExcelRange cells = book.GetWorksheet().GetCells();
        REQUIRE(cells);
        REQUIRE(cells.FindAll_("apple", found, &values, &whole, NULL, false));
        CHECK(JoinAddresses(found) == "B2 C3 D3 XFD1048576");
    }
}

TEST_CASE("wxExcelRange::Find on a single cell",
          "[excel][range][find]")
{
    ExcelTestWorkbook book;
    REQUIRE(book.IsReady());
    FillTable(book);

    XlLookAt whole = xlWhole;
    XlFindLookIn values = xlValues;

    wxExcelRange cell = book.GetRange("A1");
    wxExcelRange found = cell.Find("cherry", NULL, &values, &whole);
    REQUIRE(cell);
    // Excel searches the whole worksheet, as in its user interface
    // when a single cell is selected
    CHECK(GetAddress(found) == "D2");
}

TEST_CASE("wxExcelRange::Replace replaces in the cells of the range",
          "[excel][range][find]")
{
    ExcelTestWorkbook book;
    REQUIRE(book.IsReady());
    wxExcelRange range = FillTable(book);
    book.GetRange("F1").SetValue("apple");

    XlLookAt whole = xlWhole;
    XlLookAt part = xlPart;

    SECTION("whole cells, ignoring the case")
    {
        REQUIRE(range.Replace("apple", "pear", &whole, NULL, false));
        CHECK(GetText(book, "B2") == "pear");
        CHECK(GetText(book, "C3") == "pear");
        CHECK(GetText(book, "D3") == "pear");
        CHECK(GetText(book, "B4") == "pineapple");
        CHECK(GetText(book, "D4") == "apple pie");
        // outside the range
        CHECK(GetText(book, "F1") == "apple");
    }

    SECTION("parts of cells, matching the case")
    {
        REQUIRE(range.Replace("apple", "pear", &part, NULL, true));
        CHECK(GetText(book, "B2") == "pear");
        CHECK(GetText(book, "D3") == "APPLE");
        CHECK(GetText(book, "B4") == "pinepear");
        CHECK(GetText(book, "D4") == "pear pie");
    }

    SECTION("a number")
    {
        REQUIRE(range.Replace(10.0, 11.0, &whole));
        wxExcelRange cell = book.GetRange("B3");
        CHECK(cell.GetValue().GetDouble() == 11.0);
        REQUIRE(cell);
    }

    SECTION("a single cell replaces in the whole worksheet")
    {
        wxExcelRange cell = book.GetRange("A1");
        REQUIRE(cell.Replace("cherry", "plum", &whole));
        CHECK(GetText(book, "D2") == "plum");
    }

    SECTION("nothing to replace is not an error")
    {
        CHECK(range.Replace("durian", "pear", &whole));
        CHECK(range);
    }
}
