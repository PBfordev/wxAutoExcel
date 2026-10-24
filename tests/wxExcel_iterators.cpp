/////////////////////////////////////////////////////////////////////////////
// Purpose:     Tests for iterating over collections and ranges
// Author:      Claude Opus 5.5 (Anthropic), under the direction of PB
// Copyright:   (c) 2026 PB <pbfordev@gmail.com>
// License:     MIT license
/////////////////////////////////////////////////////////////////////////////

#include <catch2/catch_test_macros.hpp>

#include <iterator>

#include "excel_test_fixture.h"

using namespace wxAutoExcel;

namespace
{

wxString GetAddress(wxExcelRange& range)
{
    REQUIRE(range.IsOk_());
    const wxString address = range.GetAddress(false, false);
    REQUIRE(range);
    return address;
}

} // unnamed namespace

TEST_CASE("Collections are iterated over with a range-based for loop",
          "[excel][iterator]")
{
    ExcelTestWorkbook book;
    REQUIRE(book.IsReady());

    wxExcelWorksheets worksheets = book.GetWorksheets();
    const long count = worksheets.GetCount();
    REQUIRE(worksheets);

    wxExcelWorksheet added = worksheets.Add();
    REQUIRE(added);
    added.SetName("Added");
    REQUIRE(added);

    wxArrayString names;
    for ( wxExcelWorksheet worksheet : worksheets )
    {
        REQUIRE(worksheet);
        names.push_back(worksheet.GetName());
        REQUIRE(worksheet);
    }

    REQUIRE(names.size() == static_cast<size_t>(count + 1));
    CHECK(names.Index("Added") != wxNOT_FOUND);

    // the order is that of the indices
    for ( wxExcelWorksheets::iterator it = worksheets.begin(); it != worksheets.end(); ++it )
    {
        wxExcelWorksheet worksheet = *it;
        CHECK(worksheet.GetIndex() == it.GetIndex());
        REQUIRE(worksheet);
    }
}

TEST_CASE("An empty collection has nothing to iterate over",
          "[excel][iterator]")
{
    ExcelTestWorkbook book;
    REQUIRE(book.IsReady());

    wxExcelNames names = book.GetWorkbook().GetNames();
    REQUIRE(names);

    long iterations = 0;
    for ( wxExcelName name : names )
    {
        wxUnusedVar(name);
        ++iterations;
    }
    CHECK(iterations == 0);
    CHECK(names.begin() == names.end());
}

TEST_CASE("wxExcelCollectionIterator works as an input iterator",
          "[excel][iterator]")
{
    ExcelTestWorkbook book;
    REQUIRE(book.IsReady());

    wxExcelWorksheets worksheets = book.GetWorksheets();
    worksheets.Add();
    worksheets.Add();
    REQUIRE(worksheets);
    const long count = worksheets.GetCount();
    REQUIRE(count >= 3);

    CHECK(std::distance(worksheets.begin(), worksheets.end()) == count);

    wxExcelWorksheets::iterator it = worksheets.begin();
    CHECK(it.GetIndex() == 1);
    wxExcelWorksheets::iterator previous = it++;
    CHECK(previous.GetIndex() == 1);
    CHECK(it.GetIndex() == 2);
    CHECK(previous != it);
    CHECK(++previous == it);

    // iterators of different collections differ
    wxExcelWorksheets other = book.GetWorksheets();
    REQUIRE(other);
    CHECK(other.begin() != worksheets.begin());

    // a default-constructed iterator equals only another one
    CHECK(wxExcelWorksheets::iterator() == wxExcelWorksheets::iterator());
    CHECK(wxExcelWorksheets::iterator() != worksheets.begin());
}

TEST_CASE("wxExcelRange iterates over its cells, rows, or columns",
          "[excel][iterator][range]")
{
    ExcelTestWorkbook book;
    REQUIRE(book.IsReady());

    wxExcelRange range = book.GetRange("B2:D3");
    wxString addresses;

    SECTION("cells, across the rows")
    {
        for ( wxExcelRange cell : range )
            addresses += GetAddress(cell) + " ";
        CHECK(addresses == "B2 C2 D2 B3 C3 D3 ");
    }

    SECTION("rows")
    {
        wxExcelRange rows = range.GetRows();
        REQUIRE(rows);
        for ( wxExcelRange row : rows )
            addresses += GetAddress(row) + " ";
        CHECK(addresses == "B2:D2 B3:D3 ");
    }

    SECTION("columns")
    {
        wxExcelRange columns = range.GetColumns();
        REQUIRE(columns);
        for ( wxExcelRange column : columns )
            addresses += GetAddress(column) + " ";
        CHECK(addresses == "B2:B3 C2:C3 D2:D3 ");
    }

    SECTION("the cells of each row")
    {
        for ( wxExcelRange row : range.GetRows() )
        {
            for ( wxExcelRange cell : row.GetCells() )
                addresses += GetAddress(cell) + " ";
            addresses += "| ";
        }
        CHECK(addresses == "B2 C2 D2 | B3 C3 D3 | ");
    }

    SECTION("a row is a range of rows with a single row")
    {
        wxExcelRange rows = range.GetRows();
        REQUIRE(rows);
        wxExcelRange row = rows[1];
        for ( wxExcelRange item : row )
            addresses += GetAddress(item) + " ";
        CHECK(addresses == "B2:D2 ");
    }

    SECTION("the index operator")
    {
        wxExcelRange cell = range[5];
        CHECK(GetAddress(cell) == "C3");

        wxExcelRange rows = range.GetRows();
        REQUIRE(rows);
        wxExcelRange row = rows[2];
        CHECK(GetAddress(row) == "B3:D3");

        wxExcelRange columns = range.GetColumns();
        REQUIRE(columns);
        wxExcelRange column = columns[3];
        CHECK(GetAddress(column) == "D2:D3");
    }

    SECTION("a single cell")
    {
        wxExcelRange single = book.GetRange("E5");
        for ( wxExcelRange cell : single )
            addresses += GetAddress(cell) + " ";
        CHECK(addresses == "E5 ");
    }
}

TEST_CASE("wxExcelRange areas are iterated over one by one",
          "[excel][iterator][range]")
{
    ExcelTestWorkbook book;
    REQUIRE(book.IsReady());

    wxExcelRangeVector parts;
    parts.push_back(book.GetRange("A1:B1"));
    parts.push_back(book.GetRange("D4"));
    wxExcelRange united = wxAutoExcelTests::GetApplication().Union(parts);
    REQUIRE(united);

    wxString addresses;
    for ( wxExcelRange area : united.GetAreas() )
    {
        for ( wxExcelRange cell : area )
            addresses += GetAddress(cell) + " ";
    }
    CHECK(addresses == "A1 B1 D4 ");
}
