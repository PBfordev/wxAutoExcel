/////////////////////////////////////////////////////////////////////////////
// Purpose:     Tests for wxExcelListObject (Excel tables) and its rows and
//              columns
// Author:      Claude Opus 5.5 (Anthropic), under the direction of PB
// Copyright:   (c) 2026 PB <pbfordev@gmail.com>
// License:     MIT license
/////////////////////////////////////////////////////////////////////////////

#include <catch2/catch_test_macros.hpp>

#include "excel_test_fixture.h"

using namespace wxAutoExcel;

namespace
{

wxString GetAddress(wxExcelRange range)
{
    REQUIRE(range);
    const wxString address = range.GetAddress();
    REQUIRE(range);
    return address;
}

wxVariant GetValue(ExcelTestWorkbook& book, const wxString& address)
{
    wxExcelRange range = book.GetRange(address);
    const wxVariant value = range.GetValue();
    REQUIRE(range);
    return value;
}

// Fills A1:C5 with a header row and four rows of data:
// Name, Region, Amount; Anna, North, 10; Bob, South, 20; ...
void FillSales(ExcelTestWorkbook& book)
{
    static const char* const names[] = { "Anna", "Bob", "Carl", "Dana" };
    static const char* const regions[] = { "North", "South", "North", "South" };

    wxExcelRange header = book.GetRange("A1");
    header.SetValue("Name");
    book.GetRange("B1").SetValue("Region");
    book.GetRange("C1").SetValue("Amount");
    REQUIRE(header);

    for ( long i = 0; i < 4; ++i )
    {
        const wxString row = wxString::Format("%ld", i + 2);
        wxExcelRange name = book.GetRange("A" + row);
        name.SetValue(names[i]);
        REQUIRE(name);
        wxExcelRange region = book.GetRange("B" + row);
        region.SetValue(regions[i]);
        REQUIRE(region);
        wxExcelRange amount = book.GetRange("C" + row);
        amount.SetValue(wxVariant(10.0 * (i + 1)));
        REQUIRE(amount);
    }
}

wxExcelListObject AddSalesTable(ExcelTestWorkbook& book)
{
    FillSales(book);

    wxExcelListObjects listObjects = book.GetWorksheet().GetListObjects();
    REQUIRE(listObjects);

    wxExcelRange source = book.GetRange("A1:C5");
    XlYesNoGuess hasHeaders = xlYes;
    wxExcelListObject table = listObjects.Add(&source, &hasHeaders);
    REQUIRE(table);

    table.SetName("Sales");
    REQUIRE(table);
    return table;
}

long GetTableCount(ExcelTestWorkbook& book)
{
    wxExcelListObjects listObjects = book.GetWorksheet().GetListObjects();
    REQUIRE(listObjects);
    const long count = listObjects.GetCount();
    REQUIRE(listObjects);
    return count;
}

} // unnamed namespace

TEST_CASE("wxExcelListObjects creates a table with headers from a range",
          "[excel][listobject]")
{
    ExcelTestWorkbook book;
    REQUIRE(book.IsReady());

    wxExcelListObject table = AddSalesTable(book);

    CHECK(GetTableCount(book) == 1);
    CHECK(table.GetName() == "Sales");
    REQUIRE(table);
    CHECK(table.GetShowHeaders());
    REQUIRE(table);
    CHECK(table.GetSourceType() == xlSrcRange);
    REQUIRE(table);

    CHECK(GetAddress(table.GetRange()) == "$A$1:$C$5");
    CHECK(GetAddress(table.GetHeaderRowRange()) == "$A$1:$C$1");
    CHECK(GetAddress(table.GetDataBodyRange()) == "$A$2:$C$5");

    wxExcelListObjects listObjects = book.GetWorksheet().GetListObjects();
    REQUIRE(listObjects);
    wxExcelListObject byName = listObjects.GetItem("Sales");
    REQUIRE(byName);
    CHECK(GetAddress(byName.GetRange()) == "$A$1:$C$5");

    wxExcelListColumns columns = table.GetListColumns();
    REQUIRE(columns);
    CHECK(columns.GetCount() == 3);
    REQUIRE(columns);

    wxExcelListColumn amount = columns.GetItem("Amount");
    REQUIRE(amount);
    CHECK(amount.GetIndex() == 3);
    REQUIRE(amount);
    CHECK(GetAddress(amount.GetDataBodyRange()) == "$C$2:$C$5");

    wxExcelListRows rows = table.GetListRows();
    REQUIRE(rows);
    CHECK(rows.GetCount() == 4);
    REQUIRE(rows);

    wxExcelListRow secondRow = rows.GetItem(2);
    REQUIRE(secondRow);
    CHECK(GetAddress(secondRow.GetRange()) == "$A$3:$C$3");
}

TEST_CASE("wxExcelListRows adds rows at the end and at a position and deletes them",
          "[excel][listobject][listrow]")
{
    ExcelTestWorkbook book;
    REQUIRE(book.IsReady());

    wxExcelListObject table = AddSalesTable(book);
    wxExcelListRows rows = table.GetListRows();
    REQUIRE(rows);

    wxExcelListRow added = rows.Add();
    REQUIRE(added);
    CHECK(added.GetIndex() == 5);
    REQUIRE(added);
    CHECK(GetAddress(added.GetRange()) == "$A$6:$C$6");
    CHECK(GetAddress(table.GetDataBodyRange()) == "$A$2:$C$6");

    long position = 1;
    wxExcelListRow first = rows.Add(&position);
    REQUIRE(first);
    CHECK(GetAddress(first.GetRange()) == "$A$2:$C$2");
    CHECK(rows.GetCount() == 6);
    REQUIRE(rows);

    // the original first row moved down
    const wxVariant anna = GetValue(book, "A3");
    REQUIRE(anna.IsType("string"));
    CHECK(anna.GetString() == "Anna");

    first.Delete();
    REQUIRE(first);
    CHECK(rows.GetCount() == 5);
    REQUIRE(rows);

    const wxVariant annaAgain = GetValue(book, "A2");
    REQUIRE(annaAgain.IsType("string"));
    CHECK(annaAgain.GetString() == "Anna");
}

TEST_CASE("wxExcelListColumns adds, renames, and deletes columns",
          "[excel][listobject][listcolumn]")
{
    ExcelTestWorkbook book;
    REQUIRE(book.IsReady());

    wxExcelListObject table = AddSalesTable(book);
    wxExcelListColumns columns = table.GetListColumns();
    REQUIRE(columns);

    wxExcelListColumn added = columns.Add();
    REQUIRE(added);
    CHECK(added.GetIndex() == 4);
    REQUIRE(added);

    added.SetName("Note");
    REQUIRE(added);
    CHECK(added.GetName() == "Note");
    REQUIRE(added);

    const wxVariant header = GetValue(book, "D1");
    REQUIRE(header.IsType("string"));
    CHECK(header.GetString() == "Note");
    CHECK(GetAddress(table.GetRange()) == "$A$1:$D$5");

    long position = 1;
    wxExcelListColumn first = columns.Add(&position);
    REQUIRE(first);
    CHECK(GetAddress(first.GetRange()) == "$A$1:$A$5");
    CHECK(columns.GetCount() == 5);
    REQUIRE(columns);

    first.Delete();
    REQUIRE(first);
    columns.GetItem("Note").Delete();
    CHECK(columns.GetCount() == 3);
    REQUIRE(columns);
    CHECK(GetAddress(table.GetRange()) == "$A$1:$C$5");
}

TEST_CASE("wxExcelListObject shows a totals row and is used in structured references",
          "[excel][listobject][totals]")
{
    ExcelTestWorkbook book;
    REQUIRE(book.IsReady());

    wxExcelListObject table = AddSalesTable(book);

    table.SetShowTotals(true);
    REQUIRE(table);
    CHECK(table.GetShowTotals());
    REQUIRE(table);
    CHECK(GetAddress(table.GetTotalsRowRange()) == "$A$6:$C$6");

    wxExcelListColumn amount = table.GetListColumns().GetItem("Amount");
    REQUIRE(amount);
    amount.SetTotalsCalculation(xlTotalsCalculationAverage);
    REQUIRE(amount);
    CHECK(amount.GetTotalsCalculation() == xlTotalsCalculationAverage);
    REQUIRE(amount);

    wxExcelRange total = amount.GetTotal();
    REQUIRE(total);
    const wxVariant average = total.GetValue();
    REQUIRE(total);
    REQUIRE(average.IsType("double"));
    CHECK(average.GetDouble() == 25.0);

    wxExcelRange formulaCell = book.GetRange("E1");
    formulaCell.SetFormula("=SUM(Sales[Amount])");
    REQUIRE(formulaCell);
    const wxVariant sum = formulaCell.GetValue();
    REQUIRE(formulaCell);
    REQUIRE(sum.IsType("double"));
    CHECK(sum.GetDouble() == 100.0);

    table.SetShowTotals(false);
    REQUIRE(table);
    CHECK(GetAddress(table.GetRange()) == "$A$1:$C$5");
}

TEST_CASE("wxExcelListObject resizes, converts to a range, and deletes",
          "[excel][listobject]")
{
    ExcelTestWorkbook book;
    REQUIRE(book.IsReady());

    wxExcelListObject table = AddSalesTable(book);

    table.Resize(book.GetRange("A1:B3"));
    REQUIRE(table);
    CHECK(GetAddress(table.GetRange()) == "$A$1:$B$3");
    CHECK(table.GetListRows().GetCount() == 2);
    REQUIRE(table);

    table.Unlist();
    REQUIRE(table);
    CHECK(GetTableCount(book) == 0);

    // converting to a range keeps the values
    const wxVariant bob = GetValue(book, "A3");
    REQUIRE(bob.IsType("string"));
    CHECK(bob.GetString() == "Bob");

    wxExcelListObject another = AddSalesTable(book);
    CHECK(GetTableCount(book) == 1);

    another.Delete();
    REQUIRE(another);
    CHECK(GetTableCount(book) == 0);

    // deleting the table clears its cells
    CHECK(GetValue(book, "A1").IsNull());
    CHECK(GetValue(book, "C5").IsNull());
}
