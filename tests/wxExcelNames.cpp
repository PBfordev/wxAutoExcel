/////////////////////////////////////////////////////////////////////////////
// Purpose:     Tests for wxExcelNames and wxExcelName
// Author:      Claude Opus 5.5 (Anthropic), under the direction of PB
// Copyright:   (c) 2026 PB <pbfordev@gmail.com>
// License:     MIT license
/////////////////////////////////////////////////////////////////////////////

#include <catch2/catch_test_macros.hpp>

#include "excel_test_fixture.h"

using namespace wxAutoExcel;

namespace
{

// Gives the worksheet a known name, so that the references do not depend
// on the language of Excel.
void NameWorksheet(ExcelTestWorkbook& book)
{
    wxExcelWorksheet& worksheet = book.GetWorksheet();
    worksheet.SetName("Data");
    REQUIRE(worksheet);
}

wxExcelNames GetWorkbookNames(ExcelTestWorkbook& book)
{
    wxExcelNames names = book.GetWorkbook().GetNames();
    REQUIRE(names);
    return names;
}

long GetCount(wxExcelNames& names)
{
    const long count = names.GetCount();
    REQUIRE(names);
    return count;
}

double GetNumber(ExcelTestWorkbook& book, const wxString& address)
{
    wxExcelRange range = book.GetRange(address);
    const wxVariant value = range.GetValue();
    REQUIRE(range);
    REQUIRE(value.IsType("double"));
    return value.GetDouble();
}

} // unnamed namespace

TEST_CASE("wxExcelNames adds a name referring to a range",
          "[excel][names]")
{
    ExcelTestWorkbook book;
    REQUIRE(book.IsReady());
    NameWorksheet(book);

    wxExcelRange rateCell = book.GetRange("A1");
    rateCell.SetValue(wxVariant(0.5));
    REQUIRE(rateCell);

    wxExcelNames names = GetWorkbookNames(book);
    const long initialCount = GetCount(names);

    wxExcelName rate = names.Add("Rate", "=Data!$A$1");
    REQUIRE(rate);
    CHECK(GetCount(names) == initialCount + 1);

    CHECK(rate.GetName() == "Rate");
    REQUIRE(rate);
    CHECK(rate.GetRefersTo() == "=Data!$A$1");
    REQUIRE(rate);
    CHECK(rate.GetRefersToR1C1() == "=Data!R1C1");
    REQUIRE(rate);
    CHECK(rate.GetVisible());
    REQUIRE(rate);

    wxExcelRange refersToRange = rate.GetRefersToRange();
    REQUIRE(refersToRange);
    CHECK(refersToRange.GetAddress() == "$A$1");
    REQUIRE(refersToRange);

    wxExcelName byName = names.Item("Rate");
    REQUIRE(byName);
    CHECK(byName.GetRefersTo() == "=Data!$A$1");
    REQUIRE(byName);

    wxExcelRange formulaCell = book.GetRange("B1");
    formulaCell.SetFormula("=Rate*4");
    REQUIRE(formulaCell);
    CHECK(GetNumber(book, "B1") == 2.0);

    // changing what the name refers to changes the result of the formula
    wxExcelRange otherCell = book.GetRange("A2");
    otherCell.SetValue(wxVariant(3.0));
    REQUIRE(otherCell);
    rate.SetRefersTo("=Data!$A$2");
    REQUIRE(rate);
    CHECK(GetNumber(book, "B1") == 12.0);
}

TEST_CASE("wxExcelRange names a range and returns its name",
          "[excel][names][range]")
{
    ExcelTestWorkbook book;
    REQUIRE(book.IsReady());
    NameWorksheet(book);

    wxExcelRange range = book.GetRange("B2:C3");
    range.SetName("Block");
    REQUIRE(range);

    wxExcelName name = range.GetName();
    REQUIRE(name);
    CHECK(name.GetName() == "Block");
    REQUIRE(name);
    CHECK(name.GetRefersTo() == "=Data!$B$2:$C$3");
    REQUIRE(name);

    wxExcelRange byName = book.GetWorksheet().GetRange("Block");
    REQUIRE(byName);
    CHECK(byName.GetAddress() == "$B$2:$C$3");
    REQUIRE(byName);
}

TEST_CASE("wxExcelNames adds names of constants and formulas",
          "[excel][names]")
{
    ExcelTestWorkbook book;
    REQUIRE(book.IsReady());
    NameWorksheet(book);

    wxExcelNames names = GetWorkbookNames(book);

    // Excel interprets the formula passed to Add with the separators of the
    // system, so the number is set with SetRefersTo, which uses the English ones
    wxExcelName constant = names.Add("Tax", "=1");
    REQUIRE(constant);
    constant.SetRefersTo("=0.25");
    REQUIRE(constant);
    CHECK(constant.GetValue() == "=0.25");
    REQUIRE(constant);

    // a name which does not refer to a range has no RefersToRange
    {
        wxLogNull noLog;
        wxExcelName copy = constant;
        wxExcelRange noRange = copy.GetRefersToRange();
        CHECK_FALSE(noRange);
        CHECK_FALSE(copy);
    }

    wxExcelName formula = names.Add("Doubled", "=Tax*2");
    REQUIRE(formula);

    wxExcelRange cell = book.GetRange("A1");
    cell.SetFormula("=Doubled+Tax");
    REQUIRE(cell);
    CHECK(GetNumber(book, "A1") == 0.75);

    wxExcelName hidden = names.Add("Hidden", "=1", false);
    REQUIRE(hidden);
    CHECK_FALSE(hidden.GetVisible());
    REQUIRE(hidden);

    hidden.SetVisible(true);
    REQUIRE(hidden);
    CHECK(hidden.GetVisible());
    REQUIRE(hidden);
}

TEST_CASE("wxExcelNames of a worksheet contain names with worksheet scope",
          "[excel][names][worksheet]")
{
    ExcelTestWorkbook book;
    REQUIRE(book.IsReady());
    NameWorksheet(book);

    wxExcelNames workbookNames = GetWorkbookNames(book);
    const long initialCount = GetCount(workbookNames);

    wxExcelNames worksheetNames = book.GetWorksheet().GetNames();
    REQUIRE(worksheetNames);
    REQUIRE(GetCount(worksheetNames) == 0);

    wxExcelName local = worksheetNames.Add("Local", "=Data!$C$3");
    REQUIRE(local);
    CHECK(GetCount(worksheetNames) == 1);

    // the names with worksheet scope are prefixed with the worksheet name
    CHECK(local.GetName() == "Data!Local");
    REQUIRE(local);

    // and they are in the workbook names too
    CHECK(GetCount(workbookNames) == initialCount + 1);
    wxExcelName fromWorkbook = workbookNames.Item("Data!Local");
    REQUIRE(fromWorkbook);
    CHECK(fromWorkbook.GetRefersTo() == "=Data!$C$3");
    REQUIRE(fromWorkbook);
}

TEST_CASE("wxExcelName is deleted and retrieved by its index",
          "[excel][names]")
{
    ExcelTestWorkbook book;
    REQUIRE(book.IsReady());
    NameWorksheet(book);

    wxExcelNames names = GetWorkbookNames(book);
    const long initialCount = GetCount(names);

    REQUIRE(names.Add("First", "=1"));
    REQUIRE(names.Add("Second", "=2"));
    REQUIRE(GetCount(names) == initialCount + 2);

    bool foundSecond = false;
    for ( long i = 1; i <= GetCount(names); ++i )
    {
        wxExcelName name = names.Item(i);
        REQUIRE(name);
        CHECK(name.GetIndex() == i);
        REQUIRE(name);
        if ( name.GetName() == "Second" )
            foundSecond = true;
        REQUIRE(name);
    }
    CHECK(foundSecond);

    wxExcelName first = names.Item("First");
    REQUIRE(first);
    first.Delete();
    REQUIRE(first);
    CHECK(GetCount(names) == initialCount + 1);

    // the deleted name can no longer be retrieved
    {
        wxLogNull noLog;
        wxExcelName deleted = names.Item("First");
        CHECK_FALSE(deleted);
    }
}
