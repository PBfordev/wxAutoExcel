/////////////////////////////////////////////////////////////////////////////
// Purpose:     Tests for sorting and filtering: wxExcelSort, wxExcelSortFields,
//              wxExcelAutoFilter, wxExcelFilters, and the wxExcelRange
//              methods which filter or remove duplicates
// Author:      Claude Opus 5.5 (Anthropic), under the direction of PB
// Copyright:   (c) 2026 PB <pbfordev@gmail.com>
// License:     MIT license
/////////////////////////////////////////////////////////////////////////////

#include <catch2/catch_test_macros.hpp>

#include "excel_test_fixture.h"

using namespace wxAutoExcel;

namespace
{

wxVariant GetValue(ExcelTestWorkbook& book, const wxString& address)
{
    wxExcelRange range = book.GetRange(address);
    const wxVariant value = range.GetValue();
    REQUIRE(range);
    return value;
}

wxString GetString(ExcelTestWorkbook& book, const wxString& address)
{
    const wxVariant value = GetValue(book, address);
    REQUIRE(value.IsType("string"));
    return value.GetString();
}

double GetNumber(ExcelTestWorkbook& book, const wxString& address)
{
    const wxVariant value = GetValue(book, address);
    REQUIRE(value.IsType("double"));
    return value.GetDouble();
}

void SetValue(ExcelTestWorkbook& book, const wxString& address, const wxVariant& value)
{
    wxExcelRange range = book.GetRange(address);
    range.SetValue(value);
    REQUIRE(range);
}

// Fills A1:C7 with a header row (Name, Team, Score) and six rows of data.
void FillScores(ExcelTestWorkbook& book)
{
    struct Row
    {
        const char* name;
        const char* team;
        double score;
    };
    static const Row rows[] =
    {
        { "Eve",   "Blue", 70 },
        { "Adam",  "Red",  90 },
        { "Dora",  "Blue", 40 },
        { "Bert",  "Red",  60 },
        { "Fred",  "Blue", 90 },
        { "Cecil", "Red",  30 },
    };

    SetValue(book, "A1", "Name");
    SetValue(book, "B1", "Team");
    SetValue(book, "C1", "Score");

    for ( size_t i = 0; i < WXSIZEOF(rows); ++i )
    {
        const wxString row = wxString::Format("%ld", static_cast<long>(i) + 2);
        SetValue(book, "A" + row, rows[i].name);
        SetValue(book, "B" + row, rows[i].team);
        SetValue(book, "C" + row, wxVariant(rows[i].score));
    }
}

// Returns the names in A2:A7, separated with commas.
wxString GetNames(ExcelTestWorkbook& book)
{
    wxString names;
    for ( long row = 2; row <= 7; ++row )
    {
        if ( !names.empty() )
            names += ",";
        names += GetString(book, wxString::Format("A%ld", row));
    }
    return names;
}

bool IsRowHidden(ExcelTestWorkbook& book, long row)
{
    wxExcelRange entireRow = book.GetRange(wxString::Format("A%ld", row)).GetEntireRow();
    REQUIRE(entireRow);
    const bool hidden = entireRow.GetHidden();
    REQUIRE(entireRow);
    return hidden;
}

long GetVisibleRowCount(ExcelTestWorkbook& book)
{
    long count = 0;
    for ( long row = 2; row <= 7; ++row )
    {
        if ( !IsRowHidden(book, row) )
            ++count;
    }
    return count;
}

wxExcelSort PrepareSort(ExcelTestWorkbook& book)
{
    wxExcelSort sort = book.GetWorksheet().GetSort();
    REQUIRE(sort);

    wxExcelSortFields fields = sort.GetSortFields();
    REQUIRE(fields);
    fields.Clear();
    REQUIRE(fields);

    return sort;
}

void ApplySort(ExcelTestWorkbook& book, wxExcelSort& sort)
{
    sort.SetRange(book.GetRange("A1:C7"));
    REQUIRE(sort);
    sort.SetHeader(xlYes);
    REQUIRE(sort);
    sort.Apply();
    REQUIRE(sort);
}

} // unnamed namespace

TEST_CASE("wxExcelSort sorts by one key in both orders",
          "[excel][sort]")
{
    ExcelTestWorkbook book;
    REQUIRE(book.IsReady());
    FillScores(book);

    wxExcelSort sort = PrepareSort(book);
    wxExcelSortFields fields = sort.GetSortFields();
    REQUIRE(fields);

    XlSortOrder ascending = xlAscending;
    wxExcelSortField field = fields.Add(book.GetRange("A2:A7"), xlSortOnValues, &ascending);
    REQUIRE(field);
    CHECK(fields.GetCount() == 1);
    REQUIRE(fields);
    CHECK(field.GetOrder() == xlAscending);
    REQUIRE(field);

    ApplySort(book, sort);
    CHECK(GetNames(book) == "Adam,Bert,Cecil,Dora,Eve,Fred");

    // the header row is not sorted
    CHECK(GetString(book, "A1") == "Name");

    field.SetOrder(xlDescending);
    REQUIRE(field);
    ApplySort(book, sort);
    CHECK(GetNames(book) == "Fred,Eve,Dora,Cecil,Bert,Adam");
}

TEST_CASE("wxExcelSort sorts by several keys and a custom order",
          "[excel][sort]")
{
    ExcelTestWorkbook book;
    REQUIRE(book.IsReady());
    FillScores(book);

    wxExcelSort sort = PrepareSort(book);
    wxExcelSortFields fields = sort.GetSortFields();
    REQUIRE(fields);

    XlSortOrder ascending = xlAscending;
    XlSortOrder descending = xlDescending;

    wxExcelSortField team = fields.Add(book.GetRange("B2:B7"), xlSortOnValues, &ascending);
    REQUIRE(team);
    wxExcelSortField score = fields.Add(book.GetRange("C2:C7"), xlSortOnValues, &descending);
    REQUIRE(score);
    CHECK(fields.GetCount() == 2);
    REQUIRE(fields);
    CHECK(team.GetPriority() == 1);
    REQUIRE(team);
    CHECK(score.GetPriority() == 2);
    REQUIRE(score);

    ApplySort(book, sort);
    CHECK(GetNames(book) == "Fred,Eve,Dora,Adam,Bert,Cecil");
    CHECK(GetNumber(book, "C2") == 90.0);
    CHECK(GetNumber(book, "C7") == 30.0);

    // the team field with a custom order, Red before Blue
    fields.Clear();
    REQUIRE(fields);
    team = fields.Add(book.GetRange("B2:B7"), xlSortOnValues, &ascending, "Red,Blue");
    REQUIRE(team);
    CHECK(team.GetCustomOrder() == "Red,Blue");
    REQUIRE(team);
    REQUIRE(fields.Add(book.GetRange("A2:A7"), xlSortOnValues, &ascending));

    ApplySort(book, sort);
    CHECK(GetNames(book) == "Adam,Bert,Cecil,Dora,Eve,Fred");
    CHECK(GetString(book, "B2") == "Red");
    CHECK(GetString(book, "B5") == "Blue");
}

TEST_CASE("wxExcelSort of a table sorts the table",
          "[excel][sort][listobject]")
{
    ExcelTestWorkbook book;
    REQUIRE(book.IsReady());
    FillScores(book);

    wxExcelRange source = book.GetRange("A1:C7");
    XlYesNoGuess hasHeaders = xlYes;
    wxExcelListObject table = book.GetWorksheet().GetListObjects().Add(&source, &hasHeaders);
    REQUIRE(table);

    wxExcelSort sort = table.GetSort();
    REQUIRE(sort);
    wxExcelSortFields fields = sort.GetSortFields();
    REQUIRE(fields);
    fields.Clear();
    REQUIRE(fields);

    XlSortOrder descending = xlDescending;
    wxExcelListColumn scoreColumn = table.GetListColumns().GetItem("Score");
    REQUIRE(scoreColumn);
    REQUIRE(fields.Add(scoreColumn.GetDataBodyRange(), xlSortOnValues, &descending));
    XlSortOrder ascending = xlAscending;
    REQUIRE(fields.Add(book.GetRange("A2:A7"), xlSortOnValues, &ascending));

    sort.Apply();
    REQUIRE(sort);
    CHECK(GetNames(book) == "Adam,Fred,Eve,Bert,Dora,Cecil");
}

TEST_CASE("wxExcelRange AutoFilter filters rows and wxExcelAutoFilter describes the filter",
          "[excel][autofilter]")
{
    ExcelTestWorkbook book;
    REQUIRE(book.IsReady());
    FillScores(book);

    wxExcelWorksheet& worksheet = book.GetWorksheet();
    CHECK_FALSE(worksheet.GetAutoFilterMode());
    REQUIRE(worksheet);

    wxExcelRange range = book.GetRange("A1:C7");
    long scoreField = 3;
    REQUIRE(range.AutoFilter(&scoreField, ">=60"));

    CHECK(worksheet.GetAutoFilterMode());
    REQUIRE(worksheet);
    CHECK(worksheet.GetFilterMode());
    REQUIRE(worksheet);

    // Eve 70, Adam 90, Bert 60, and Fred 90 remain
    CHECK(GetVisibleRowCount(book) == 4);
    CHECK(IsRowHidden(book, 4));
    CHECK(IsRowHidden(book, 7));
    CHECK_FALSE(IsRowHidden(book, 1));

    wxExcelAutoFilter autoFilter = worksheet.GetAutoFilter();
    REQUIRE(autoFilter);
    CHECK(autoFilter.GetFilterMode());
    REQUIRE(autoFilter);

    wxExcelRange filterRange = autoFilter.GetRange();
    REQUIRE(filterRange);
    CHECK(filterRange.GetAddress() == "$A$1:$C$7");
    REQUIRE(filterRange);

    wxExcelFilters filters = autoFilter.GetFilters();
    REQUIRE(filters);
    CHECK(filters.GetCount() == 3);
    REQUIRE(filters);

    wxExcelFilter nameFilter = filters.GetItem(1);
    REQUIRE(nameFilter);
    CHECK_FALSE(nameFilter.GetOn());
    REQUIRE(nameFilter);

    wxExcelFilter scoreFilter = filters.GetItem(3);
    REQUIRE(scoreFilter);
    CHECK(scoreFilter.GetOn());
    REQUIRE(scoreFilter);
    const wxVariant criteria = scoreFilter.GetCriteria1();
    REQUIRE(scoreFilter);
    REQUIRE(criteria.IsType("string"));
    CHECK(criteria.GetString() == ">=60");

    // filtering another column narrows the result
    long teamField = 2;
    REQUIRE(range.AutoFilter(&teamField, "Blue"));
    CHECK(GetVisibleRowCount(book) == 2);

    // ShowAllData shows all rows but keeps the AutoFilter
    autoFilter.ShowAllData();
    REQUIRE(autoFilter);
    CHECK(GetVisibleRowCount(book) == 6);
    CHECK_FALSE(worksheet.GetFilterMode());
    REQUIRE(worksheet);
    CHECK(worksheet.GetAutoFilterMode());
    REQUIRE(worksheet);

    // AutoFilter without arguments removes the AutoFilter
    REQUIRE(range.AutoFilter());
    CHECK_FALSE(worksheet.GetAutoFilterMode());
    REQUIRE(worksheet);
}

TEST_CASE("wxExcelRange AutoFilter combines two criteria",
          "[excel][autofilter]")
{
    ExcelTestWorkbook book;
    REQUIRE(book.IsReady());
    FillScores(book);

    wxExcelRange range = book.GetRange("A1:C7");
    long scoreField = 3;

    XlAutoFilterOperator orOperator = xlOr;
    REQUIRE(range.AutoFilter(&scoreField, "<40", &orOperator, ">80"));
    // Adam 90, Fred 90, and Cecil 30
    CHECK(GetVisibleRowCount(book) == 3);

    wxExcelFilter filter = book.GetWorksheet().GetAutoFilter().GetFilters().GetItem(3);
    REQUIRE(filter);
    CHECK(filter.GetOperator() == xlOr);
    REQUIRE(filter);
    const wxVariant criteria2 = filter.GetCriteria2();
    REQUIRE(filter);
    REQUIRE(criteria2.IsType("string"));
    CHECK(criteria2.GetString() == ">80");

    XlAutoFilterOperator andOperator = xlAnd;
    REQUIRE(range.AutoFilter(&scoreField, ">=40", &andOperator, "<=70"));
    // Eve 70, Dora 40, and Bert 60
    CHECK(GetVisibleRowCount(book) == 3);
    CHECK(IsRowHidden(book, 3));

    wxExcelWorksheet& worksheet = book.GetWorksheet();
    worksheet.ShowAllData();
    REQUIRE(worksheet);
    CHECK(GetVisibleRowCount(book) == 6);
}

TEST_CASE("wxExcelRange AdvancedFilter copies unique values",
          "[excel][autofilter][advancedfilter]")
{
    ExcelTestWorkbook book;
    REQUIRE(book.IsReady());
    FillScores(book);

    wxExcelRange teams = book.GetRange("B1:B7");
    wxExcelRange destination = book.GetRange("E1");
    REQUIRE(teams.AdvancedFilter(xlFilterCopy, NULL, &destination, true));

    CHECK(GetString(book, "E1") == "Team");
    CHECK(GetString(book, "E2") == "Blue");
    CHECK(GetString(book, "E3") == "Red");
    CHECK(GetValue(book, "E4").IsNull());
}

TEST_CASE("wxExcelRange RemoveDuplicates removes rows with duplicate values",
          "[excel][removeduplicates]")
{
    ExcelTestWorkbook book;
    REQUIRE(book.IsReady());
    FillScores(book);

    wxExcelRange range = book.GetRange("A1:C7");
    wxArrayLong columns;
    columns.push_back(3);
    XlYesNoGuess header = xlYes;
    range.RemoveDuplicates(columns, &header);
    REQUIRE(range);

    // the second row with score 90 (Fred) is removed
    CHECK(GetString(book, "A6") == "Cecil");
    CHECK(GetValue(book, "A7").IsNull());

    // with several columns, only the rows with all the values equal are removed
    SetValue(book, "A7", "Adam");
    SetValue(book, "B7", "Red");
    SetValue(book, "C7", wxVariant(91.0));
    columns.clear();
    columns.push_back(1);
    columns.push_back(2);
    range.RemoveDuplicates(columns, &header);
    REQUIRE(range);

    CHECK(GetString(book, "A6") == "Cecil");
    CHECK(GetValue(book, "A7").IsNull());
}
