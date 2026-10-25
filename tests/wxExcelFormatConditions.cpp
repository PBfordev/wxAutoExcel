/////////////////////////////////////////////////////////////////////////////
// Purpose:     Tests for wxExcelFormatConditions and the conditional
//              formatting rules
// Author:      Claude Opus 5.5 (Anthropic), under the direction of PB
// Copyright:   (c) 2026 PB <pbfordev@gmail.com>
// License:     MIT license
/////////////////////////////////////////////////////////////////////////////

#include <catch2/catch_test_macros.hpp>

#include "excel_test_fixture.h"

#if WXAUTOEXCEL_USE_CONDFORMAT

using namespace wxAutoExcel;

namespace
{

const wxColour Highlight(255, 199, 206);
const wxColour Other(198, 239, 206);
const wxColour White(255, 255, 255);

// Fills A1:A10 with the numbers 1 to 10 and returns the range.
wxExcelRange FillNumbers(ExcelTestWorkbook& book)
{
    wxExcelRange range = book.GetRange("A1:A10");
    for ( long row = 1; row <= 10; ++row )
    {
        wxExcelRange cell = range.GetItem(row);
        REQUIRE(cell);
        cell.SetValue(wxVariant(static_cast<double>(row)));
        REQUIRE(cell);
    }
    return range;
}

wxExcelFormatConditions GetFormatConditions(wxExcelRange range)
{
    wxExcelFormatConditions formatConditions = range.GetFormatConditions();
    REQUIRE(formatConditions);
    return formatConditions;
}

long GetCount(wxExcelFormatConditions formatConditions)
{
    const long count = formatConditions.GetCount();
    REQUIRE(formatConditions);
    return count;
}

void SetInteriorColor(wxExcelInterior interior, const wxColour& colour)
{
    REQUIRE(interior);
    interior.SetColor(colour);
    REQUIRE(interior);
}

// Returns the interior colour of the cell as displayed, i.e. including
// the conditional formatting.
wxColour GetDisplayedColor(ExcelTestWorkbook& book, const wxString& address)
{
    wxExcelDisplayFormat displayFormat = book.GetRange(address).GetDisplayFormat();
    REQUIRE(displayFormat);
    wxExcelInterior interior = displayFormat.GetInterior();
    REQUIRE(interior);
    const wxColour colour = interior.GetColor();
    REQUIRE(interior);
    return colour;
}

} // unnamed namespace

TEST_CASE("wxExcelFormatConditions adds a cell value rule which formats matching cells",
          "[excel][formatconditions]")
{
    ExcelTestWorkbook book;
    REQUIRE(book.IsReady());

    wxExcelRange range = FillNumbers(book);
    wxExcelFormatConditions formatConditions = GetFormatConditions(range);
    CHECK(GetCount(formatConditions) == 0);

    XlFormatConditionOperator greater = xlGreater;
    wxExcelFormatCondition condition = formatConditions.Add(xlCellValue, &greater, "=7");
    REQUIRE(condition);
    CHECK(GetCount(formatConditions) == 1);

    CHECK(condition.GetType() == xlCellValue);
    REQUIRE(condition);
    CHECK(condition.GetOperator() == xlGreater);
    REQUIRE(condition);
    CHECK(condition.GetFormula1() == "=7");
    REQUIRE(condition);

    wxExcelRange appliesTo = condition.GetAppliesTo();
    REQUIRE(appliesTo);
    CHECK(appliesTo.GetAddress() == "$A$1:$A$10");
    REQUIRE(appliesTo);

    SetInteriorColor(condition.GetInterior(), Highlight);

    CHECK(GetDisplayedColor(book, "A8") == Highlight);
    CHECK(GetDisplayedColor(book, "A10") == Highlight);
    CHECK(GetDisplayedColor(book, "A7") == White);

    // the rule does not change the format of the cell itself
    wxExcelInterior cellInterior = book.GetRange("A8").GetInterior();
    REQUIRE(cellInterior);
    CHECK(cellInterior.GetColor() == White);
    REQUIRE(cellInterior);

    // the rule is evaluated again when the value changes
    wxExcelRange cell = book.GetRange("A1");
    cell.SetValue(wxVariant(100.0));
    REQUIRE(cell);
    CHECK(GetDisplayedColor(book, "A1") == Highlight);

    XlFormatConditionOperator between = xlBetween;
    condition.Modify(xlCellValue, &between, "=2", "=4");
    REQUIRE(condition);
    CHECK(condition.GetFormula2() == "=4");
    REQUIRE(condition);
    CHECK(GetDisplayedColor(book, "A3") == Highlight);
    CHECK(GetDisplayedColor(book, "A8") == White);
}

TEST_CASE("wxExcelFormatConditions adds an expression rule",
          "[excel][formatconditions]")
{
    ExcelTestWorkbook book;
    REQUIRE(book.IsReady());

    wxExcelRange range = FillNumbers(book);
    wxExcelFormatConditions formatConditions = GetFormatConditions(range);

    // Excel interprets the formula with the separators and function names of
    // the system, so it uses none of them; the relative reference is relative
    // to the first cell of the range
    wxExcelFormatCondition belowA5 = formatConditions.Add(xlExpression, NULL, "=$A1<$A$5");
    REQUIRE(belowA5);
    CHECK(belowA5.GetType() == xlExpression);
    REQUIRE(belowA5);
    CHECK(belowA5.GetFormula1() == "=$A1<$A$5");
    REQUIRE(belowA5);
    SetInteriorColor(belowA5.GetInterior(), Highlight);

    CHECK(GetDisplayedColor(book, "A1") == Highlight);
    CHECK(GetDisplayedColor(book, "A4") == Highlight);
    CHECK(GetDisplayedColor(book, "A5") == White);

    belowA5.ModifyAppliesToRange(book.GetRange("A1:A2"));
    REQUIRE(belowA5);
    CHECK(GetDisplayedColor(book, "A2") == Highlight);
    CHECK(GetDisplayedColor(book, "A3") == White);

    wxExcelRange appliesTo = belowA5.GetAppliesTo();
    REQUIRE(appliesTo);
    CHECK(appliesTo.GetAddress() == "$A$1:$A$2");
    REQUIRE(appliesTo);
}

TEST_CASE("wxExcelFormatCondition priority decides which rule is applied",
          "[excel][formatconditions]")
{
    ExcelTestWorkbook book;
    REQUIRE(book.IsReady());

    wxExcelRange range = FillNumbers(book);
    wxExcelFormatConditions formatConditions = GetFormatConditions(range);

    XlFormatConditionOperator greater = xlGreater;
    wxExcelFormatCondition aboveFive = formatConditions.Add(xlCellValue, &greater, "=5");
    REQUIRE(aboveFive);
    SetInteriorColor(aboveFive.GetInterior(), Highlight);

    wxExcelFormatCondition aboveEight = formatConditions.Add(xlCellValue, &greater, "=8");
    REQUIRE(aboveEight);
    SetInteriorColor(aboveEight.GetInterior(), Other);

    CHECK(GetCount(formatConditions) == 2);

    aboveEight.SetFirstPriority();
    REQUIRE(aboveEight);
    CHECK(aboveEight.GetPriority() == 1);
    REQUIRE(aboveEight);
    CHECK(aboveFive.GetPriority() == 2);
    REQUIRE(aboveFive);

    // both rules are applied, the one with the higher priority wins
    CHECK(GetDisplayedColor(book, "A9") == Other);
    CHECK(GetDisplayedColor(book, "A6") == Highlight);

    aboveFive.SetFirstPriority();
    REQUIRE(aboveFive);
    CHECK(aboveFive.GetPriority() == 1);
    REQUIRE(aboveFive);
    CHECK(GetDisplayedColor(book, "A9") == Highlight);

    wxExcelFormatCondition first = formatConditions[1];
    REQUIRE(first);
    CHECK(first.GetFormula1() == "=5");
    REQUIRE(first);

    aboveFive.Delete();
    REQUIRE(aboveFive);
    CHECK(GetCount(formatConditions) == 1);
    CHECK(GetDisplayedColor(book, "A9") == Other);
    CHECK(GetDisplayedColor(book, "A6") == White);
}

TEST_CASE("wxExcelFormatConditions adds unique, top 10, data bar, and color scale rules",
          "[excel][formatconditions]")
{
    ExcelTestWorkbook book;
    REQUIRE(book.IsReady());

    wxExcelRange range = FillNumbers(book);
    // make 3 a duplicate
    wxExcelRange cell = book.GetRange("A4");
    cell.SetValue(wxVariant(3.0));
    REQUIRE(cell);

    wxExcelFormatConditions formatConditions = GetFormatConditions(range);

    wxExcelUniqueValues duplicates = formatConditions.AddUniqueValues();
    REQUIRE(duplicates);
    duplicates.SetDupeUnique(xlDuplicate);
    REQUIRE(duplicates);
    CHECK(duplicates.GetDupeUnique() == xlDuplicate);
    REQUIRE(duplicates);
    CHECK(duplicates.GetType() == xlUniqueValues);
    REQUIRE(duplicates);
    SetInteriorColor(duplicates.GetInterior(), Highlight);

    CHECK(GetDisplayedColor(book, "A3") == Highlight);
    CHECK(GetDisplayedColor(book, "A4") == Highlight);
    CHECK(GetDisplayedColor(book, "A5") == White);

    wxExcelTop10 top = formatConditions.AddTop10();
    REQUIRE(top);
    top.SetTopBottom(xlTop10Top);
    top.SetRank(2);
    top.SetPercent(false);
    REQUIRE(top);
    CHECK(top.GetRank() == 2);
    REQUIRE(top);
    CHECK(top.GetType() == xlTop10);
    REQUIRE(top);
    SetInteriorColor(top.GetInterior(), Other);

    CHECK(GetDisplayedColor(book, "A10") == Other);
    CHECK(GetDisplayedColor(book, "A9") == Other);
    CHECK(GetDisplayedColor(book, "A8") == White);

    REQUIRE(formatConditions.AddDatabar());
    REQUIRE(formatConditions.AddColorScale(3));
    CHECK(GetCount(formatConditions) == 4);

    formatConditions.Delete();
    REQUIRE(formatConditions);
    CHECK(GetCount(GetFormatConditions(range)) == 0);
    CHECK(GetDisplayedColor(book, "A10") == White);
}

TEST_CASE("wxExcelIconSetCondition uses an icon set of the workbook",
          "[excel][formatconditions][iconsets]")
{
    ExcelTestWorkbook book;
    REQUIRE(book.IsReady());

    wxExcelRange range = book.GetRange("A1:A10");
    range.SetFormula("=ROW()");
    REQUIRE(range);

    wxExcelIconSetCondition condition = GetFormatConditions(range).AddIconSetCondition();
    REQUIRE(condition);

    wxExcelIconSet iconSet = condition.GetIconSet();
    REQUIRE(iconSet);
    CHECK(iconSet.GetCount() == 3);
    REQUIRE(iconSet);

    wxExcelIconSets iconSets = book.GetWorkbook().GetIconSets();
    REQUIRE(iconSets);
    CHECK(iconSets.GetCount() >= 17);
    REQUIRE(iconSets);

    wxExcelIconSet arrows = book.GetWorkbook().GetIconSets(xl5Arrows);
    REQUIRE(arrows);
    CHECK(arrows.GetID() == xl5Arrows);
    CHECK(arrows.GetCount() == 5);
    REQUIRE(arrows);

    condition.SetIconSet(arrows);
    REQUIRE(condition);
    wxExcelIconSet changed = condition.GetIconSet();
    REQUIRE(changed);
    CHECK(changed.GetID() == xl5Arrows);
    CHECK(changed.GetCount() == 5);
    REQUIRE(changed);

    wxExcelIcon icon = changed.GetItem(5);
    REQUIRE(icon);
    CHECK(icon.GetIndex() == 5);
}

#endif // WXAUTOEXCEL_USE_CONDFORMAT
