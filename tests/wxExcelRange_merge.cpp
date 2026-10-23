/////////////////////////////////////////////////////////////////////////////
// Purpose:     Tests for merged wxExcelRange objects
// Author:      OpenAI Codex, under the direction of PB
//              Additional tests by Claude Opus 5.5 (Anthropic)
// Generated:   AI-generated using GPT-5.6 Sol with High reasoning effort
// Copyright:   (c) 2026 PB <pbfordev@gmail.com>
// License:     MIT license
/////////////////////////////////////////////////////////////////////////////

#include <catch2/catch_test_macros.hpp>

#include "excel_test_fixture.h"

using namespace wxAutoExcel;

TEST_CASE("wxExcelRange merges and unmerges a rectangular area",
          "[excel][range][merge]")
{
    ExcelTestWorkbook excel;
    REQUIRE(excel.IsReady());

    wxExcelRange range = excel.GetRange("B2:D3");
    wxExcelRange topLeft = excel.GetRange("B2");
    topLeft.SetValue("merged value");
    REQUIRE(topLeft);

    range.Merge();
    REQUIRE(range);
    CHECK(range.GetMergeCells());
    REQUIRE(range);

    wxExcelRange mergedCell = excel.GetRange("C3");
    CHECK(mergedCell.GetMergeCells());
    REQUIRE(mergedCell);

    wxExcelRange mergeArea = mergedCell.GetMergeArea();
    REQUIRE(mergeArea);
    CHECK(mergeArea.GetAddress() == "$B$2:$D$3");
    REQUIRE(mergeArea);
    CHECK(mergeArea.GetCount() == 6);
    REQUIRE(mergeArea);

    const wxVariant mergedValue = topLeft.GetValue();
    REQUIRE(topLeft);
    REQUIRE(mergedValue.IsType("string"));
    CHECK(mergedValue.GetString() == "merged value");

    // Excel's UnMerge method has no return value, so verify its postcondition.
    range.UnMerge();
    CHECK_FALSE(range.GetMergeCells());
    REQUIRE(range);

    const wxVariant retainedValue = topLeft.GetValue();
    REQUIRE(topLeft);
    REQUIRE(retainedValue.IsType("string"));
    CHECK(retainedValue.GetString() == "merged value");
}

TEST_CASE("wxExcelRange can merge each row separately",
          "[excel][range][merge][across]")
{
    ExcelTestWorkbook excel;
    REQUIRE(excel.IsReady());

    wxExcelRange range = excel.GetRange("B5:D6");
    range.Merge(true);
    REQUIRE(range);

    wxExcelRange firstRowArea = excel.GetRange("C5").GetMergeArea();
    REQUIRE(firstRowArea);
    CHECK(firstRowArea.GetAddress() == "$B$5:$D$5");
    REQUIRE(firstRowArea);
    CHECK(firstRowArea.GetCount() == 3);
    REQUIRE(firstRowArea);

    wxExcelRange secondRowArea = excel.GetRange("C6").GetMergeArea();
    REQUIRE(secondRowArea);
    CHECK(secondRowArea.GetAddress() == "$B$6:$D$6");
    REQUIRE(secondRowArea);
    CHECK(secondRowArea.GetCount() == 3);
    REQUIRE(secondRowArea);

    // Excel's UnMerge method has no return value, so verify its postcondition.
    range.UnMerge();
    CHECK_FALSE(range.GetMergeCells());
    REQUIRE(range);
}

TEST_CASE("wxExcelRange MergeArea of a cell which is not merged is the cell",
          "[excel][range][merge]")
{
    ExcelTestWorkbook excel;
    REQUIRE(excel.IsReady());

    wxExcelRange cell = excel.GetRange("F7");
    CHECK_FALSE(cell.GetMergeCells());
    REQUIRE(cell);

    wxExcelRange mergeArea = cell.GetMergeArea();
    REQUIRE(mergeArea);
    CHECK(mergeArea.GetAddress() == "$F$7");
    REQUIRE(mergeArea);
}

TEST_CASE("wxExcelRange merges each area of a multiple-area range",
          "[excel][range][merge][areas]")
{
    ExcelTestWorkbook excel;
    REQUIRE(excel.IsReady());

    // the areas are joined with the list separator of the system, so the
    // range is a union of two ranges instead of an address
    wxExcelRangeVector areas;
    areas.push_back(excel.GetRange("A1:B2"));
    areas.push_back(excel.GetRange("D1:E2"));
    wxExcelRange range = wxAutoExcelTests::GetApplication().Union(areas);
    REQUIRE(range);

    range.Merge();
    REQUIRE(range);

    wxExcelRange firstArea = excel.GetRange("B2").GetMergeArea();
    REQUIRE(firstArea);
    CHECK(firstArea.GetAddress() == "$A$1:$B$2");
    REQUIRE(firstArea);

    wxExcelRange secondArea = excel.GetRange("E2").GetMergeArea();
    REQUIRE(secondArea);
    CHECK(secondArea.GetAddress() == "$D$1:$E$2");
    REQUIRE(secondArea);

    // the cell between the areas is not merged
    CHECK_FALSE(excel.GetRange("C1").GetMergeCells());
}

TEST_CASE("wxExcelRange SetMergeCells merges and unmerges",
          "[excel][range][merge]")
{
    ExcelTestWorkbook excel;
    REQUIRE(excel.IsReady());

    wxExcelRange range = excel.GetRange("B2:C4");
    range.SetMergeCells(true);
    REQUIRE(range);
    CHECK(range.GetMergeCells());
    REQUIRE(range);

    wxExcelRange mergeArea = excel.GetRange("C4").GetMergeArea();
    REQUIRE(mergeArea);
    CHECK(mergeArea.GetAddress() == "$B$2:$C$4");
    REQUIRE(mergeArea);

    range.SetMergeCells(false);
    REQUIRE(range);
    CHECK_FALSE(range.GetMergeCells());
    REQUIRE(range);
}

TEST_CASE("wxExcelRange merging keeps only the value of the top left cell",
          "[excel][range][merge]")
{
    ExcelTestWorkbook excel;
    REQUIRE(excel.IsReady());

    wxExcelRange topLeft = excel.GetRange("A1");
    topLeft.SetValue("kept");
    REQUIRE(topLeft);
    wxExcelRange other = excel.GetRange("B2");
    other.SetValue("discarded");
    REQUIRE(other);

    // the tests run with DisplayAlerts off, otherwise Excel would ask
    // whether to discard the other values
    wxExcelRange range = excel.GetRange("A1:B2");
    range.Merge();
    REQUIRE(range);

    const wxVariant value = topLeft.GetValue();
    REQUIRE(topLeft);
    REQUIRE(value.IsType("string"));
    CHECK(value.GetString() == "kept");

    range.UnMerge();
    REQUIRE(range);
    CHECK(other.GetValue().IsNull());
    REQUIRE(other);
}

TEST_CASE("wxExcelRange GetMergeCells reports Null for partly merged ranges",
          "[excel][range][merge]")
{
    ExcelTestWorkbook excel;
    REQUIRE(excel.IsReady());

    wxExcelRange merged = excel.GetRange("A1:B1");
    merged.Merge();
    REQUIRE(merged);

    // Excel returns Null instead of a Boolean for a range which has both
    // merged and not merged cells, which is not an error
    wxExcelRange range = excel.GetRange("A1:A3");
    CHECK_FALSE(range.GetMergeCells());
    REQUIRE(range);
    CHECK(range.IsLastResultNull_());

    CHECK(merged.GetMergeCells());
    REQUIRE(merged);
    CHECK_FALSE(merged.IsLastResultNull_());
}
