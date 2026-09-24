/////////////////////////////////////////////////////////////////////////////
// Purpose:     Tests for merged wxExcelRange objects
// Author:      OpenAI Codex, under the direction of PB
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
