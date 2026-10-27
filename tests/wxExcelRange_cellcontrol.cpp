/////////////////////////////////////////////////////////////////////////////
// Purpose:     Tests for wxExcelCellControl and the wxExcelRange methods
//              for cell controls
// Author:      Claude Opus 5.5 (Anthropic), under the direction of PB
// Copyright:   (c) 2026 PB <pbfordev@gmail.com>
// License:     MIT license
/////////////////////////////////////////////////////////////////////////////

#include <catch2/catch_test_macros.hpp>

#include "excel_test_fixture.h"

using namespace wxAutoExcel;

namespace
{

// Skips the test when Excel does not support cell controls.
void SkipWithoutCellControls(ExcelTestWorkbook& excel)
{
    if ( !excel.GetRange("A1").HasProperty_("CellControl") )
        SKIP("Excel does not support cell controls");
}

// Returns the type of the cell controls of the range.
XlCellControlType GetControlType(ExcelTestWorkbook& excel, const wxString& address)
{
    wxExcelCellControl cellControl = excel.GetRange(address).GetCellControl();
    REQUIRE(cellControl);
    return cellControl.GetType();
}

} // namespace

TEST_CASE("wxExcelCellControl adds checkboxes and returns their type",
          "[excel][range][cellcontrol]")
{
    ExcelTestWorkbook excel;
    REQUIRE(excel.IsReady());
    SkipWithoutCellControls(excel);

    CHECK(GetControlType(excel, "A1") == xlTypeNone);

    wxExcelCellControl cellControl = excel.GetRange("A1:A3").GetCellControl();
    REQUIRE(cellControl);
    CHECK(cellControl.SetCheckbox());

    CHECK(GetControlType(excel, "A1") == xlTypeCheckbox);
    CHECK(cellControl.GetType() == xlTypeCheckbox);
    CHECK_FALSE(cellControl.IsLastResultNull_());

    // an empty cell gets the value of an unchecked checkbox
    const wxVariant value = excel.GetRange("A1").GetValue();
    REQUIRE(value.IsType("bool"));
    CHECK_FALSE(value.GetBool());

    SECTION("a range with different types of cell controls")
    {
        wxExcelCellControl mixed = excel.GetRange("A1:A4").GetCellControl();
        REQUIRE(mixed);
        CHECK(mixed.GetType() == xlTypeNone);
        CHECK(mixed);
        CHECK(mixed.IsLastResultNull_());
    }

    SECTION("checking a checkbox")
    {
        wxExcelRange cell = excel.GetRange("A2");
        cell.SetValue(true);
        REQUIRE(cell);
        CHECK(cell.GetValue().GetBool());
        CHECK(GetControlType(excel, "A2") == xlTypeCheckbox);
    }
}

TEST_CASE("wxExcelRange ResetContents resets and then removes cell controls",
          "[excel][range][cellcontrol]")
{
    ExcelTestWorkbook excel;
    REQUIRE(excel.IsReady());
    SkipWithoutCellControls(excel);

    REQUIRE(excel.GetRange("A1:A2").GetCellControl().SetCheckbox());
    excel.GetRange("A1").SetValue(true);
    excel.GetRange("A3").SetValue("text");

    wxExcelRange range = excel.GetRange("A1:A3");

    // the checkboxes are unchecked and the other cells cleared
    CHECK(range.ResetContents());
    CHECK(GetControlType(excel, "A1") == xlTypeCheckbox);
    CHECK_FALSE(excel.GetRange("A1").GetValue().GetBool());
    CHECK(excel.GetRange("A3").GetValue().IsNull());

    // with only unchecked checkboxes and empty cells, the checkboxes are removed
    CHECK(range.ResetContents());
    CHECK(GetControlType(excel, "A1") == xlTypeNone);
    CHECK(excel.GetRange("A1").GetValue().IsNull());
}

TEST_CASE("wxExcelRange RemoveControls and ClearContents remove cell controls",
          "[excel][range][cellcontrol]")
{
    ExcelTestWorkbook excel;
    REQUIRE(excel.IsReady());
    SkipWithoutCellControls(excel);

    wxExcelRange cell = excel.GetRange("B1");
    REQUIRE(cell.GetCellControl().SetCheckbox());
    cell.SetValue(true);
    REQUIRE(cell);

    SECTION("RemoveControls keeps the values")
    {
        CHECK(cell.RemoveControls());
        CHECK(GetControlType(excel, "B1") == xlTypeNone);
        CHECK(cell.GetValue().GetBool());
    }

    SECTION("ClearContents keeps the cell controls by default")
    {
        CHECK(cell.ClearContents());
        CHECK(GetControlType(excel, "B1") == xlTypeCheckbox);
        CHECK(cell.GetValue().IsNull());
    }

    SECTION("ClearContents removes the cell controls when asked to")
    {
        CHECK(cell.ClearContents(true));
        CHECK(GetControlType(excel, "B1") == xlTypeNone);
        CHECK(cell.GetValue().IsNull());
    }
}
