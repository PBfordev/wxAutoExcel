/////////////////////////////////////////////////////////////////////////////
// Purpose:     Tests for wxExcelRange dynamic array formulas
// Author:      Claude Opus 5.5 (Anthropic), under the direction of PB
// Copyright:   (c) 2026 PB <pbfordev@gmail.com>
// License:     MIT license
/////////////////////////////////////////////////////////////////////////////

#include <catch2/catch_test_macros.hpp>

#include <vector>

#include <wx/msw/ole/oleutils.h>

#include "excel_test_fixture.h"

using namespace wxAutoExcel;

namespace
{

// Fills A1:A5 with the numbers 5, 4, 3, 2, 1.
void FillNumbers(ExcelTestWorkbook& excel)
{
    for ( long row = 1; row <= 5; ++row )
    {
        wxExcelRange cell = excel.GetRange(wxString::Format("A%ld", row));
        cell.SetValue(6 - row);
        REQUIRE(cell);
    }
}

// Skips the test when Excel does not support dynamic arrays, i.e. it is
// older than Excel 2021.
void SkipWithoutDynamicArrays(ExcelTestWorkbook& excel)
{
    if ( !excel.GetRange("A1").HasProperty_("Formula2") )
        SKIP("Excel does not support dynamic arrays");
}

SCODE GetExcelErrorCode(const wxVariant& value)
{
    REQUIRE(value.IsType("errorcode"));
    return static_cast<const wxVariantDataErrorCode*>(value.GetData())->GetValue();
}

SCODE MakeExcelErrorCode(XlCVError error)
{
    return static_cast<SCODE>(
        0x800A0000UL | static_cast<unsigned long>(error));
}

} // namespace

TEST_CASE("wxExcelRange Formula2 spills, Formula uses implicit intersection",
          "[excel][range][formula][spill]")
{
    ExcelTestWorkbook excel;
    REQUIRE(excel.IsReady());
    SkipWithoutDynamicArrays(excel);
    FillNumbers(excel);

    SECTION("Formula2")
    {
        wxExcelRange cell = excel.GetRange("C1");
        cell.SetFormula2("=SORT(A1:A5)");
        REQUIRE(cell);

        CHECK(cell.GetFormula2() == "=SORT(A1:A5)");
        CHECK(cell.GetFormula2R1C1() == "=SORT(RC[-2]:R[4]C[-2])");
        // Formula shows no difference
        CHECK(cell.GetFormula() == "=SORT(A1:A5)");
        CHECK(cell.GetHasSpill().IsTrue());
        CHECK(cell.GetSavedAsArray().IsTrue());
        CHECK(cell.GetHasArray().IsFalse());

        wxExcelRange spilled = cell.GetSpillingToRange();
        REQUIRE(cell);
        REQUIRE(spilled.IsOk_());
        CHECK(spilled.GetAddress() == "$C$1:$C$5");

        wxExcelRangeData values;
        REQUIRE(spilled.GetValues_(values));
        for ( size_t row = 0; row < 5; ++row )
        {
            INFO("row " << row);
            double value = 0;
            CHECK(values.GetCell(row, 0, value));
            CHECK(value == static_cast<double>(row + 1));
        }
    }

    SECTION("Formula")
    {
        wxExcelRange cell = excel.GetRange("C1");
        cell.SetFormula("=SORT(A1:A5)");
        REQUIRE(cell);

        CHECK(cell.GetFormula() == "=SORT(A1:A5)");
        CHECK(cell.GetFormula2() == "=@SORT(A1:A5)");
        CHECK(cell.GetHasSpill().IsFalse());
        CHECK(cell.GetSavedAsArray().IsFalse());

        CHECK(excel.GetRange("C2").GetValue().IsNull());
    }

    SECTION("Formula2Local")
    {
        // no functions or separators, so the formula is the same in the
        // language of the user
        wxExcelRange cell = excel.GetRange("C1");
        cell.SetFormula2Local("=A1:A3*2");
        REQUIRE(cell);

        CHECK(cell.GetFormula2Local() == "=A1:A3*2");
        CHECK(cell.GetSpillingToRange().GetAddress() == "$C$1:$C$3");

        cell.SetFormula2R1C1Local("=RC[-2]:R[1]C[-2]*3");
        REQUIRE(cell);
        CHECK(cell.GetFormula2R1C1Local() == "=RC[-2]:R[1]C[-2]*3");
        CHECK(cell.GetSpillingToRange().GetAddress() == "$C$1:$C$2");
    }
}

TEST_CASE("wxExcelRange spill properties for the cells of a spill range",
          "[excel][range][formula][spill]")
{
    ExcelTestWorkbook excel;
    REQUIRE(excel.IsReady());
    SkipWithoutDynamicArrays(excel);

    excel.GetRange("D1").SetFormula2("=SEQUENCE(3)");
    excel.GetRange("E1").SetValue(1L);

    SECTION("a cell into which the formula spills")
    {
        wxExcelRange cell = excel.GetRange("D2");
        CHECK(cell.GetHasSpill().IsTrue());
        CHECK(cell.GetFormula2().empty());

        wxExcelRange parent = cell.GetSpillParent();
        REQUIRE(cell);
        REQUIRE(parent.IsOk_());
        CHECK(parent.GetAddress() == "$D$1");

        // only the cell with the formula has the range
        wxExcelRange spilled = cell.GetSpillingToRange();
        CHECK(cell);
        CHECK_FALSE(spilled.IsOk_());
    }

    SECTION("a cell outside the spill range")
    {
        wxExcelRange cell = excel.GetRange("E1");
        CHECK(cell.GetHasSpill().IsFalse());
        CHECK(cell.GetSavedAsArray().IsFalse());

        wxExcelRange parent = cell.GetSpillParent();
        CHECK(cell);
        CHECK_FALSE(parent.IsOk_());
    }

    SECTION("a range partly in the spill range")
    {
        wxExcelRange range = excel.GetRange("D1:E1");
        CHECK(range.GetHasSpill().IsDefault());
        CHECK(range.GetSavedAsArray().IsDefault());
    }

    SECTION("a range with more than one cell")
    {
        // SpillingToRange and SpillParent are only for a single cell
        wxLogNull noLog;
        wxExcelRange range = excel.GetRange("D1:D3");

        range.GetSpillingToRange();
        CHECK_FALSE(range);
        range.GetSpillParent();
        CHECK_FALSE(range);
    }

    SECTION("the formulas of all cells")
    {
        // ordered by columns: D1, D2, E1, E2
        wxVariant formulas;
        CHECK(excel.GetRange("D1:E2").GetFormula2(formulas));
        REQUIRE(formulas.IsType("list"));
        REQUIRE(formulas.GetCount() == 4);
        CHECK(formulas[0].GetString() == "=SEQUENCE(3)");
        CHECK(formulas[1].GetString() == "");
        CHECK(formulas[2].GetString() == "1");
        CHECK(formulas[3].GetString() == "");

        wxExcelRangeData data;
        CHECK(excel.GetRange("D1:E2").GetFormulas_(data, true));
        REQUIRE(data.GetRowCount() == 2);
        REQUIRE(data.GetColumnCount() == 2);
        CHECK(data(0, 0).GetString() == "=SEQUENCE(3)");
        CHECK(data(1, 0).GetString() == "");
    }

    SECTION("a reference to the spill range")
    {
        wxExcelRange cell = excel.GetRange("G1");
        cell.SetFormula2("=D1#*10");
        REQUIRE(cell);
        CHECK(cell.GetFormula2() == "=D1#*10");
        CHECK(cell.GetSpillingToRange().GetAddress() == "$G$1:$G$3");
    }
}

TEST_CASE("wxExcelRange formula which cannot spill returns #SPILL!",
          "[excel][range][formula][spill][error]")
{
    ExcelTestWorkbook excel;
    REQUIRE(excel.IsReady());
    SkipWithoutDynamicArrays(excel);

    excel.GetRange("A2").SetValue("in the way");
    wxExcelRange cell = excel.GetRange("A1");
    cell.SetFormula2("=SEQUENCE(3)");
    REQUIRE(cell);

    CHECK(cell.GetHasSpill().IsFalse());
    CHECK(GetExcelErrorCode(cell.GetValue()) == MakeExcelErrorCode(xlErrSpill));

    wxExcelRangeData data;
    REQUIRE(cell.GetValues_(data));
    XlCVError error = xlErrNA;
    CHECK(data.IsCellError(0, 0, &error));
    CHECK(error == xlErrSpill);

    // removing what is in the way lets the formula spill
    excel.GetRange("A2").ClearContents();
    CHECK(cell.GetHasSpill().IsTrue());
    CHECK(cell.GetSpillingToRange().GetAddress() == "$A$1:$A$3");
}

TEST_CASE("wxExcelRange returns the errors of dynamic array functions",
          "[excel][range][formula][spill][error]")
{
    ExcelTestWorkbook excel;
    REQUIRE(excel.IsReady());
    SkipWithoutDynamicArrays(excel);
    FillNumbers(excel);

    // FILTER without any matching value returns #CALC!
    wxExcelRange cell = excel.GetRange("C1");
    cell.SetFormula2("=FILTER(A1:A5,A1:A5>10)");
    REQUIRE(cell);

    wxExcelRangeData data;
    REQUIRE(cell.GetValues_(data));
    XlCVError error = xlErrNA;
    CHECK(data.IsCellError(0, 0, &error));
    CHECK(error == xlErrCalc);
}
