/////////////////////////////////////////////////////////////////////////////
// Purpose:     Tests for wxExcelRange formula values
// Author:      OpenAI Codex, under the direction of PB
// Generated:   AI-generated using GPT-5.6 Sol with High reasoning effort
// Copyright:   (c) 2026 PB <pbfordev@gmail.com>
// License:     MIT license
/////////////////////////////////////////////////////////////////////////////

#include <catch2/catch_approx.hpp>
#include <catch2/catch_test_macros.hpp>

#include <wx/msw/ole/oleutils.h>

#include "excel_test_fixture.h"

using namespace wxAutoExcel;

namespace
{

SCODE MakeExcelErrorCode(XlCVError error)
{
    return static_cast<SCODE>(
        0x800A0000UL | static_cast<unsigned long>(error));
}

} // namespace

TEST_CASE("wxExcelRange A1 formula returns its calculated value",
          "[excel][range][formula]")
{
    ExcelTestWorkbook excel;
    REQUIRE(excel.IsReady());

    wxExcelRange range = excel.GetRange("B2");
    range.SetFormula("=SUM(1,2,3)");
    REQUIRE(range);

    CHECK(range.GetFormula() == "=SUM(1,2,3)");
    REQUIRE(range);
    CHECK(range.GetHasFormula().IsTrue());
    REQUIRE(range);

    const wxVariant value = range.GetValue();
    REQUIRE(range);
    REQUIRE(value.IsType("double"));
    CHECK(value.GetDouble() == Catch::Approx(6.0));
}

TEST_CASE("wxExcelRange R1C1 formula resolves relative references",
          "[excel][range][formula][r1c1]")
{
    ExcelTestWorkbook excel;
    REQUIRE(excel.IsReady());

    wxExcelRange input = excel.GetRange("A3");
    input.SetValue(10.0);
    REQUIRE(input);

    wxExcelRange formula = excel.GetRange("B3");
    formula.SetFormulaR1C1("=RC[-1]*3");
    REQUIRE(formula);

    CHECK(formula.GetFormulaR1C1() == "=RC[-1]*3");
    REQUIRE(formula);
    CHECK(formula.GetFormula() == "=A3*3");
    REQUIRE(formula);

    const wxVariant value = formula.GetValue();
    REQUIRE(formula);
    REQUIRE(value.IsType("double"));
    CHECK(value.GetDouble() == Catch::Approx(30.0));
}

TEST_CASE("wxExcelRange HasFormula distinguishes formulas and constants",
          "[excel][range][formula][has-formula]")
{
    ExcelTestWorkbook excel;
    REQUIRE(excel.IsReady());

    wxExcelRange formula = excel.GetRange("A5");
    formula.SetFormula("=2+2");
    REQUIRE(formula);

    wxExcelRange constant = excel.GetRange("A6");
    constant.SetValue(4.0);
    REQUIRE(constant);

    CHECK(formula.GetHasFormula().IsTrue());
    REQUIRE(formula);
    CHECK(constant.GetHasFormula().IsFalse());
    REQUIRE(constant);

    wxExcelRange mixedRange = excel.GetRange("A5:A6");
    CHECK(mixedRange.GetHasFormula().IsDefault());
    REQUIRE(mixedRange);
}

TEST_CASE("wxExcelRange array formula calculates every result cell",
          "[excel][range][formula][array]")
{
    ExcelTestWorkbook excel;
    REQUIRE(excel.IsReady());

    for ( long row = 0; row < 3; ++row )
    {
        wxExcelRange input = excel.GetRange(
            wxString::Format("A%ld", row + 8));
        input.SetValue(static_cast<double>(row + 1));
        REQUIRE(input);
    }

    wxExcelRange formulaRange = excel.GetRange("B8:B10");
    formulaRange.SetFormulaArray("=A8:A10*2");
    REQUIRE(formulaRange);

    CHECK(formulaRange.GetFormulaArray() == "=A8:A10*2");
    REQUIRE(formulaRange);
    CHECK(formulaRange.GetHasArray().IsTrue());
    REQUIRE(formulaRange);

    const wxVariant values = formulaRange.GetValue();
    REQUIRE(formulaRange);
    REQUIRE(values.IsType("list"));
    REQUIRE(values.GetCount() == 3);
    for ( size_t i = 0; i < 3; ++i )
    {
        REQUIRE(values[i].IsType("double"));
        CHECK(values[i].GetDouble() ==
              Catch::Approx(2.0 * static_cast<double>(i + 1)));
    }
}

TEST_CASE("wxExcelRange formula errors are returned as error values",
          "[excel][range][formula][error]")
{
    ExcelTestWorkbook excel;
    REQUIRE(excel.IsReady());

    wxExcelRange range = excel.GetRange("B12");
    range.SetFormula("=1/0");
    REQUIRE(range);

    CHECK(range.GetHasFormula().IsTrue());
    REQUIRE(range);

    const wxVariant value = range.GetValue();
    REQUIRE(range);
    REQUIRE(value.IsType("errorcode"));
    const wxVariantDataErrorCode* const error =
        static_cast<const wxVariantDataErrorCode*>(value.GetData());
    REQUIRE(error != nullptr);
    CHECK(error->GetValue() == MakeExcelErrorCode(xlErrDiv0));
}
