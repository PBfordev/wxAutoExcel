/////////////////////////////////////////////////////////////////////////////
// Purpose:     Tests for wxExcelRange formula values
// Author:      OpenAI Codex, under the direction of PB
// Generated:   AI-generated using GPT-5.6 Sol with High reasoning effort
// Copyright:   (c) 2026 PB <pbfordev@gmail.com>
// License:     MIT license
/////////////////////////////////////////////////////////////////////////////

#include <catch2/catch_approx.hpp>
#include <catch2/catch_test_macros.hpp>

#include <vector>

#include <wx/msw/ole/oleutils.h>
#include <wx/msw/ole/safearray.h>

#include "excel_test_fixture.h"

using namespace wxAutoExcel;

namespace
{

SCODE MakeExcelErrorCode(XlCVError error)
{
    return static_cast<SCODE>(
        0x800A0000UL | static_cast<unsigned long>(error));
}

// Checks that formulas is a list with the expected strings.
void CheckFormulaList(const wxVariant& formulas,
                      const std::vector<wxString>& expected)
{
    REQUIRE(formulas.IsType("list"));
    REQUIRE(formulas.GetCount() == expected.size());
    for ( size_t i = 0; i < expected.size(); ++i )
    {
        INFO("formula " << i);
        REQUIRE(formulas[i].IsType("string"));
        CHECK(formulas[i].GetString() == expected[i]);
    }
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

TEST_CASE("wxExcelRange formula getters return the formulas of all cells",
          "[excel][range][formula][multi-cell]")
{
    ExcelTestWorkbook excel;
    REQUIRE(excel.IsReady());

    // A1:B2 has formulas in A1 and B1, a constant in A2 and B2 is empty;
    // the formulas use no functions or separators, so that they are the
    // same in the language of the user
    excel.GetRange("A1").SetFormula("=C1*2");
    excel.GetRange("B1").SetFormula("=$C$1+1");
    excel.GetRange("A2").SetValue(5L);
    wxExcelRange range = excel.GetRange("A1:B2");
    REQUIRE(range);

    // the formulas are ordered by columns: A1, A2, B1, B2
    wxVariant formulas;
    CHECK(range.GetFormula(formulas));
    CheckFormulaList(formulas, { "=C1*2", "5", "=$C$1+1", "" });

    formulas.Clear();
    CHECK(range.GetFormulaLocal(formulas));
    CheckFormulaList(formulas, { "=C1*2", "5", "=$C$1+1", "" });

    formulas.Clear();
    CHECK(range.GetFormulaR1C1(formulas));
    CheckFormulaList(formulas, { "=RC[2]*2", "5", "=R1C3+1", "" });

    formulas.Clear();
    CHECK(range.GetFormulaR1C1Local(formulas));
    REQUIRE(formulas.IsType("list"));
    CHECK(formulas.GetCount() == 4);

    // a single cell has just a string
    wxExcelRange cell = excel.GetRange("A1");
    formulas.Clear();
    CHECK(cell.GetFormula(formulas));
    REQUIRE(formulas.IsType("string"));
    CHECK(formulas.GetString() == "=C1*2");

    // the getters returning a string can't return the formulas of more cells
    {
        wxLogNull noLog;
        wxAutoExcelObjectErrorModeOverrider noErrors(wxExcelObject::Err_DoNothing);

        CHECK(range.GetFormula().empty());
        CHECK_FALSE(range);
        CHECK(range.GetFormulaR1C1().empty());
        CHECK_FALSE(range);
    }
}

TEST_CASE("wxExcelRange formula getter can return a SAFEARRAY",
          "[excel][range][formula][multi-cell][safearray]")
{
    ExcelTestWorkbook excel;
    REQUIRE(excel.IsReady());

    excel.GetRange("A1:C2").SetFormula("=D1+1");
    wxExcelRange range = excel.GetRange("A1:C2");
    REQUIRE(range);

    range.SetConvertVariantFlags_(wxOleConvertVariant_ReturnSafeArrays);
    wxVariant formulas;
    REQUIRE(range.GetFormula(formulas));
    REQUIRE(formulas.IsType("safearray"));

    wxSafeArray<VT_VARIANT> safeArray;
    wxVariantDataSafeArray* const data =
        wxStaticCastVariantData(formulas.GetData(), wxVariantDataSafeArray);
    REQUIRE(safeArray.Attach(data->GetValue()));
    REQUIRE(safeArray.GetDim() == 2);
    CHECK(safeArray.GetCount(1) == 2); // rows
    CHECK(safeArray.GetCount(2) == 3); // columns

    // Excel arrays start at 1, the second cell of the second row is B2
    long indices[2] = { 2, 2 };
    wxVariant formula;
    REQUIRE(safeArray.GetElement(indices, formula));
    REQUIRE(formula.IsType("string"));
    CHECK(formula.GetString() == "=E2+1");
}

TEST_CASE("wxExcelRange formula setters accept the formulas of all cells",
          "[excel][range][formula][multi-cell][set]")
{
    ExcelTestWorkbook excel;
    REQUIRE(excel.IsReady());

    // a string sets the same formula with relative references adjusted
    wxExcelRange column = excel.GetRange("A1:A3");
    column.SetFormula("=B1*2");
    REQUIRE(column);
    wxVariant formulas;
    CHECK(column.GetFormula(formulas));
    CheckFormulaList(formulas, { "=B1*2", "=B2*2", "=B3*2" });

    // a list sets a formula for each cell of a single row
    wxVariant rowFormulas;
    rowFormulas.NullList();
    rowFormulas.Append(wxVariant("=1+1"));
    rowFormulas.Append(wxVariant("=2+2"));
    rowFormulas.Append(wxVariant("text"));
    wxExcelRange row = excel.GetRange("C1:E1");
    row.SetFormula(rowFormulas);
    REQUIRE(row);
    formulas.Clear();
    CHECK(row.GetFormula(formulas));
    CheckFormulaList(formulas, { "=1+1", "=2+2", "text" });

    // a two-dimensional SAFEARRAY sets a formula for each cell of more rows
    SAFEARRAYBOUND bounds[2];
    bounds[0].lLbound = 0;
    bounds[0].cElements = 2; // rows
    bounds[1].lLbound = 0;
    bounds[1].cElements = 2; // columns
    wxSafeArray<VT_VARIANT> safeArray;
    REQUIRE(safeArray.Create(bounds, 2));
    const char* const elements[2][2] = { { "=RC[2]", "=R1C1" }, { "=10", "=RC[-1]+1" } };
    for ( long r = 0; r < 2; ++r )
    {
        for ( long c = 0; c < 2; ++c )
        {
            long indices[2] = { r, c };
            REQUIRE(safeArray.SetElement(indices, wxVariant(elements[r][c])));
        }
    }
    wxExcelRange block = excel.GetRange("G1:H2");
    block.SetFormulaR1C1(wxVariant(new wxVariantDataSafeArray(safeArray.Detach())));
    REQUIRE(block);
    formulas.Clear();
    CHECK(block.GetFormula(formulas));
    CheckFormulaList(formulas, { "=I1", "=10", "=$A$1", "=G2+1" });

    // a null value clears the cells
    block.SetFormula(wxVariant());
    REQUIRE(block);
    formulas.Clear();
    CHECK(block.GetFormula(formulas));
    CheckFormulaList(formulas, { "", "", "", "" });
}
