/////////////////////////////////////////////////////////////////////////////
// Purpose:     Tests for wxExcelApplication
// Author:      Claude Opus 5.5 (Anthropic), under the direction of PB
// Copyright:   (c) 2026 PB <pbfordev@gmail.com>
// License:     MIT license
/////////////////////////////////////////////////////////////////////////////

#include <catch2/catch_test_macros.hpp>

#include "excel_test_fixture.h"

using namespace wxAutoExcel;

namespace
{

// Restores the calculation mode of Excel when the test ends.
class CalculationRestorer
{
public:
    CalculationRestorer()
        : calculation(wxAutoExcelTests::GetApplication().GetCalculation())
    {}

    ~CalculationRestorer()
    {
        wxAutoExcelTests::GetApplication().SetCalculation(calculation);
    }

private:
    XlCalculation calculation;
};

double GetNumber(wxExcelRange range)
{
    const wxVariant value = range.GetValue();
    REQUIRE(range);
    REQUIRE(value.IsType("double"));
    return value.GetDouble();
}

} // unnamed namespace

TEST_CASE("wxExcelApplication manual calculation is done by Calculate",
          "[excel][application][calculation]")
{
    ExcelTestWorkbook book;
    REQUIRE(book.IsReady());
    wxExcelApplication& application = wxAutoExcelTests::GetApplication();
    CalculationRestorer restorer;

    book.GetRange("A1").SetValue(1.0);
    book.GetRange("A2").SetFormula("=A1*2");

    application.SetCalculation(xlCalculationManual);
    CHECK(application.GetCalculation() == xlCalculationManual);

    book.GetRange("A1").SetValue(5.0);
    CHECK(GetNumber(book.GetRange("A2")) == 2.0);

    application.Calculate();
    REQUIRE(application);
    CHECK(GetNumber(book.GetRange("A2")) == 10.0);

    book.GetRange("A1").SetValue(7.0);
    application.SetCalculation(xlCalculationAutomatic);
    CHECK(application.GetCalculation() == xlCalculationAutomatic);
    CHECK(GetNumber(book.GetRange("A2")) == 14.0);
}

TEST_CASE("wxExcelApplication Boolean settings round-trip",
          "[excel][application][settings]")
{
    wxExcelApplication& application = wxAutoExcelTests::GetApplication();

    SECTION("DisplayAlerts")
    {
        const bool saved = application.GetDisplayAlerts();
        REQUIRE(application);
        application.SetDisplayAlerts(!saved);
        CHECK(application.GetDisplayAlerts() == !saved);
        application.SetDisplayAlerts(saved);
        CHECK(application.GetDisplayAlerts() == saved);
    }

    SECTION("ScreenUpdating")
    {
        const bool saved = application.GetScreenUpdating();
        REQUIRE(application);
        application.SetScreenUpdating(!saved);
        CHECK(application.GetScreenUpdating() == !saved);
        application.SetScreenUpdating(saved);
        CHECK(application.GetScreenUpdating() == saved);
    }

    SECTION("Interactive")
    {
        const bool saved = application.GetInteractive();
        REQUIRE(application);
        application.SetInteractive(!saved);
        CHECK(application.GetInteractive() == !saved);
        application.SetInteractive(saved);
        CHECK(application.GetInteractive() == saved);
    }
}

TEST_CASE("wxExcelApplication status bar text is set and reset",
          "[excel][application][settings]")
{
    wxExcelApplication& application = wxAutoExcelTests::GetApplication();

    application.SetStatusBar("wxAutoExcel test");
    CHECK(application.GetStatusBar() == "wxAutoExcel test");

    // without an argument Excel shows its own status bar text again,
    // for which StatusBar returns False
    application.SetStatusBar();
    REQUIRE(application);
}

TEST_CASE("wxExcelApplication combines ranges with Union and Intersect",
          "[excel][application][ranges]")
{
    ExcelTestWorkbook book;
    REQUIRE(book.IsReady());
    wxExcelApplication& application = wxAutoExcelTests::GetApplication();

    wxExcelRangeVector ranges;
    ranges.push_back(book.GetRange("A1:C3"));
    ranges.push_back(book.GetRange("B2:D4"));

    wxExcelRange intersection = application.Intersect(ranges);
    REQUIRE(intersection);
    CHECK(intersection.GetAddress() == "$B$2:$C$3");

    // ranges which overlap are united into the areas as passed
    wxExcelRange united = application.Union(ranges);
    REQUIRE(united);
    CHECK(united.GetAreas().GetCount() == 2);
    CHECK(united.GetCount() == 18);

    // ranges which don't intersect give Nothing, which is not an error
    wxExcelRangeVector disjoint;
    disjoint.push_back(book.GetRange("A1"));
    disjoint.push_back(book.GetRange("C3"));
    wxExcelRange none = application.Intersect(disjoint);
    CHECK_FALSE(none.IsOk_());
    CHECK(application);
}

TEST_CASE("wxExcelApplication converts formulas between reference styles",
          "[excel][application][formula]")
{
    ExcelTestWorkbook book;
    REQUIRE(book.IsReady());
    wxExcelApplication& application = wxAutoExcelTests::GetApplication();

    XlReferenceStyle toStyle = xlR1C1;
    wxExcelRange relativeTo = book.GetRange("C3");
    CHECK(application.ConvertFormula("=A1+$B$2", xlA1, &toStyle, NULL, &relativeTo)
          == "=R[-2]C[-2]+R2C2");
    REQUIRE(application);

    toStyle = xlA1;
    CHECK(application.ConvertFormula("=R[-2]C[-2]+R2C2", xlR1C1, &toStyle, NULL, &relativeTo)
          == "=A1+$B$2");
    REQUIRE(application);
}

TEST_CASE("wxExcelApplication returns its version",
          "[excel][application][version]")
{
    wxExcelApplication& application = wxAutoExcelTests::GetApplication();

    const wxString version = application.GetVersion();
    REQUIRE(application);
    CHECK_FALSE(version.empty());

    double versionNumber = 0;
    REQUIRE(application.GetVersionAsDouble_(versionNumber));
    double parsed = 0;
    REQUIRE(version.ToCDouble(&parsed));
    CHECK(versionNumber == parsed);

    wxExcelApplication::ExcelForWindowsVersions versionEnum;
    REQUIRE(application.GetVersionAsEnum_(versionEnum));
    CHECK(versionEnum == static_cast<int>(versionNumber * 10));

    // the tests need at least Excel 2007, e.g. for the xlsx format
    CHECK(application.IsVersionAtLeast_(wxExcelApplication::evExcel2007));
    CHECK(application.IsVersionAtLeast_(versionEnum));
}

TEST_CASE("wxExcelApplication returns the international settings",
          "[excel][application][international]")
{
    wxExcelApplication& application = wxAutoExcelTests::GetApplication();

    // all the settings, as a list
    const wxVariant all = application.GetInternational();
    REQUIRE(application);
    REQUIRE(all.IsType("list"));
    CHECK(all.GetCount() > static_cast<size_t>(xlListSeparator));

    XlApplicationInternational index = xlDecimalSeparator;
    const wxVariant decimalSeparator = application.GetInternational(&index);
    REQUIRE(application);
    REQUIRE(decimalSeparator.IsType("string"));
    CHECK(decimalSeparator.GetString().length() >= 1);

    // the list separator differs from the decimal separator,
    // so that lists of numbers can be parsed
    const wxString listSeparator = application.GetListSeparator_();
    REQUIRE(application);
    CHECK_FALSE(listSeparator.empty());
    CHECK(listSeparator != decimalSeparator.GetString());
}

TEST_CASE("wxExcelApplication returns the active workbook, sheet, and cell",
          "[excel][application][active]")
{
    ExcelTestWorkbook book;
    REQUIRE(book.IsReady());
    wxExcelApplication& application = wxAutoExcelTests::GetApplication();

    REQUIRE(book.GetWorkbook().Activate());
    REQUIRE(book.GetRange("C4").Activate());

    wxExcelWorkbook workbook = application.GetActiveWorkbook();
    REQUIRE(workbook);
    CHECK(workbook.GetName() == book.GetWorkbook().GetName());

    wxExcelSheet sheet = application.GetActiveSheet();
    REQUIRE(sheet);
    CHECK(sheet.GetName() == book.GetWorksheet().GetName());

    wxExcelRange cell = application.GetActiveCell();
    REQUIRE(cell);
    CHECK(cell.GetAddress() == "$C$4");
}
