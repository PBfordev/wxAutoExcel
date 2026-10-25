/////////////////////////////////////////////////////////////////////////////
// Purpose:     Tests for wxExcelApplication
// Author:      Claude Opus 5.5 (Anthropic), under the direction of PB
// Copyright:   (c) 2026 PB <pbfordev@gmail.com>
// License:     MIT license
/////////////////////////////////////////////////////////////////////////////

#include <catch2/catch_approx.hpp>
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

TEST_CASE("wxExcelApplication converts formulas between absolute and relative references",
          "[excel][application][formula]")
{
    wxExcelApplication& application = wxAutoExcelTests::GetApplication();
    XlReferenceStyle toStyle = xlA1;

    XlReferenceType toAbsolute = xlAbsolute;
    CHECK(application.ConvertFormula("=A1+B$2", xlA1, &toStyle, &toAbsolute) == "=$A$1+$B$2");

    toAbsolute = xlRelative;
    CHECK(application.ConvertFormula("=$A$1+B$2", xlA1, &toStyle, &toAbsolute) == "=A1+B2");

    toAbsolute = xlAbsRowRelColumn;
    CHECK(application.ConvertFormula("=A1", xlA1, &toStyle, &toAbsolute) == "=A$1");

    toAbsolute = xlRelRowAbsColumn;
    CHECK(application.ConvertFormula("=A1", xlA1, &toStyle, &toAbsolute) == "=$A1");
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

TEST_CASE("wxExcelApplication returns the full build of its EXCEL.EXE",
          "[excel][application][version]")
{
    wxExcelApplication& application = wxAutoExcelTests::GetApplication();

    const wxString fullBuild = application.GetFullBuild_();
    REQUIRE(application);
    INFO("full build: " << fullBuild);

    // e.g. 16.0.20430.20182: Version and Build, followed by the revision
    const wxArrayString parts = wxSplit(fullBuild, '.', '\0');
    REQUIRE(parts.size() == 4);
    for ( size_t i = 0; i < parts.size(); ++i )
    {
        unsigned long number;
        CHECK(parts[i].ToULong(&number));
    }

    const wxString version = application.GetVersion();
    CHECK(parts[0] + "." + parts[1] == version);

    const double build = application.GetBuild();
    REQUIRE(application);
    CHECK(parts[2] == wxString::Format("%.0f", build));
}

TEST_CASE("wxExcelApplication GetFullBuild_ reports an error when it fails",
          "[excel][application][version]")
{
    wxAutoExcelObjectErrorModeOverrider mode(wxExcelObject::Err_ThrowOnInvalidDispatch);

    // an object without Excel can't get the path
    wxExcelApplication application;
    CHECK_THROWS_AS(application.GetFullBuild_(), wxExcelException);

    wxAutoExcelObjectErrorModeOverrider noErrors(wxExcelObject::Err_DoNothing);
    CHECK(application.GetFullBuild_().empty());
    CHECK_FALSE(application);
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

TEST_CASE("wxExcelApplication converts centimeters and inches to points",
          "[excel][application]")
{
    wxExcelApplication& application = wxAutoExcelTests::GetApplication();

    CHECK(application.CentimetersToPoints(2.54) == Catch::Approx(72.0));
    CHECK(application.InchesToPoints(0.5) == Catch::Approx(36.0));
    REQUIRE(application);
}

TEST_CASE("wxExcelApplication recalculates formulas with manual calculation",
          "[excel][application][calculation]")
{
    ExcelTestWorkbook book;
    REQUIRE(book.IsReady());
    wxExcelApplication& application = wxAutoExcelTests::GetApplication();

    wxExcelRange input = book.GetRange("A1");
    wxExcelRange result = book.GetRange("B1");
    input.SetValue(wxVariant(1L));
    result.SetFormula("=A1*2");
    REQUIRE(result);

    const XlCalculation calculation = application.GetCalculation();
    REQUIRE(application);
    application.SetCalculation(xlCalculationManual);
    REQUIRE(application);

    input.SetValue(wxVariant(5L));
    CHECK(result.GetValue().GetDouble() == 2.0);

    application.CalculateFull();
    REQUIRE(application);
    CHECK(result.GetValue().GetDouble() == 10.0);

    input.SetValue(wxVariant(6L));
    application.CalculateFullRebuild();
    REQUIRE(application);
    CHECK(result.GetValue().GetDouble() == 12.0);

    application.SetCalculation(calculation);
    REQUIRE(application);
}

TEST_CASE("wxExcelApplication adds and deletes a custom list",
          "[excel][application][customlists]")
{
    wxExcelApplication& application = wxAutoExcelTests::GetApplication();

    const long countBefore = application.GetCustomListCount();
    REQUIRE(application);

    wxArrayString list;
    list.push_back("wxAutoExcelAlpha");
    list.push_back("wxAutoExcelBeta");
    list.push_back("wxAutoExcelGamma");
    REQUIRE(application.AddCustomList(list));
    CHECK(application.GetCustomListCount() == countBefore + 1);

    const long listNum = application.GetCustomListNum(list);
    REQUIRE(application);
    CHECK(listNum > 0);

    const wxArrayString contents = application.GetCustomListContents(listNum);
    REQUIRE(application);
    CHECK(contents == list);

    REQUIRE(application.DeleteCustomList(listNum));
    CHECK(application.GetCustomListCount() == countBefore);
}
