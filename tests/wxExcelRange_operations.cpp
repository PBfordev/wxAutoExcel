/////////////////////////////////////////////////////////////////////////////
// Purpose:     Tests for wxExcelRange methods which copy, insert, delete,
//              clear, and navigate ranges
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

double GetNumber(ExcelTestWorkbook& book, const wxString& address)
{
    const wxVariant value = GetValue(book, address);
    REQUIRE(value.IsType("double"));
    return value.GetDouble();
}

wxString GetAddress(wxExcelRange range)
{
    REQUIRE(range);
    const wxString address = range.GetAddress();
    REQUIRE(range);
    return address;
}

bool IsEmpty(ExcelTestWorkbook& book, const wxString& address)
{
    return GetValue(book, address).IsNull();
}

wxString GetInternationalString(XlApplicationInternational index)
{
    wxExcelApplication& application = wxAutoExcelTests::GetApplication();
    const wxVariant value = application.GetInternational(&index);
    REQUIRE(application);
    REQUIRE(value.IsType("string"));
    return value.GetString();
}

// Fills A1:B3 with numbers 1 to 6, row by row.
void FillNumbers(ExcelTestWorkbook& book)
{
    double number = 1.0;
    for ( long row = 1; row <= 3; ++row )
    {
        book.GetRange(wxString::Format("A%ld", row)).SetValue(number++);
        book.GetRange(wxString::Format("B%ld", row)).SetValue(number++);
    }
}

// Empties the clipboard of Excel when the test ends, so that the copied
// range is no longer marked.
class CutCopyModeResetter
{
public:
    ~CutCopyModeResetter()
    {
        wxAutoExcelTests::GetApplication().SetCutCopyMode();
    }
};

} // unnamed namespace

TEST_CASE("wxExcelRange copies values, formulas, and formats to a destination",
          "[excel][range][operations][copy]")
{
    ExcelTestWorkbook book;
    REQUIRE(book.IsReady());

    book.GetRange("A1").SetValue(2.0);
    book.GetRange("A2").SetFormula("=A1*10");
    book.GetRange("A1:A2").GetFont().SetBold(true);

    wxExcelRange destination = book.GetRange("C5");
    REQUIRE(book.GetRange("A1:A2").Copy(&destination));

    CHECK(GetNumber(book, "C5") == 2.0);
    // relative references are adjusted
    CHECK(book.GetRange("C6").GetFormula() == "=C5*10");
    CHECK(GetNumber(book, "C6") == 20.0);
    CHECK(book.GetRange("C6").GetFont().GetBold());
    // the source is unchanged
    CHECK(GetNumber(book, "A1") == 2.0);
}

TEST_CASE("wxExcelRange cuts to a destination",
          "[excel][range][operations][cut]")
{
    ExcelTestWorkbook book;
    REQUIRE(book.IsReady());

    book.GetRange("A1").SetValue(wxVariant("moved"));
    book.GetRange("B1").SetFormula("=A1");

    wxExcelRange destination = book.GetRange("D4");
    REQUIRE(book.GetRange("A1").Cut(&destination));

    CHECK(GetValue(book, "D4").GetString() == "moved");
    CHECK(IsEmpty(book, "A1"));
    // references to the cut cells follow them
    CHECK(book.GetRange("B1").GetFormula() == "=D4");
}

TEST_CASE("wxExcelRange pastes values and transposes from the clipboard",
          "[excel][range][operations][paste-special]")
{
    ExcelTestWorkbook book;
    REQUIRE(book.IsReady());
    CutCopyModeResetter resetter;

    book.GetRange("A1").SetValue(3.0);
    book.GetRange("B1").SetFormula("=A1*2");
    book.GetRange("A1:B1").GetFont().SetItalic(true);

    REQUIRE(book.GetRange("A1:B1").Copy());

    SECTION("values only")
    {
        XlPasteType paste = xlPasteValues;
        REQUIRE(book.GetRange("A3").PasteSpecial(&paste));

        CHECK(GetNumber(book, "A3") == 3.0);
        CHECK(GetNumber(book, "B3") == 6.0);
        CHECK(book.GetRange("B3").GetHasFormula().IsFalse());
        CHECK_FALSE(book.GetRange("B3").GetFont().GetItalic());
    }

    SECTION("transposed")
    {
        XlPasteType paste = xlPasteAll;
        REQUIRE(book.GetRange("D1").PasteSpecial(&paste, NULL, wxDefaultXlTribool, true));

        CHECK(GetNumber(book, "D1") == 3.0);
        CHECK(GetNumber(book, "D2") == 6.0);
        CHECK(book.GetRange("D2").GetFormula() == "=D1*2");
        CHECK(IsEmpty(book, "E1"));
    }

    SECTION("adding to the destination")
    {
        book.GetRange("A5").SetValue(10.0);
        XlPasteType paste = xlPasteValues;
        XlPasteSpecialOperation operation = xlPasteSpecialOperationAdd;
        REQUIRE(book.GetRange("A5").PasteSpecial(&paste, &operation));

        CHECK(GetNumber(book, "A5") == 13.0);
        CHECK(GetNumber(book, "B5") == 6.0);
    }
}

TEST_CASE("wxExcelRange inserts and deletes cells in both directions",
          "[excel][range][operations][insert-delete]")
{
    ExcelTestWorkbook book;
    REQUIRE(book.IsReady());
    FillNumbers(book);

    SECTION("insert shifting down")
    {
        XlInsertShiftDirection shift = xlShiftDown;
        REQUIRE(book.GetRange("A1").Insert(&shift));

        CHECK(IsEmpty(book, "A1"));
        CHECK(GetNumber(book, "A2") == 1.0);
        CHECK(GetNumber(book, "A4") == 5.0);
        CHECK(GetNumber(book, "B1") == 2.0);
    }

    SECTION("insert shifting right")
    {
        XlInsertShiftDirection shift = xlShiftToRight;
        REQUIRE(book.GetRange("A1").Insert(&shift));

        CHECK(IsEmpty(book, "A1"));
        CHECK(GetNumber(book, "B1") == 1.0);
        CHECK(GetNumber(book, "C1") == 2.0);
        CHECK(GetNumber(book, "A2") == 3.0);
    }

    SECTION("delete shifting up")
    {
        XlDeleteShiftDirection shift = xlShiftUp;
        REQUIRE(book.GetRange("A1").Delete(&shift));

        CHECK(GetNumber(book, "A1") == 3.0);
        CHECK(GetNumber(book, "A2") == 5.0);
        CHECK(IsEmpty(book, "A3"));
        CHECK(GetNumber(book, "B1") == 2.0);
    }

    SECTION("delete shifting left")
    {
        XlDeleteShiftDirection shift = xlShiftToLeft;
        REQUIRE(book.GetRange("A1").Delete(&shift));

        CHECK(GetNumber(book, "A1") == 2.0);
        CHECK(IsEmpty(book, "B1"));
        CHECK(GetNumber(book, "A2") == 3.0);
    }

    SECTION("entire rows")
    {
        REQUIRE(book.GetRange("A2").GetEntireRow().Delete());

        CHECK(GetNumber(book, "A2") == 5.0);
        CHECK(GetNumber(book, "B2") == 6.0);

        XlInsertShiftDirection shift = xlShiftDown;
        REQUIRE(book.GetRange("A1").GetEntireRow().Insert(&shift));
        CHECK(IsEmpty(book, "B1"));
        CHECK(GetNumber(book, "B2") == 2.0);
    }
}

TEST_CASE("wxExcelRange clears contents, formats, or both",
          "[excel][range][operations][clear]")
{
    ExcelTestWorkbook book;
    REQUIRE(book.IsReady());

    wxExcelRange range = book.GetRange("A1");
    range.SetValue(1234.5);
    range.SetNumberformat("0.00");
    range.GetFont().SetBold(true);
    REQUIRE(range);

    SECTION("contents")
    {
        REQUIRE(range.ClearContents());
        CHECK(IsEmpty(book, "A1"));
        CHECK(range.GetNumberformat() == "0.00");
        CHECK(range.GetFont().GetBold());
    }

    SECTION("formats")
    {
        REQUIRE(range.ClearFormats());
        CHECK(GetNumber(book, "A1") == 1234.5);
        CHECK(range.GetNumberformat() == "General");
        CHECK_FALSE(range.GetFont().GetBold());
    }

    SECTION("everything")
    {
        REQUIRE(range.Clear());
        CHECK(IsEmpty(book, "A1"));
        CHECK(range.GetNumberformat() == "General");
        CHECK_FALSE(range.GetFont().GetBold());
    }
}

TEST_CASE("wxExcelRange GetText returns the value as displayed",
          "[excel][range][operations][text]")
{
    ExcelTestWorkbook book;
    REQUIRE(book.IsReady());

    // the text uses the separators of Excel, which by default are those of
    // the Windows regional settings, regardless of the automation LCID
    const wxString decimal = GetInternationalString(xlDecimalSeparator);
    const wxString thousands = GetInternationalString(xlThousandsSeparator);

    wxExcelRange range = book.GetRange("A1");
    range.SetValue(0.25);
    range.SetNumberformat("0.0%");
    CHECK(range.GetText() == "25" + decimal + "0%");

    range.SetNumberformat("#,##0.00");
    range.SetValue(1234.5);
    CHECK(range.GetText() == "1" + thousands + "234" + decimal + "50");
    REQUIRE(range);
}

TEST_CASE("wxExcelRange offsets, resizes, and extends to entire rows and columns",
          "[excel][range][operations][navigation]")
{
    ExcelTestWorkbook book;
    REQUIRE(book.IsReady());

    wxExcelRange range = book.GetRange("B2");

    CHECK(GetAddress(range.GetOffset(2, 3)) == "$E$4");
    CHECK(GetAddress(range.GetOffset(-1)) == "$B$1");

    long rows = 3;
    long columns = 2;
    CHECK(GetAddress(range.GetResize(&rows, &columns)) == "$B$2:$C$4");
    CHECK(GetAddress(range.GetResize(&rows)) == "$B$2:$B$4");

    CHECK(GetAddress(range.GetEntireRow()) == "$2:$2");
    CHECK(GetAddress(range.GetEntireColumn()) == "$B:$B");

    wxExcelRange area = book.GetRange("B2:D5");
    CHECK(area.GetRow() == 2);
    CHECK(area.GetColumn() == 2);
    CHECK(area.GetCount() == 12);
    REQUIRE(area);
}

TEST_CASE("wxExcelRange finds the end of data and the current region",
          "[excel][range][operations][navigation]")
{
    ExcelTestWorkbook book;
    REQUIRE(book.IsReady());

    // data in B2:D5 with the header in row 2
    for ( long row = 2; row <= 5; ++row )
        book.GetRange(wxString::Format("B%ld:D%ld", row, row)).SetValue(static_cast<double>(row));

    wxExcelRange start = book.GetRange("B2");
    CHECK(GetAddress(start.GetEnd(xlDown)) == "$B$5");
    CHECK(GetAddress(start.GetEnd(xlToRight)) == "$D$2");
    CHECK(GetAddress(book.GetRange("D5").GetEnd(xlUp)) == "$D$2");
    CHECK(GetAddress(book.GetRange("D5").GetEnd(xlToLeft)) == "$B$5");

    CHECK(GetAddress(book.GetRange("C3").GetCurrentRegion()) == "$B$2:$D$5");
    CHECK(GetAddress(book.GetWorksheet().GetUsedRange()) == "$B$2:$D$5");
}

TEST_CASE("wxExcelRange SpecialCells finds constants, formulas, and blanks",
          "[excel][range][operations][special-cells]")
{
    ExcelTestWorkbook book;
    REQUIRE(book.IsReady());

    book.GetRange("A1").SetValue(1.0);
    book.GetRange("A2").SetValue(wxVariant("text"));
    book.GetRange("A3").SetFormula("=A1+1");
    // A4 is blank
    book.GetRange("A5").SetValue(5.0);

    wxExcelRange range = book.GetRange("A1:A5");

    CHECK(GetAddress(range.SpecialCells(xlCellTypeFormulas)) == "$A$3");
    CHECK(GetAddress(range.SpecialCells(xlCellTypeBlanks)) == "$A$4");
    CHECK(range.SpecialCells(xlCellTypeConstants).GetCount() == 3);
    CHECK(range.SpecialCells(xlCellTypeConstants, static_cast<long>(xlNumbers)).GetCount() == 2);
    CHECK(GetAddress(range.SpecialCells(xlCellTypeConstants, static_cast<long>(xlTextValues))) == "$A$2");
}

TEST_CASE("wxExcelRange fills down, right, and with a series",
          "[excel][range][operations][fill]")
{
    ExcelTestWorkbook book;
    REQUIRE(book.IsReady());

    book.GetRange("A1").SetValue(1.0);
    book.GetRange("B1").SetFormula("=A1*2");

    REQUIRE(book.GetRange("A1:B3").FillDown());
    CHECK(GetNumber(book, "A3") == 1.0);
    CHECK(book.GetRange("B3").GetFormula() == "=A3*2");

    book.GetRange("D1").SetValue(7.0);
    REQUIRE(book.GetRange("D1:F1").FillRight());
    CHECK(GetNumber(book, "F1") == 7.0);

    book.GetRange("H1").SetValue(1.0);
    book.GetRange("H2").SetValue(2.0);
    XlAutoFillType type = xlFillSeries;
    REQUIRE(book.GetRange("H1:H2").AutoFill(book.GetRange("H1:H5"), &type));
    CHECK(GetNumber(book, "H5") == 5.0);
}

TEST_CASE("wxExcelRange AutoFit widens a column to fit its contents",
          "[excel][range][operations][autofit]")
{
    ExcelTestWorkbook book;
    REQUIRE(book.IsReady());

    wxExcelRange column = book.GetRange("C1").GetEntireColumn();
    column.SetColumnWidth(2.0);
    REQUIRE(column);

    book.GetRange("C1").SetValue(wxVariant("A rather long text which needs a wide column"));
    REQUIRE(column.AutoFit());

    const double width = column.GetColumnWidth();
    REQUIRE(column);
    CHECK(width > 20.0);
}
