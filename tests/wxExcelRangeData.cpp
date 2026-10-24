/////////////////////////////////////////////////////////////////////////////
// Purpose:     Tests for wxExcelRangeData and the wxExcelRange methods
//              getting and setting the values and formulas of all cells
// Author:      Claude Opus 5.5 (Anthropic), under the direction of PB
// Copyright:   (c) 2026 PB <pbfordev@gmail.com>
// License:     MIT license
/////////////////////////////////////////////////////////////////////////////

#include <catch2/catch_test_macros.hpp>

#include <wx/datetime.h>
#include <wx/msw/ole/oleutils.h>

#include "excel_test_fixture.h"

using namespace wxAutoExcel;

namespace
{

wxVariant MakeExcelError(XlCVError error)
{
    return wxVariant(new wxVariantDataErrorCode(
        static_cast<SCODE>(0x800A0000UL | static_cast<unsigned long>(error))));
}

wxVariant GetValue(ExcelTestWorkbook& book, const wxString& address)
{
    wxExcelRange range = book.GetRange(address);
    const wxVariant value = range.GetValue();
    REQUIRE(range);
    return value;
}

void DestroyArray(wxVariant& variant)
{
    REQUIRE(variant.IsType("safearray"));
    SafeArrayDestroy(wxStaticCastVariantData(variant.GetData(), wxVariantDataSafeArray)->GetValue());
}

// Returns data of 3 rows and 2 columns with values of different types.
wxExcelRangeData MakeMixedData()
{
    wxExcelRangeData data(3, 2);

    data(0, 0) = "text";
    data(0, 1) = 1.5;
    data(1, 0) = true;
    // (1, 1) is empty
    data(2, 0) = MakeExcelError(xlErrDiv0);
    data(2, 1) = -42.0;
    return data;
}

void CheckMixedData(const wxExcelRangeData& data)
{
    REQUIRE(data.GetRowCount() == 3);
    REQUIRE(data.GetColumnCount() == 2);

    REQUIRE(data(0, 0).IsType("string"));
    CHECK(data(0, 0).GetString() == "text");
    REQUIRE(data(0, 1).IsType("double"));
    CHECK(data(0, 1).GetDouble() == 1.5);
    REQUIRE(data(1, 0).IsType("bool"));
    CHECK(data(1, 0).GetBool());
    CHECK(data(1, 1).IsNull());
    REQUIRE(data(2, 0).IsType("errorcode"));
    CHECK(static_cast<const wxVariantDataErrorCode*>(data(2, 0).GetData())->GetValue()
          == static_cast<SCODE>(0x800A0000UL | xlErrDiv0));
    REQUIRE(data(2, 1).IsType("double"));
    CHECK(data(2, 1).GetDouble() == -42.0);
}

} // unnamed namespace

TEST_CASE("wxExcelRangeData stores cells by row and column",
          "[rangedata]")
{
    wxExcelRangeData empty;
    CHECK(empty.IsEmpty());
    CHECK(empty.GetRowCount() == 0);
    CHECK(empty.GetColumnCount() == 0);

    wxExcelRangeData data(2, 3);
    CHECK_FALSE(data.IsEmpty());
    CHECK(data.GetRowCount() == 2);
    CHECK(data.GetColumnCount() == 3);
    CHECK(data(1, 2).IsNull());

    for ( size_t row = 0; row < 2; ++row )
        for ( size_t column = 0; column < 3; ++column )
            data(row, column) = static_cast<double>(row * 10 + column);

    const wxVector<wxVariant> secondRow = data.GetRow(1);
    REQUIRE(secondRow.size() == 3);
    CHECK(secondRow[0].GetDouble() == 10.0);
    CHECK(secondRow[2].GetDouble() == 12.0);

    const wxVector<wxVariant> thirdColumn = data.GetColumn(2);
    REQUIRE(thirdColumn.size() == 2);
    CHECK(thirdColumn[0].GetDouble() == 2.0);
    CHECK(thirdColumn[1].GetDouble() == 12.0);

    // the cells which remain keep their values
    data.Resize(3, 2);
    CHECK(data.GetRowCount() == 3);
    CHECK(data.GetColumnCount() == 2);
    CHECK(data(0, 1).GetDouble() == 1.0);
    CHECK(data(1, 1).GetDouble() == 11.0);
    CHECK(data(2, 0).IsNull());

    data.Resize(0, 5);
    CHECK(data.IsEmpty());
    CHECK(data.GetColumnCount() == 0);

    data.Resize(1, 1);
    data.Clear();
    CHECK(data.IsEmpty());
}

TEST_CASE("wxExcelRangeData converts to and from a SAFEARRAY",
          "[rangedata]")
{
    const wxExcelRangeData data = MakeMixedData();

    wxVariant variant = data.ToVariant();
    REQUIRE(variant.IsType("safearray"));

    wxExcelRangeData converted;
    CHECK(wxExcelRangeData::FromVariant(variant, converted));
    CheckMixedData(converted);
    DestroyArray(variant);

    CHECK(wxExcelRangeData().ToVariant().IsNull());

    // a single value is a single cell
    CHECK(wxExcelRangeData::FromVariant(wxVariant(3.0), converted));
    REQUIRE(converted.GetRowCount() == 1);
    REQUIRE(converted.GetColumnCount() == 1);
    CHECK(converted(0, 0).GetDouble() == 3.0);

    // a list has no dimensions
    wxVariant list;
    list.NullList();
    list.Append(wxVariant(1.0));
    CHECK_FALSE(wxExcelRangeData::FromVariant(list, converted));
    CHECK(converted.IsEmpty());
}

TEST_CASE("wxExcelRange SetValues_ and GetValues_ transfer all cells at once",
          "[excel][range][rangedata]")
{
    ExcelTestWorkbook book;
    REQUIRE(book.IsReady());

    wxExcelRange range = book.GetRange("B2:C4");
    REQUIRE(range.SetValues_(MakeMixedData()));

    // the values are in the cells
    const wxVariant text = GetValue(book, "B2");
    REQUIRE(text.IsType("string"));
    CHECK(text.GetString() == "text");
    const wxVariant number = GetValue(book, "C4");
    REQUIRE(number.IsType("double"));
    CHECK(number.GetDouble() == -42.0);
    CHECK(GetValue(book, "C3").IsNull());

    wxExcelRangeData data;
    REQUIRE(range.GetValues_(data));
    CheckMixedData(data);

    // GetValues_ does not change the flags used by GetValue()
    CHECK(range.GetConvertVariantFlags_() == wxOleConvertVariant_Default);

    // a null value clears the cell
    data(0, 0) = wxVariant();
    REQUIRE(range.SetValues_(data));
    CHECK(GetValue(book, "B2").IsNull());
}

TEST_CASE("wxExcelRange GetValues_ and SetValues_ work with a single cell",
          "[excel][range][rangedata]")
{
    ExcelTestWorkbook book;
    REQUIRE(book.IsReady());

    wxExcelRange cell = book.GetRange("D5");
    wxExcelRangeData data(1, 1);
    data(0, 0) = "single";
    REQUIRE(cell.SetValues_(data));

    wxExcelRangeData result;
    REQUIRE(cell.GetValues_(result));
    REQUIRE(result.GetRowCount() == 1);
    REQUIRE(result.GetColumnCount() == 1);
    REQUIRE(result(0, 0).IsType("string"));
    CHECK(result(0, 0).GetString() == "single");
}

TEST_CASE("wxExcelRange GetValues_ returns dates with Value and numbers with Value2",
          "[excel][range][rangedata]")
{
    ExcelTestWorkbook book;
    REQUIRE(book.IsReady());

    wxExcelRange range = book.GetRange("A1:B1");
    wxExcelRangeData data(1, 2);
    data(0, 0) = wxDateTime(24, wxDateTime::Dec, 2026);
    data(0, 1) = 7.0;
    REQUIRE(range.SetValues_(data, false));

    wxExcelRangeData values;
    REQUIRE(range.GetValues_(values, false));
    REQUIRE(values(0, 0).IsType("datetime"));
    CHECK(values(0, 0).GetDateTime().IsSameDate(wxDateTime(24, wxDateTime::Dec, 2026)));

    wxExcelRangeData values2;
    REQUIRE(range.GetValues_(values2));
    REQUIRE(values2(0, 0).IsType("double"));
    // the serial number of 24 December 2026
    CHECK(values2(0, 0).GetDouble() == 46380.0);
    CHECK(values2(0, 1).GetDouble() == 7.0);
}

TEST_CASE("wxExcelRange SetValues_ fails when the size of the data does not match",
          "[excel][range][rangedata]")
{
    ExcelTestWorkbook book;
    REQUIRE(book.IsReady());

    wxExcelRange range = book.GetRange("A1:B3");
    wxLogNull noLog;

    // more rows, fewer columns, and no cells
    CHECK_FALSE(range.SetValues_(wxExcelRangeData(4, 2)));
    CHECK_FALSE(range);
    CHECK_FALSE(range.SetValues_(wxExcelRangeData(3, 1)));
    CHECK_FALSE(range.SetFormulas_(wxExcelRangeData(2, 2)));
    CHECK_FALSE(range.SetValues_(wxExcelRangeData()));

    // a range with several areas
    wxExcelRangeVector areas;
    areas.push_back(book.GetRange("A1:B3"));
    areas.push_back(book.GetRange("D1:E3"));
    wxExcelRange united = wxAutoExcelTests::GetApplication().Union(areas);
    REQUIRE(united);
    wxExcelRangeData data(3, 2);
    data(0, 0) = 1.0;
    CHECK_FALSE(united.SetValues_(data));
    CHECK_FALSE(united);

    // nothing was set
    CHECK(GetValue(book, "A1").IsNull());
    CHECK(GetValue(book, "D1").IsNull());

    // the matching size works
    CHECK(range.SetValues_(data));
    CHECK(range);
    CHECK(GetValue(book, "A1").GetDouble() == 1.0);
}

TEST_CASE("wxExcelRange SetValues_ reports the mismatch as an invalid argument",
          "[excel][range][rangedata]")
{
    ExcelTestWorkbook book;
    REQUIRE(book.IsReady());

    wxExcelRange range = book.GetRange("A1:B3");
    wxAutoExcelObjectErrorModeOverrider throwOnInvalidArgument(
        wxExcelObject::Err_ThrowOnInvalidArgument);

    CHECK_THROWS_AS(range.SetValues_(wxExcelRangeData(2, 2)), wxExcelException);
}

TEST_CASE("wxExcelRange SetFormulas_ and GetFormulas_ transfer all formulas at once",
          "[excel][range][rangedata][formula]")
{
    ExcelTestWorkbook book;
    REQUIRE(book.IsReady());

    wxExcelRange range = book.GetRange("A1:B2");
    wxExcelRangeData formulas(2, 2);
    formulas(0, 0) = "3";
    formulas(0, 1) = "=A1*2";
    formulas(1, 0) = "=SUM(A1,B1)";
    formulas(1, 1) = "plain text";
    REQUIRE(range.SetFormulas_(formulas));

    CHECK(GetValue(book, "B1").GetDouble() == 6.0);
    CHECK(GetValue(book, "A2").GetDouble() == 9.0);

    wxExcelRangeData result;
    REQUIRE(range.GetFormulas_(result));
    REQUIRE(result.GetRowCount() == 2);
    REQUIRE(result.GetColumnCount() == 2);
    CHECK(result(0, 0).GetString() == "3");
    CHECK(result(0, 1).GetString() == "=A1*2");
    CHECK(result(1, 0).GetString() == "=SUM(A1,B1)");
    CHECK(result(1, 1).GetString() == "plain text");
}

TEST_CASE("wxExcelRange GetValues_ returns the first area of a range with several areas",
          "[excel][range][rangedata]")
{
    ExcelTestWorkbook book;
    REQUIRE(book.IsReady());

    book.GetRange("A1").SetValue(wxVariant(1.0));
    book.GetRange("D1").SetValue(wxVariant(2.0));

    wxExcelRangeVector areas;
    areas.push_back(book.GetRange("A1:B2"));
    areas.push_back(book.GetRange("D1:D3"));
    wxExcelRange united = wxAutoExcelTests::GetApplication().Union(areas);
    REQUIRE(united);

    wxExcelRangeData data;
    REQUIRE(united.GetValues_(data));
    CHECK(data.GetRowCount() == 2);
    CHECK(data.GetColumnCount() == 2);
    CHECK(data(0, 0).GetDouble() == 1.0);
}

TEST_CASE("wxExcelRange transfers a large block of values",
          "[excel][range][rangedata]")
{
    ExcelTestWorkbook book;
    REQUIRE(book.IsReady());

    const size_t rows = 500;
    const size_t columns = 40;
    wxExcelRangeData data(rows, columns);

    for ( size_t row = 0; row < rows; ++row )
        for ( size_t column = 0; column < columns; ++column )
            data(row, column) = static_cast<double>(row * columns + column);

    long rowSize = static_cast<long>(rows);
    long columnSize = static_cast<long>(columns);
    wxExcelRange range = book.GetRange("A1").GetResize(&rowSize, &columnSize);
    REQUIRE(range);
    REQUIRE(range.SetValues_(data));

    wxExcelRangeData result;
    REQUIRE(range.GetValues_(result));
    REQUIRE(result.GetRowCount() == rows);
    REQUIRE(result.GetColumnCount() == columns);
    CHECK(result(0, 0).GetDouble() == 0.0);
    CHECK(result(rows - 1, columns - 1).GetDouble() == static_cast<double>(rows * columns - 1));
    CHECK(result(123, 7).GetDouble() == static_cast<double>(123 * columns + 7));
}

TEST_CASE("wxExcelRangeData gets the cells of the requested types",
          "[rangedata]")
{
    wxExcelRangeData data(1, 8);
    CURRENCY currency;
    currency.int64 = 12345; // 1.2345

    data(0, 0) = 2.5;
    data(0, 1) = -3.0;
    data(0, 2) = 7L;
    data(0, 3) = wxVariant(new wxVariantDataCurrency(currency));
    data(0, 4) = "text";
    data(0, 5) = true;
    data(0, 6) = wxDateTime(24, wxDateTime::Dec, 2026);
    data(0, 7) = MakeExcelError(xlErrNA);

    double d = 0;
    long l = 0;
    wxString s;
    bool b = false;
    wxDateTime dt;

    SECTION("doubles from any number")
    {
        CHECK(data.GetCell(0, 0, d));
        CHECK(d == 2.5);
        CHECK(data.GetCell(0, 2, d));
        CHECK(d == 7.0);
        CHECK(data.GetCell(0, 3, d));
        CHECK(d == 1.2345);
        d = 99.0;
        CHECK_FALSE(data.GetCell(0, 4, d));
        CHECK_FALSE(data.GetCell(0, 5, d));
        CHECK_FALSE(data.GetCell(0, 6, d));
        CHECK_FALSE(data.GetCell(0, 7, d));
        CHECK(d == 99.0);
    }

    SECTION("longs from whole numbers only")
    {
        CHECK(data.GetCell(0, 1, l));
        CHECK(l == -3);
        CHECK(data.GetCell(0, 2, l));
        CHECK(l == 7);
        l = 99;
        CHECK_FALSE(data.GetCell(0, 0, l));
        CHECK_FALSE(data.GetCell(0, 3, l));
        CHECK_FALSE(data.GetCell(0, 4, l));
        CHECK(l == 99);

        wxExcelRangeData big(1, 1);
        big(0, 0) = 1e12;
        CHECK_FALSE(big.GetCell(0, 0, l));
    }

    SECTION("strings, bools, and dates are not converted")
    {
        CHECK(data.GetCell(0, 4, s));
        CHECK(s == "text");
        CHECK_FALSE(data.GetCell(0, 0, s));

        CHECK(data.GetCell(0, 5, b));
        CHECK(b);
        CHECK_FALSE(data.GetCell(0, 0, b));

        CHECK(data.GetCell(0, 6, dt));
        CHECK(dt.IsSameDate(wxDateTime(24, wxDateTime::Dec, 2026)));
        CHECK_FALSE(data.GetCell(0, 0, dt));
    }

    SECTION("empty cells and errors")
    {
        wxExcelRangeData other(1, 2);
        other(0, 1) = 1.0;
        CHECK(other.IsCellEmpty(0, 0));
        CHECK_FALSE(other.IsCellEmpty(0, 1));
        CHECK_FALSE(other.GetCell(0, 0, d));
        CHECK_FALSE(other.GetCell(0, 0, s));

        XlCVError error = xlErrNull;
        CHECK(data.IsCellError(0, 7, &error));
        CHECK(error == xlErrNA);
        CHECK(data.IsCellError(0, 7));
        CHECK_FALSE(data.IsCellError(0, 0, &error));
        CHECK(error == xlErrNA);
    }
}

TEST_CASE("wxExcelRangeData gets the typed values of cells obtained from Excel",
          "[excel][range][rangedata]")
{
    ExcelTestWorkbook book;
    REQUIRE(book.IsReady());

    wxExcelRangeData data(1, 4);
    CURRENCY currency;
    currency.int64 = 25000; // 2.5

    data(0, 0) = wxDateTime(24, wxDateTime::Dec, 2026);
    data(0, 1) = wxVariant(new wxVariantDataCurrency(currency));
    data(0, 2) = 42.0;
    // (0, 3) is empty

    wxExcelRange range = book.GetRange("A1:E1");
    wxExcelRange values = book.GetRange("A1:D1");
    REQUIRE(values.SetValues_(data, false));
    wxExcelRange formula = book.GetRange("E1");
    formula.SetFormula("=1/0");
    REQUIRE(formula);

    wxExcelRangeData result;
    REQUIRE(range.GetValues_(result, false));
    REQUIRE(result.GetColumnCount() == 5);

    wxDateTime date;
    CHECK(result.GetCell(0, 0, date));
    CHECK(date.IsSameDate(wxDateTime(24, wxDateTime::Dec, 2026)));

    double d = 0;
    CHECK(result.GetCell(0, 1, d));
    CHECK(d == 2.5);

    long l = 0;
    CHECK(result.GetCell(0, 2, l));
    CHECK(l == 42);

    CHECK(result.IsCellEmpty(0, 3));

    XlCVError error = xlErrNA;
    CHECK(result.IsCellError(0, 4, &error));
    CHECK(error == xlErrDiv0);

    // with Value2, the date is a number
    wxExcelRangeData result2;
    REQUIRE(range.GetValues_(result2));
    CHECK_FALSE(result2.GetCell(0, 0, date));
    CHECK(result2.GetCell(0, 0, l));
    CHECK(l == 46380);
}
