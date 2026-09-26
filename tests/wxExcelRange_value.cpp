/////////////////////////////////////////////////////////////////////////////
// Purpose:     Tests for wxExcelRange value access and conversion
// Author:      OpenAI Codex, under the direction of PB
// Generated:   AI-generated using GPT-5.6 Sol with High reasoning effort
// Copyright:   (c) 2026 PB <pbfordev@gmail.com>
// License:     MIT license
/////////////////////////////////////////////////////////////////////////////

#include <catch2/catch_approx.hpp>
#include <catch2/catch_message.hpp>
#include <catch2/catch_test_macros.hpp>

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

wxVariant MakeExcelError(XlCVError error)
{
    return wxVariant(new wxVariantDataErrorCode(MakeExcelErrorCode(error)));
}

wxVariant MakeCurrencyVariant(wxLongLong_t scaledValue)
{
    CURRENCY currency = {};
    currency.int64 = scaledValue;
    return wxVariant(new wxVariantDataCurrency(currency));
}

SCODE GetExcelErrorCode(const wxVariant& value)
{
    REQUIRE(value.IsType("errorcode"));
    const wxVariantDataErrorCode* const error =
        static_cast<const wxVariantDataErrorCode*>(value.GetData());
    REQUIRE(error != nullptr);
    return error->GetValue();
}

wxVariant SetAndGet(ExcelTestWorkbook& excel, const wxString& address,
                    const wxVariant& value)
{
    wxExcelRange range = excel.GetRange(address);
    range.SetValue(value);
    REQUIRE(range);

    wxVariant result = range.GetValue();
    REQUIRE(range);
    return result;
}

wxVariant MakeVariantList(const double* values, size_t count)
{
    wxVariantList list;
    list.DeleteContents(true);
    for ( size_t i = 0; i < count; ++i )
        list.Append(new wxVariant(values[i]));
    return wxVariant(list);
}

void CheckListValues(wxExcelRange& range, const double* expected, size_t count)
{
    const wxVariant result = range.GetValue();
    REQUIRE(range);
    REQUIRE(result.IsType("list"));
    REQUIRE(result.GetCount() == count);

    for ( size_t i = 0; i < count; ++i )
    {
        INFO("flattened value index: " << i);
        REQUIRE(result[i].IsType("double"));
        CHECK(result[i].GetDouble() == Catch::Approx(expected[i]));
    }
}

void CheckSafeArrayRoundTrip(ExcelTestWorkbook& excel,
                             const wxString& address,
                             long rowCount, long columnCount)
{
    SAFEARRAYBOUND bounds[2] = {};
    bounds[0].lLbound = 0;
    bounds[0].cElements = rowCount;
    bounds[1].lLbound = 0;
    bounds[1].cElements = columnCount;

    wxSafeArray<VT_R8> input;
    REQUIRE(input.Create(bounds, 2));

    long indices[2] = {};
    for ( long row = 0; row < rowCount; ++row )
    {
        indices[0] = row;
        for ( long column = 0; column < columnCount; ++column )
        {
            indices[1] = column;
            const double value = 100.0 * row + column + 0.25;
            REQUIRE(input.SetElement(indices, value));
        }
    }

    wxExcelRange range = excel.GetRange(address);
    range.SetValue(wxVariant(new wxVariantDataSafeArray(input.Detach())));
    REQUIRE(range);
    REQUIRE(range.SetConvertVariantFlags_(wxOleConvertVariant_ReturnSafeArrays));

    const wxVariant result = range.GetValue();
    REQUIRE(range);
    REQUIRE(result.IsType("safearray"));

    wxVariantDataSafeArray* const resultData =
        static_cast<wxVariantDataSafeArray*>(result.GetData());
    REQUIRE(resultData != nullptr);

    wxSafeArray<VT_VARIANT> output;
    REQUIRE(output.Attach(resultData->GetValue()));
    REQUIRE(output.GetDim() == 2);
    CHECK(output.GetCount(1) == static_cast<size_t>(rowCount));
    CHECK(output.GetCount(2) == static_cast<size_t>(columnCount));

    long rowLowerBound = 0;
    long columnLowerBound = 0;
    REQUIRE(output.GetLBound(1, rowLowerBound));
    REQUIRE(output.GetLBound(2, columnLowerBound));

    for ( long row = 0; row < rowCount; ++row )
    {
        indices[0] = rowLowerBound + row;
        for ( long column = 0; column < columnCount; ++column )
        {
            indices[1] = columnLowerBound + column;
            wxVariant value;
            REQUIRE(output.GetElement(indices, value));
            REQUIRE(value.IsType("double"));
            CHECK(value.GetDouble() ==
                  Catch::Approx(100.0 * row + column + 0.25));
        }
    }
}

} // namespace

TEST_CASE("wxExcelRange SetValue and GetValue round-trip scalar values",
          "[excel][range][value]")
{
    ExcelTestWorkbook excel;
    REQUIRE(excel.IsReady());

    SECTION("double")
    {
        const wxVariant result = SetAndGet(excel, "A1", wxVariant(1234.5));

        REQUIRE(result.IsType("double"));
        CHECK(result.GetDouble() == Catch::Approx(1234.5));
    }

    SECTION("string")
    {
        const wxString expected = wxString::FromUTF8("P\xC5\x99\xC3\xADli\xC5\xA1 \xE2\x80\x93 wxAutoExcel");
        const wxVariant result = SetAndGet(excel, "A2", wxVariant(expected));

        REQUIRE(result.IsType("string"));
        CHECK(result.GetString() == expected);
    }

    SECTION("integer")
    {
        const wxVariant result = SetAndGet(excel, "A3", wxVariant(42L));

        // Excel stores worksheet numbers as floating-point values.
        REQUIRE(result.IsType("double"));
        CHECK(result.GetDouble() == Catch::Approx(42.0));
    }

    SECTION("true")
    {
        const wxVariant result = SetAndGet(excel, "A4", wxVariant(true));

        REQUIRE(result.IsType("bool"));
        CHECK(result.GetBool());
    }

    SECTION("false")
    {
        const wxVariant result = SetAndGet(excel, "A5", wxVariant(false));

        REQUIRE(result.IsType("bool"));
        CHECK_FALSE(result.GetBool());
    }

    SECTION("formula-generated error")
    {
        const wxVariant result = SetAndGet(excel, "A6", wxVariant("=NA()"));

        // Excel returns CVErr values as FACILITY_CONTROL error codes.
        CHECK(GetExcelErrorCode(result) == MakeExcelErrorCode(xlErrNA));
    }

    SECTION("date and time")
    {
        const wxDateTime expected(15, wxDateTime::Mar, 2024, 12, 34, 56);
        const wxVariant result = SetAndGet(excel, "A7", wxVariant(expected));

        REQUIRE(result.IsType("datetime"));
        const wxDateTime actual = result.GetDateTime();
        CHECK(actual == expected);
    }

    SECTION("currency")
    {
        CURRENCY expected = {};
        expected.int64 = 12345600; // Currency values have four implied decimals.
        const wxVariant input(new wxVariantDataCurrency(expected));
        const wxVariant result = SetAndGet(excel, "A8", input);

        REQUIRE(result.IsType("currency"));
        const wxVariantDataCurrency* const currency =
            static_cast<const wxVariantDataCurrency*>(result.GetData());
        REQUIRE(currency != nullptr);
        CHECK(currency->GetValue().int64 == expected.int64);
    }
}

TEST_CASE("wxExcelRange reports blank and empty value semantics",
          "[excel][range][value][empty]")
{
    ExcelTestWorkbook excel;
    REQUIRE(excel.IsReady());

    wxExcelRange range = excel.GetRange("A10");

    const wxVariant untouchedValue = range.GetValue();
    REQUIRE(range);
    CHECK(untouchedValue.IsNull());

    range.SetValue(wxVariant("non-empty"));
    REQUIRE(range);
    range.SetValue(wxVariant());
    CHECK_FALSE(range);
    const wxVariant nullValue = range.GetValue();
    REQUIRE(range);
    REQUIRE(nullValue.IsType("string"));
    CHECK(nullValue.GetString() == "non-empty");

    range.SetValue(wxVariant(""));
    REQUIRE(range);
    const wxVariant emptyStringValue = range.GetValue();
    REQUIRE(range);
    CHECK(emptyStringValue.IsNull());

    range.SetValue(wxVariant("clear me"));
    REQUIRE(range);
    range.ClearContents();
    const wxVariant clearedValue = range.GetValue();
    REQUIRE(range);
    CHECK(clearedValue.IsNull());
}

TEST_CASE("wxExcelRange directly round-trips every standard Excel error",
          "[excel][range][value][error]")
{
    ExcelTestWorkbook excel;
    REQUIRE(excel.IsReady());

    struct ErrorCase
    {
        const char* name;
        XlCVError error;
    };

    const ErrorCase errors[] = {
        { "null", xlErrNull },
        { "divide by zero", xlErrDiv0 },
        { "value", xlErrValue },
        { "reference", xlErrRef },
        { "name", xlErrName },
        { "number", xlErrNum },
        { "not available", xlErrNA },
    };

    wxExcelRange range = excel.GetRange("A12");
    for ( const ErrorCase& error : errors )
    {
        CAPTURE(error.name);

        range.SetValue(MakeExcelError(error.error));
        REQUIRE(range);
        const wxVariant result = range.GetValue();
        REQUIRE(range);
        CHECK(GetExcelErrorCode(result) == MakeExcelErrorCode(error.error));
    }
}

TEST_CASE("wxExcelRange Value and Value2 preserve their documented types",
          "[excel][range][value][value2]")
{
    ExcelTestWorkbook excel;
    REQUIRE(excel.IsReady());

    SECTION("date and time")
    {
        wxExcelRange range = excel.GetRange("A14");
        const wxDateTime expected(15, wxDateTime::Mar, 2024, 12, 34, 56);
        range.SetValue(wxVariant(expected));
        REQUIRE(range);

        const wxVariant value = range.GetValue();
        REQUIRE(range);
        REQUIRE(value.IsType("datetime"));
        CHECK(value.GetDateTime() == expected);

        const wxVariant value2 = range.GetValue2();
        REQUIRE(range);
        REQUIRE(value2.IsType("double"));
        CHECK(value2.GetDouble() ==
              Catch::Approx(45366.5242592593).margin(0.0000001));
    }

    SECTION("negative currency with four decimal places")
    {
        wxExcelRange range = excel.GetRange("A15");
        const wxLongLong_t inputScaledValue = -9876543;
        range.SetValue(MakeCurrencyVariant(inputScaledValue));
        REQUIRE(range);

        const wxVariant value = range.GetValue();
        REQUIRE(range);
        REQUIRE(value.IsType("currency"));
        const wxVariantDataCurrency* const currency =
            static_cast<const wxVariantDataCurrency*>(value.GetData());
        REQUIRE(currency != nullptr);
        // Range.Value returns Currency rounded to two decimal places.
        CHECK(currency->GetValue().int64 == -9876500);

        const wxVariant value2 = range.GetValue2();
        REQUIRE(range);
        REQUIRE(value2.IsType("double"));
        CHECK(value2.GetDouble() == Catch::Approx(-987.6543));
    }

    SECTION("SetValue2 and GetValue2")
    {
        wxExcelRange range = excel.GetRange("A16");
        range.SetValue2(wxVariant(42.75));
        REQUIRE(range);

        const wxVariant result = range.GetValue2();
        REQUIRE(range);
        REQUIRE(result.IsType("double"));
        CHECK(result.GetDouble() == Catch::Approx(42.75));
    }
}

TEST_CASE("wxExcelRange values use wxVariantList in a single row",
          "[excel][range][value][variant-list][row]")
{
    ExcelTestWorkbook excel;
    REQUIRE(excel.IsReady());

    const double values[] = {1.25, 2.25, 3.25, 4.25};
    wxExcelRange range = excel.GetRange("B2:E2");
    range.SetValue(MakeVariantList(values, 4));
    REQUIRE(range);

    CheckListValues(range, values, 4);
}

TEST_CASE("wxExcelRange values use wxVariantList in a single column",
          "[excel][range][value][variant-list][column]")
{
    ExcelTestWorkbook excel;
    REQUIRE(excel.IsReady());

    const double values[] = {11.25, 12.25, 13.25, 14.25};
    wxExcelRange range = excel.GetRange("B2:B5");
    // A wxVariantList is one-dimensional and Excel treats it as a row, so
    // populate a column one row at a time as in the bulk-data sample.
    for ( long row = 1; row <= 4; ++row )
    {
        wxExcelRange rowRange = range.GetRows(row);
        REQUIRE(rowRange);
        rowRange.SetValue(MakeVariantList(&values[row - 1], 1));
        REQUIRE(rowRange);
    }

    CheckListValues(range, values, 4);
}

TEST_CASE("wxExcelRange values use wxVariantList across rows and columns",
          "[excel][range][value][variant-list]")
{
    ExcelTestWorkbook excel;
    REQUIRE(excel.IsReady());

    wxExcelRange range = excel.GetRange("B2:D4");

    double nextValue = 1.0;
    for ( long row = 1; row <= 3; ++row )
    {
        wxVariantList values;
        values.DeleteContents(true);
        for ( long column = 1; column <= 3; ++column )
            values.Append(new wxVariant(nextValue++));

        wxExcelRange rowRange = range.GetRows(row);
        REQUIRE(rowRange);
        rowRange.SetValue(wxVariant(values));
        REQUIRE(rowRange);
    }

    // Multidimensional Excel values are flattened by wxWidgets column first.
    const double expected[] = {1.0, 4.0, 7.0, 2.0, 5.0, 8.0, 3.0, 6.0, 9.0};
    CheckListValues(range, expected, 9);
}

TEST_CASE("wxExcelRange values use wxSafeArray in a single row",
          "[excel][range][value][safe-array][row]")
{
    ExcelTestWorkbook excel;
    REQUIRE(excel.IsReady());

    CheckSafeArrayRoundTrip(excel, "F2:I2", 1, 4);
}

TEST_CASE("wxExcelRange values use wxSafeArray in a single column",
          "[excel][range][value][safe-array][column]")
{
    ExcelTestWorkbook excel;
    REQUIRE(excel.IsReady());

    CheckSafeArrayRoundTrip(excel, "F2:F5", 4, 1);
}

TEST_CASE("wxExcelRange values use wxSafeArray across rows and columns",
          "[excel][range][value][safe-array][rectangular]")
{
    ExcelTestWorkbook excel;
    REQUIRE(excel.IsReady());

    CheckSafeArrayRoundTrip(excel, "F2:I4", 3, 4);
}

TEST_CASE("wxExcelRange round-trips mixed types in a rectangular wxSafeArray",
          "[excel][range][value][safe-array][mixed][rectangular]")
{
    ExcelTestWorkbook excel;
    REQUIRE(excel.IsReady());

    const wxDateTime expectedDate(5, wxDateTime::Apr, 2024, 6, 7, 8);
    const wxLongLong_t expectedCurrency = -1234567;
    const wxVariant expected[] = {
        wxVariant(12.5),
        wxVariant("mixed text"),
        wxVariant(true),
        wxVariant(),
        wxVariant(expectedDate),
        MakeCurrencyVariant(expectedCurrency),
        MakeExcelError(xlErrDiv0),
        wxVariant(false),
        wxVariant(wxString::FromUTF8("P\xC5\x99\xC3\xADli\xC5\xA1")),
    };

    SAFEARRAYBOUND bounds[2] = {};
    bounds[0].lLbound = 0;
    bounds[0].cElements = 3;
    bounds[1].lLbound = 0;
    bounds[1].cElements = 3;

    wxSafeArray<VT_VARIANT> input;
    REQUIRE(input.Create(bounds, 2));

    long indices[2] = {};
    for ( long row = 0; row < 3; ++row )
    {
        indices[0] = row;
        for ( long column = 0; column < 3; ++column )
        {
            indices[1] = column;
            REQUIRE(input.SetElement(indices, expected[3 * row + column]));
        }
    }

    wxExcelRange range = excel.GetRange("C12:E14");
    range.SetValue(wxVariant(new wxVariantDataSafeArray(input.Detach())));
    REQUIRE(range);
    REQUIRE(range.SetConvertVariantFlags_(wxOleConvertVariant_ReturnSafeArrays));

    const wxVariant result = range.GetValue();
    REQUIRE(range);
    REQUIRE(result.IsType("safearray"));

    wxVariantDataSafeArray* const resultData =
        static_cast<wxVariantDataSafeArray*>(result.GetData());
    REQUIRE(resultData != nullptr);

    wxSafeArray<VT_VARIANT> output;
    REQUIRE(output.Attach(resultData->GetValue()));
    REQUIRE(output.GetDim() == 2);
    REQUIRE(output.GetCount(1) == 3);
    REQUIRE(output.GetCount(2) == 3);

    long rowLowerBound = 0;
    long columnLowerBound = 0;
    REQUIRE(output.GetLBound(1, rowLowerBound));
    REQUIRE(output.GetLBound(2, columnLowerBound));

    wxVariant actual[9];
    for ( long row = 0; row < 3; ++row )
    {
        indices[0] = rowLowerBound + row;
        for ( long column = 0; column < 3; ++column )
        {
            indices[1] = columnLowerBound + column;
            CAPTURE(row, column);
            REQUIRE(output.GetElement(indices, actual[3 * row + column]));
        }
    }

    REQUIRE(actual[0].IsType("double"));
    CHECK(actual[0].GetDouble() == Catch::Approx(12.5));
    REQUIRE(actual[1].IsType("string"));
    CHECK(actual[1].GetString() == "mixed text");
    REQUIRE(actual[2].IsType("bool"));
    CHECK(actual[2].GetBool());
    CHECK(actual[3].IsNull());
    REQUIRE(actual[4].IsType("datetime"));
    CHECK(actual[4].GetDateTime() == expectedDate);
    REQUIRE(actual[5].IsType("currency"));
    const wxVariantDataCurrency* const currency =
        static_cast<const wxVariantDataCurrency*>(actual[5].GetData());
    REQUIRE(currency != nullptr);
    // Range.Value returns Currency rounded to two decimal places.
    CHECK(currency->GetValue().int64 == -1234600);
    CHECK(GetExcelErrorCode(actual[6]) == MakeExcelErrorCode(xlErrDiv0));
    REQUIRE(actual[7].IsType("bool"));
    CHECK_FALSE(actual[7].GetBool());
    REQUIRE(actual[8].IsType("string"));
    CHECK(actual[8].GetString() ==
          wxString::FromUTF8("P\xC5\x99\xC3\xADli\xC5\xA1"));
}
