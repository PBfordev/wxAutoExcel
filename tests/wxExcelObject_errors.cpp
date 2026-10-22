/////////////////////////////////////////////////////////////////////////////
// Purpose:     Tests for the error modes of wxExcelObject
// Author:      Claude Opus 5.5 (Anthropic), under the direction of PB
// Copyright:   (c) 2026 PB <pbfordev@gmail.com>
// License:     MIT license
/////////////////////////////////////////////////////////////////////////////

#include <catch2/catch_test_macros.hpp>
#include <catch2/matchers/catch_matchers_string.hpp>

#include "excel_test_fixture.h"

using namespace wxAutoExcel;

namespace
{

// Counts the errors logged while it exists, instead of logging them.
class ErrorLogCounter : public wxLog
{
public:
    ErrorLogCounter()
        : m_previous(wxLog::SetActiveTarget(this))
    {}

    ~ErrorLogCounter()
    {
        wxLog::SetActiveTarget(m_previous);
    }

    size_t GetCount() const { return m_count; }

protected:
    void DoLogRecord(wxLogLevel level, const wxString& WXUNUSED(msg),
                     const wxLogRecordInfo& WXUNUSED(info)) override
    {
        if ( level == wxLOG_Error )
            ++m_count;
    }

private:
    wxLog* m_previous;
    size_t m_count = 0;
};

// Calling a method of an object without IDispatch is an Err_InvalidDispatch
// error, which wxAutoExcel reports without calling Excel.
wxVariant GetValueOfEmptyRange(wxExcelRange& range)
{
    return range.GetValue();
}

} // unnamed namespace

TEST_CASE("wxExcelObject error mode Err_DoNothing reports nothing",
          "[excel][errors]")
{
    wxAutoExcelObjectErrorModeOverrider mode(wxExcelObject::Err_DoNothing);
    ErrorLogCounter log;

    wxExcelRange range;
    CHECK_NOTHROW(GetValueOfEmptyRange(range));
    CHECK_FALSE(range);
    CHECK(log.GetCount() == 0);
}

TEST_CASE("wxExcelObject error mode Err_LogOn* logs the error",
          "[excel][errors]")
{
    wxAutoExcelObjectErrorModeOverrider mode(wxExcelObject::Err_LogOnInvalidDispatch);
    ErrorLogCounter log;

    wxExcelRange range;
    CHECK_NOTHROW(GetValueOfEmptyRange(range));
    CHECK_FALSE(range);
    CHECK(log.GetCount() == 1);
}

TEST_CASE("wxExcelObject error mode Err_ThrowOn* throws wxExcelException",
          "[excel][errors]")
{
    SECTION("invalid dispatch")
    {
        wxAutoExcelObjectErrorModeOverrider mode(wxExcelObject::Err_ThrowOnInvalidDispatch);

        wxExcelRange range;
        CHECK_THROWS_WITH(GetValueOfEmptyRange(range),
                          Catch::Matchers::ContainsSubstring("invalid IDispatch")
                          && Catch::Matchers::ContainsSubstring("Range.Value"));
    }

    SECTION("failed call")
    {
        ExcelTestWorkbook book;
        REQUIRE(book.IsReady());
        wxExcelWorksheet& worksheet = book.GetWorksheet();

        wxAutoExcelObjectErrorModeOverrider mode(wxExcelObject::Err_ThrowOnFailedInvoke, true);
        CHECK_THROWS_AS(worksheet.GetRange("not an address"), wxExcelException);
    }

    SECTION("only for the errors of the set flags")
    {
        wxAutoExcelObjectErrorModeOverrider mode(wxExcelObject::Err_ThrowOnFailedInvoke);

        wxExcelRange range;
        CHECK_NOTHROW(GetValueOfEmptyRange(range));
        CHECK_FALSE(range);
    }
}

#if wxDEBUG_LEVEL
// the assert handler of the tests throws an exception
TEST_CASE("wxExcelObject error mode Err_AssertOn* asserts",
          "[excel][errors]")
{
    SECTION("invalid dispatch")
    {
        wxAutoExcelObjectErrorModeOverrider mode(wxExcelObject::Err_AssertOnInvalidDispatch);

        wxExcelRange range;
        CHECK_THROWS(GetValueOfEmptyRange(range));
    }

    SECTION("invalid argument")
    {
        ExcelTestWorkbook book;
        REQUIRE(book.IsReady());
        wxExcelRange source = book.GetRange("A1");

        wxAutoExcelObjectErrorModeOverrider mode(wxExcelObject::Err_AssertOnInvalidArgument);
        wxExcelRange destination;
        CHECK_THROWS(source.Copy(&destination));
    }

    SECTION("failed call")
    {
        ExcelTestWorkbook book;
        REQUIRE(book.IsReady());
        wxExcelWorksheet& worksheet = book.GetWorksheet();

        wxAutoExcelObjectErrorModeOverrider mode(wxExcelObject::Err_AssertOnFailedInvoke, true);
        CHECK_THROWS(worksheet.GetRange("not an address"));
    }

    SECTION("only for the errors of the set flags")
    {
        ExcelTestWorkbook book;
        REQUIRE(book.IsReady());
        wxExcelWorksheet& worksheet = book.GetWorksheet();

        wxAutoExcelObjectErrorModeOverrider mode(wxExcelObject::Err_AssertOnInvalidDispatch, true);
        wxExcelRange range;
        CHECK_NOTHROW(range = worksheet.GetRange("not an address"));
        CHECK_FALSE(range.IsOk_());
    }
}
#endif // wxDEBUG_LEVEL

TEST_CASE("wxExcelObject reports an invalid argument",
          "[excel][errors]")
{
    ExcelTestWorkbook book;
    REQUIRE(book.IsReady());
    wxExcelRange source = book.GetRange("A1");

    wxAutoExcelObjectErrorModeOverrider mode(wxExcelObject::Err_ThrowOnInvalidArgument);
    wxExcelRange destination;
    CHECK_THROWS_WITH(source.Copy(&destination),
                      Catch::Matchers::ContainsSubstring("invalid argument"));
}

TEST_CASE("wxExcelObject failed calls are reported and make the object false",
          "[excel][errors]")
{
    ExcelTestWorkbook book;
    REQUIRE(book.IsReady());
    wxExcelWorksheet& worksheet = book.GetWorksheet();

    {
        wxAutoExcelObjectErrorModeOverrider mode(wxExcelObject::Err_LogOnFailedInvoke);
        ErrorLogCounter log;

        wxExcelRange range = worksheet.GetRange("not an address");
        CHECK_FALSE(range.IsOk_());
        CHECK_FALSE(worksheet);
        // wxWidgets may log the error of the call too
        CHECK(log.GetCount() >= 1);
    }

    // the next successful call makes the object true again
    CHECK(worksheet.GetName() == worksheet.GetName());
    CHECK(worksheet);
}

TEST_CASE("wxAutoExcelObjectErrorModeOverrider restores the previous mode",
          "[excel][errors]")
{
    const unsigned original = wxExcelObject::GetErrorMode_();

    {
        wxAutoExcelObjectErrorModeOverrider outer(wxExcelObject::Err_DoNothing);
        CHECK(wxExcelObject::GetErrorMode_() == wxExcelObject::Err_DoNothing);

        {
            wxAutoExcelObjectErrorModeOverrider inner(wxExcelObject::Err_ThrowOnFailedInvoke);
            CHECK(wxExcelObject::GetErrorMode_() == wxExcelObject::Err_ThrowOnFailedInvoke);
        }

        CHECK(wxExcelObject::GetErrorMode_() == wxExcelObject::Err_DoNothing);
    }

    CHECK(wxExcelObject::GetErrorMode_() == original);
}

TEST_CASE("wxAutoExcelObjectErrorModeOverrider can suppress logging",
          "[excel][errors]")
{
    ErrorLogCounter log;

    {
        wxAutoExcelObjectErrorModeOverrider mode(wxExcelObject::Err_LogOnInvalidDispatch, true);
        wxExcelRange range;
        GetValueOfEmptyRange(range);
    }
    CHECK(log.GetCount() == 0);

    // logging is enabled again
    {
        wxAutoExcelObjectErrorModeOverrider mode(wxExcelObject::Err_LogOnInvalidDispatch);
        wxExcelRange range;
        GetValueOfEmptyRange(range);
    }
    CHECK(log.GetCount() == 1);
}
