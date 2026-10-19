/////////////////////////////////////////////////////////////////////////////
// Purpose:     Shared fixture for wxAutoExcel integration tests
// Author:      OpenAI Codex, under the direction of PB
// Generated:   AI-generated using GPT-5.6 Sol with High reasoning effort
// Copyright:   (c) 2026 PB <pbfordev@gmail.com>
// License:     MIT license
/////////////////////////////////////////////////////////////////////////////

#ifndef WXAUTOEXCEL_TESTS_EXCEL_TEST_FIXTURE_H
#define WXAUTOEXCEL_TESTS_EXCEL_TEST_FIXTURE_H

#include <catch2/catch_test_macros.hpp>

#include <wx/log.h>
#include <wx/wxAutoExcel.h>

#include "catch_wx_string_makers.h"

namespace wxAutoExcelTests
{

wxAutoExcel::wxExcelApplication& GetApplication();

// Log errors without asserting or throwing: the tests check the results and
// the logged messages are shown when an assertion fails.
const unsigned ErrorMode =
    wxAutoExcel::wxExcelObject::Err_LogOnInvalidDispatch
    | wxAutoExcel::wxExcelObject::Err_LogOnInvalidArgument
    | wxAutoExcel::wxExcelObject::Err_LogOnFailedInvoke
    | wxAutoExcel::wxExcelObject::Err_LogOnInvalidReturnType
    | wxAutoExcel::wxExcelObject::Err_LogOnOtherError;

} // namespace wxAutoExcelTests

class ExcelTestWorkbook
{
public:
    ExcelTestWorkbook()
    {
        wxAutoExcel::wxExcelApplication& application =
            wxAutoExcelTests::GetApplication();
        if ( !application )
            return;

        workbook = application.GetWorkbooks().Add();
        if ( !workbook )
            return;

        if ( !workbook.SetAutomationLCID_(
                 wxAutoExcel::wxExcelObject::lcidEnglishUS) )
            return;

        worksheet = workbook.GetWorksheets()[1];
        ready = static_cast<bool>(worksheet);
    }

    ~ExcelTestWorkbook()
    {
        worksheet = wxAutoExcel::wxExcelWorksheet();

        if ( workbook.IsOk_() )
        {
            workbook.SetSaved(true);
            workbook.Close(false);
        }
        workbook = wxAutoExcel::wxExcelWorkbook();
    }

    wxAutoExcel::wxExcelRange GetRange(const wxString& address)
    {
        wxAutoExcel::wxExcelRange range = worksheet.GetRange(address);
        REQUIRE(range);
        return range;
    }

    wxAutoExcel::wxExcelWorksheets GetWorksheets()
    {
        wxAutoExcel::wxExcelWorksheets worksheets = workbook.GetWorksheets();
        REQUIRE(worksheets);
        return worksheets;
    }

    wxAutoExcel::wxExcelWorkbook& GetWorkbook()
    {
        return workbook;
    }

    wxAutoExcel::wxExcelWorksheet& GetWorksheet()
    {
        return worksheet;
    }

    bool IsReady() const
    {
        return ready;
    }

private:
    wxAutoExcel::wxAutoExcelObjectErrorModeOverrider errorMode{
        wxAutoExcelTests::ErrorMode
    };
    wxAutoExcel::wxExcelWorkbook workbook;
    wxAutoExcel::wxExcelWorksheet worksheet;
    bool ready = false;
};

#endif // WXAUTOEXCEL_TESTS_EXCEL_TEST_FIXTURE_H
