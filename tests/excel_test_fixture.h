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

class ExcelTestWorkbook
{
public:
    ExcelTestWorkbook()
    {
        application = wxAutoExcel::wxExcelApplication::CreateInstance();
        if ( !application )
            return;

        application.SetDisplayAlerts(false);
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

        if ( application.IsOk_() )
            application.Quit();
        application = wxAutoExcel::wxExcelApplication();
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

    bool IsReady() const
    {
        return ready;
    }

private:
    wxAutoExcel::wxAutoExcelObjectErrorModeOverrider errorMode{
        wxAutoExcel::wxExcelObject::Err_DoNothing, true
    };
    wxAutoExcel::wxExcelApplication application;
    wxAutoExcel::wxExcelWorkbook workbook;
    wxAutoExcel::wxExcelWorksheet worksheet;
    bool ready = false;
};

#endif // WXAUTOEXCEL_TESTS_EXCEL_TEST_FIXTURE_H
