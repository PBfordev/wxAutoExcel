/////////////////////////////////////////////////////////////////////////////
// Purpose:     Catch2 test runner for wxAutoExcel integration tests
// Author:      OpenAI Codex, under the direction of PB
// Generated:   AI-generated using GPT-5.6 Sol with High reasoning effort
// Copyright:   (c) 2026 PB <pbfordev@gmail.com>
// License:     MIT license
/////////////////////////////////////////////////////////////////////////////

#include <catch2/catch_session.hpp>

#include <wx/app.h>
#include <wx/wxAutoExcel.h>

#include "excel_test_fixture.h"

namespace wxAutoExcelTests
{

namespace
{

wxAutoExcel::wxExcelApplication* application = NULL;

} // unnamed namespace

wxAutoExcel::wxExcelApplication& GetApplication()
{
    wxASSERT(application != NULL);
    return *application;
}

} // namespace wxAutoExcelTests

int main(int argc, char* argv[])
{
    wxApp::SetInstance(new wxApp());
    if ( !wxEntryStart(argc, argv) )
        return 1;

    wxAutoExcel::wxAutoExcelObjectErrorModeOverrider errorMode{
        wxAutoExcel::wxExcelObject::Err_DoNothing, true
    };
    wxAutoExcel::wxExcelApplication application =
        wxAutoExcel::wxExcelApplication::CreateInstance();
    wxAutoExcelTests::application = &application;

    if ( application )
        application.SetDisplayAlerts(false);

    const int result = Catch::Session().run(argc, argv);

    wxAutoExcelTests::application = NULL;
    if ( application.IsOk_() )
        application.Quit();
    application = wxAutoExcel::wxExcelApplication();

    wxEntryCleanup();
    return result;
}
