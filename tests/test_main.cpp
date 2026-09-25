/////////////////////////////////////////////////////////////////////////////
// Purpose:     Catch2 test runner for wxAutoExcel integration tests
// Author:      OpenAI Codex, under the direction of PB
// Generated:   AI-generated using GPT-5.6 Sol with High reasoning effort
// Copyright:   (c) 2026 PB <pbfordev@gmail.com>
// License:     MIT license
/////////////////////////////////////////////////////////////////////////////

#include <catch2/catch_session.hpp>

#include <iostream>

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

namespace
{

int RunTests(int argc, char* argv[])
{
    wxAutoExcel::wxAutoExcelObjectErrorModeOverrider errorMode{
        wxAutoExcel::wxExcelObject::Err_DoNothing, true
    };
    wxAutoExcel::wxExcelApplication application =
        wxAutoExcel::wxExcelApplication::CreateInstance();

    if ( !application )
    {
        std::cerr
            << "Unable to create a Microsoft Excel application instance.\n"
            << "Verify that Microsoft Excel is installed and available for "
               "OLE automation.\n";
        return 1;
    }

    application.SetDisplayAlerts(false);
    wxAutoExcelTests::application = &application;

    const int result = Catch::Session().run(argc, argv);

    wxAutoExcelTests::application = NULL;
    if ( application.IsOk_() )
        application.Quit();

    return result;
}

} // unnamed namespace

int main(int argc, char* argv[])
{
    wxApp::SetInstance(new wxApp());
    if ( !wxEntryStart(argc, argv) )
        return 1;

    const int result = RunTests(argc, argv);

    wxEntryCleanup();
    return result;
}
