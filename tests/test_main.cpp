/////////////////////////////////////////////////////////////////////////////
// Purpose:     Catch2 test runner for wxAutoExcel integration tests
// Author:      OpenAI Codex, under the direction of PB
// Generated:   AI-generated using GPT-5.6 Sol with High reasoning effort
// Copyright:   (c) 2026 PB <pbfordev@gmail.com>
// License:     MIT license
/////////////////////////////////////////////////////////////////////////////

#include <catch2/catch_session.hpp>

#include <wx/app.h>

int main(int argc, char* argv[])
{
    wxApp::SetInstance(new wxApp());
    if ( !wxEntryStart(argc, argv) )
        return 1;

    const int result = Catch::Session().run(argc, argv);

    wxEntryCleanup();
    return result;
}
