/////////////////////////////////////////////////////////////////////////////
// Purpose:     Catch2 test runner for wxAutoExcel integration tests
// Author:      OpenAI Codex, under the direction of PB
// Generated:   AI-generated using GPT-5.6 Sol with High reasoning effort
// Copyright:   (c) 2026 PB <pbfordev@gmail.com>
// License:     MIT license
/////////////////////////////////////////////////////////////////////////////

#include <catch2/catch_session.hpp>
#include <catch2/reporters/catch_reporter_event_listener.hpp>
#include <catch2/reporters/catch_reporter_registrars.hpp>

#include <exception>
#include <iostream>
#include <stdexcept>
#include <vector>

#include <wx/app.h>
#include <wx/log.h>
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

// Thrown on a wxWidgets assert, so that Catch2 reports it as a failure of
// the current test instead of the default message box blocking the tests.
class TestAssertFailure : public std::runtime_error
{
public:
    explicit TestAssertFailure(const std::string& message)
        : std::runtime_error(message)
    {}
};

bool IsUnwinding()
{
#if __cplusplus >= 201703L || (defined(_MSVC_LANG) && _MSVC_LANG >= 201703L)
    return std::uncaught_exceptions() > 0;
#else
    return std::uncaught_exception();
#endif
}

void TestAssertHandler(const wxString& file, int line, const wxString& func,
                       const wxString& cond, const wxString& msg)
{
    const wxString text = wxString::Format(
        "wxWidgets assert \"%s\" failed in %s() at %s(%d): %s",
        cond, func, file, line, msg);

    // throwing while an exception is already being handled would terminate
    if ( IsUnwinding() )
    {
        std::cerr << text.utf8_string() << std::endl;
        return;
    }

    throw TestAssertFailure(text.utf8_string());
}

// Collects messages logged by wxWidgets and wxAutoExcel, they are shown only
// when an assertion fails, to explain why a call failed.
class TestLog : public wxLog
{
public:
    static std::vector<wxString>& GetMessages()
    {
        static std::vector<wxString> messages;
        return messages;
    }

protected:
    void DoLogTextAtLevel(wxLogLevel WXUNUSED(level), const wxString& msg) override
    {
        GetMessages().push_back(msg);
    }
};

class TestLogListener : public Catch::EventListenerBase
{
public:
    using Catch::EventListenerBase::EventListenerBase;

    void testCaseStarting(Catch::TestCaseInfo const& WXUNUSED(testInfo)) override
    {
        TestLog::GetMessages().clear();
    }

    void assertionEnded(Catch::AssertionStats const& stats) override
    {
        std::vector<wxString>& messages = TestLog::GetMessages();

        if ( !stats.assertionResult.succeeded() && !messages.empty() )
        {
            std::cout << "wxWidgets log messages for the failure below:\n";
            for ( const wxString& message : messages )
                std::cout << "  " << message.utf8_string() << "\n";
            std::cout << std::flush;
        }

        // only show the messages logged since the previous assertion
        messages.clear();
    }
};

CATCH_REGISTER_LISTENER(TestLogListener)

int RunTests(int argc, char* argv[])
{
    wxAutoExcel::wxAutoExcelObjectErrorModeOverrider errorMode{
        wxAutoExcelTests::ErrorMode
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

    wxSetAssertHandler(TestAssertHandler);
    wxLog::DisableTimestamp();
    delete wxLog::SetActiveTarget(new TestLog);

    int result;

    try
    {
        result = RunTests(argc, argv);
    }
    catch ( const std::exception& e )
    {
        std::cerr << "Unexpected exception: " << e.what() << std::endl;
        result = 1;
    }

    wxEntryCleanup();
    return result;
}
