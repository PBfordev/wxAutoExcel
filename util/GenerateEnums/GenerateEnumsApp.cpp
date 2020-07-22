#include <wx/wx.h>
#include <wx/config.h>

#include "MainFrame.h"
#include "version.h"

class GenerateEnumsApp : public wxApp
{
public:
    bool OnInit() override;
    int  OnExit() override;
};

bool GenerateEnumsApp::OnInit()
{
    if ( !wxApp::OnInit() )
        return false;

    SetVendorName(APP_VENDOR_STR);
    SetAppName(APP_NAME_STR);

    delete wxConfigBase::Set(new wxConfig(GetAppName(), GetVendorName()));

    MainFrame* frame = new MainFrame();
    frame->Show();

    return true;
}

int GenerateEnumsApp::OnExit()
{
    delete wxConfigBase::Set(nullptr);

    return wxApp::OnExit();
}

wxIMPLEMENT_APP(GenerateEnumsApp);