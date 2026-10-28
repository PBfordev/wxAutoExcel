/////////////////////////////////////////////////////////////////////////////
// Purpose:     Program using an installed wxAutoExcel package
// Author:      Claude Opus 5.5 (Anthropic), under the direction of PB
// Copyright:   (c) 2026 PB <pbfordev@gmail.com>
// License:     MIT license
/////////////////////////////////////////////////////////////////////////////

#include <cstdio>

#include <wx/init.h>
#include <wx/utils.h>
#include <wx/wxAutoExcel.h>

// Does not need Excel: it only checks that the program links with the
// wxAutoExcel and wxWidgets libraries and can be run.
int main(int argc, char* argv[])
{
    wxInitializer initializer(argc, argv);
    if ( !initializer.IsOk() )
    {
        std::fprintf(stderr, "Failed to initialize wxWidgets\n");
        return 1;
    }

    wxAutoExcel::wxExcelApplication application;
    if ( application.IsOk_() )
    {
        std::fprintf(stderr, "Default constructed wxExcelApplication is not empty\n");
        return 1;
    }

    std::printf("wxAutoExcel %s with wxWidgets %s\n",
                WXAUTOEXCEL_VERSION_NUM_DOT_STRING, wxVERSION_NUM_DOT_STRING);
    std::printf("OS: %s\n", static_cast<const char*>(wxGetOsDescription().utf8_str()));
    std::printf("CPU architecture: %s (native %s)\n",
                static_cast<const char*>(wxGetCpuArchitectureName().utf8_str()),
                static_cast<const char*>(wxGetNativeCpuArchitectureName().utf8_str()));
    return 0;
}
