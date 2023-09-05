/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2019 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////

/**********************************************************

wxAutoExcel PureWin32 sample shows how to use wxAutoExcel
in an application which does not use wxWidgets as its GUI toolkit.
Most of the code in this file is just the necessary staffolding for
a pure Win32 application, the interesting code is almost
all in UsewxAutoExcel.cpp.

This code has not been extensively tested and so may manifest some
issues, but perhaps it can be still of use.

This example does all its wxAutoExcel calls from a single function,
see UsewxAutoExcel() in usewxAutoExcel.h and UsewxAutoExcel.cpp.
However, the code can be split into three parts: (1) wxWidgets initialization,
(2) many calls to wxAutoExcel from many places, (3) wxWidgets shutdown.
This should be simple to implement, based on the provided example code.

If you do not want to use wxAutoExcel error reporting nor trace logging,
you can suppress all logging with creating a wxLogNull instance at an
appropriate place.

If you want to use wxAutoExcel in a "non-wxWidgets" application, you obviously
still need to do all the things described in docs/install.txt to use wxAutoExcel,
i.e., set the defines, link the libraries...

**********************************************************/

#include <windows.h>
#include <tchar.h>

#include <string>
#include <sstream>

#include "usewxAutoExcel.h"

// An example of MyLogger implementation
// This logger will dump messages sent from wxAutoExcel
// into a rich edit control with hRiched HWND.
// Debug level messages are additionally send to the debug output.
// Base class MyLogger is declared in usewxAutoExcel.h

class RichedLogger : public MyLogger
{
public:
    RichedLogger(HWND hRiched) : m_hRiched(hRiched) {}

    virtual void Log(int level, LPCWSTR message)
    {
        std::wostringstream oss;

        switch ( level )
        {
            case FatalError: oss << L"FATAL ERROR: "; break;
            case Error:      oss << L"ERROR: "; break;
            case Warning:    oss << L"WARNING: "; break;
            case Message:    oss << L"MESSAGE: "; break;
            case Status:     oss << L"STATUS: "; break;
            case Info:       oss << L"INFO: "; break;
            case Debug:      oss << L"DEBUG: "; break;
            case Trace:      oss << L"TRACE: "; break;
            default:         oss << L"<UNKNOWN LOG LEVEL>: ";
        }

        oss << message << std::endl;

        const std::wstring wstr = oss.str();
        const int index = GetWindowTextLengthW(m_hRiched);

        SendMessageW(m_hRiched, EM_SETSEL, (WPARAM)index, (LPARAM)index);
        SendMessageW(m_hRiched, EM_REPLACESEL, 0, (LPARAM)wstr.c_str());
        SendMessage(m_hRiched, WM_VSCROLL, SB_BOTTOM, 0);

        // send the debug messages to the debug output
        // (such as MSVS Output Debug window) as well
        if ( level == Debug || level == Trace )
            OutputDebugStringW(wstr.c_str());
    }
private:
    HWND m_hRiched;
};

RichedLogger* g_richedLoggerInstance = NULL;

// menu command IDs
#define IDM_SHOW_ME    101
#define IDM_EXIT    102

// log child window
HWND g_hLog = NULL;

LRESULT CALLBACK WndProc(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam)
{
    BOOL enableTrace = FALSE;

    switch ( message )
    {
        case WM_COMMAND:
            switch ( LOWORD(wParam) )
            {
                case IDM_SHOW_ME:
#ifndef NDEBUG
                    enableTrace = MessageBox(hWnd, _T("Show wxAutoExcel trace messages?"),
                                   _T("Question"), MB_YESNO | MB_DEFBUTTON2) == IDYES;
#endif
                    // Call the code using wxAutoExcel from here
                    if ( !UsewxAutoExcel(g_richedLoggerInstance, false, enableTrace == TRUE) )
                        MessageBox(hWnd, _T("There was an error when running wxAutoExcel code."), _T("Error"), MB_OK | MB_ICONERROR);
                    break;

                case IDM_EXIT:
                    DestroyWindow(hWnd);
                    break;

                default:
                    return DefWindowProc(hWnd, message, wParam, lParam);
            }
            break;

        case WM_SIZE:
            MoveWindow(g_hLog, 0, 0, LOWORD(lParam), HIWORD(lParam), TRUE);
            break;

        case WM_DESTROY:
            PostQuitMessage(0);
            break;

        default:
            return DefWindowProc(hWnd, message, wParam, lParam);
    }
    return 0;
}

void FatalError(LPCTSTR errorMessage)
{
    MessageBox(NULL, errorMessage, _T("Fatal error"), MB_OK | MB_ICONERROR);
    exit(-1);
}

int WINAPI WinMain(HINSTANCE hInstance, HINSTANCE, LPSTR, int nCmdShow)
{
    // create the app window

    LPCTSTR className = _T("SampleFrame");

    WNDCLASSEX wcex;

    ZeroMemory(&wcex, sizeof(wcex));
    wcex.cbSize         = sizeof(WNDCLASSEX);
    wcex.style          = CS_HREDRAW | CS_VREDRAW;
    wcex.lpfnWndProc    = WndProc;
    wcex.hInstance      = hInstance;
    wcex.hCursor        = LoadCursor(NULL, IDC_ARROW);
    wcex.hbrBackground  = (HBRUSH)(COLOR_WINDOW+1);
    wcex.lpszClassName  = className;

    if ( !RegisterClassEx(&wcex) )
        FatalError(_T("Could not register the class for the application window."));

    HWND hFrame = CreateWindow(className, _T("wxAutoExcel minimal sample for pure Win32 application"), WS_OVERLAPPEDWINDOW,
                    CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, NULL, NULL, hInstance, NULL);
    if ( !hFrame )
      FatalError(_T("Could not create the application window."));

    // add menu

    HMENU hMenubar = CreateMenu();
    HMENU hMenu = CreateMenu();

     if ( !hMenubar || !hMenu )
      FatalError(_T("Could not create the application menu."));

    AppendMenu(hMenu, MF_STRING, IDM_SHOW_ME, _T("&Show me!"));
    AppendMenu(hMenu, MF_STRING, IDM_EXIT, _T("E&xit"));

    AppendMenu(hMenubar, MF_POPUP, (UINT_PTR)hMenu, _T("&Sample"));
    SetMenu(hFrame, hMenubar);

    // add the log child window

    LoadLibrary(_T("riched32.dll"));
    g_hLog = CreateWindowEx(0, _T("RICHEDIT"), NULL,
        WS_CHILD | WS_VISIBLE | WS_VSCROLL | ES_LEFT | ES_MULTILINE | ES_AUTOVSCROLL | ES_READONLY,
        0, 0, 0, 0, hFrame, NULL, hInstance, NULL);
    if ( !g_hLog)
      FatalError(_T("Could not create the log window."));

    SendMessage(g_hLog, WM_SETFONT, (WPARAM)GetStockObject(DEFAULT_GUI_FONT), 0);

    RichedLogger logger(g_hLog);
    g_richedLoggerInstance = &logger;

    ShowWindow(hFrame, nCmdShow);
    UpdateWindow(hFrame);

    // run the message pump

    MSG msg;

    while ( GetMessage(&msg, NULL, 0, 0) )
    {
        TranslateMessage(&msg);
        DispatchMessage(&msg);
    }

    return (int)msg.wParam;
}