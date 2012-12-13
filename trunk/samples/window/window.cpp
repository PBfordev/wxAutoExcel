/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Modified by:
// RCS-ID:      $Id: $
// Copyright:   (c) 2012 PB <pb4dev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////


#include <wx/wx.h>

#include <wx/wxAutoExcel.h>

class MyFrame : public wxFrame
{
public:
    MyFrame();    
private:    
    void OnCreateWorksheet(wxCommandEvent& event);
    void OnQuit(wxCommandEvent& event);    
};


class MyApp : public wxApp
{
public:	
    virtual bool OnInit();    
};

using namespace wxAutoExcel;


MyFrame::MyFrame()
: wxFrame(NULL, wxID_ANY, _("wxAutoExcel window sample"))
{
    wxMenu *menu = new wxMenu;
    menu->Append(wxID_NEW, _("&Show me!"));
    menu->Append(wxID_EXIT, _("E&xit"));

    wxMenuBar *menuBar = new wxMenuBar();
    menuBar->Append(menu, _("&Sample"));
    SetMenuBar(menuBar);

    Bind(wxEVT_COMMAND_MENU_SELECTED, &MyFrame::OnCreateWorksheet, this, wxID_NEW);
    Bind(wxEVT_COMMAND_MENU_SELECTED, &MyFrame::OnQuit, this, wxID_EXIT);
}


void MyFrame::OnCreateWorksheet(wxCommandEvent& WXUNUSED(event))
{
    // first create an instance of MS Excel
    wxExcelApplication app = wxExcelApplication::CreateInstance();
    if ( !app ) 
    {
        wxLogError(_("Failed to create an instance of MS Excel application."));
        return;
    }
    app.SetVisible(true); // display MS Excel window

    // add a new workbook
    wxExcelWorkbook workbook = app.GetWorkbooks().Add();    
    if ( !workbook ) 
    {
        wxLogError(_("Failed to create a new workbook."));
        return;
    }

    // get the first worksheet in the newly added workbook
    wxExcelWorksheet worksheet = workbook.GetWorksheets()[1];
    if ( !worksheet )
    {
        wxLogError(_("Failed to obtain worksheet number 1."));
        return;
    }
 
        // Fill the worksheet with some values first    
    wxVariant variant;
    wxExcelRange range;
    long columns;    
    
    range = worksheet.GetRange("A1:BZ1");
    columns = range.GetColumns().GetCount();
    
    app.SetStatusBar("Filling data...");
    app.SetScreenUpdating(false);

    variant.ClearList();
    for (long l = 0; l < columns; l++)
        variant.Append(wxString::Format("Col %ld", l+1));
    range.SetValue(variant);
    range.GetFont().SetBold(true);

    range = worksheet.GetRange("A2:BZ2");    
    for ( int i = 0; i < 2; i++ )
    {
        variant.ClearList();        
        variant.Append(wxString::Format("Row %ld", i+2));
        for (long l = 0; l < columns - 1; l++)
            variant.Append(i+l+2);
        range.SetValue(variant);
        range = range.GetOffset(1);
    }
    worksheet.GetRange("A2:A3").GetFont().SetBold(true);
    
    range = worksheet.GetRange("A2:BZ3");    
    wxExcelRange fillRange = worksheet.GetRange("A2:BZ100");
    range.AutoFill(fillRange, WXAEEP(xlFillSeries)); 
        
    worksheet.GetUsedRange().GetEntireColumn().AutoFit();    
    
    app.SetStatusBar();
    app.SetScreenUpdating(true);
    
    // Obtain the copy of the active window
    wxExcelWindow wnd = app.GetActiveWindow().NewWindow();
    if ( !wnd )
    {
        wxLogError(_("Failed to obtain the active window."));
        return;
    }
    
    // split columns and rows
    wnd.SetSplitColumn(1);
    wnd.SetSplitRow(1);
    wnd.SetFreezePanes(true);

    wnd.SetZoom(200);

    // scroll to left and down
    wnd.SetScrollColumn(columns / 2);
    wnd.SetScrollRow(100 / 2);

    
    app.GetWindows().Arrange(WXAEEP(xlArrangeStyleVertical));    
}

void MyFrame::OnQuit(wxCommandEvent& WXUNUSED(event))
{
    Close(true);
}


bool MyApp::OnInit()
{
    if (!wxApp::OnInit())
        return false;       	
    MyFrame* frame = new MyFrame();
    frame->Show();

    wxLog::AddTraceMask(wxTRACE_AutoExcel);                                  

    return true;
}

IMPLEMENT_APP(MyApp)
