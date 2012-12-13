/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Modified by:
// RCS-ID:      $Id: $
// Copyright:   (c) 2012 PB <pb4dev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////


#include <wx/wx.h>
#include <wx/vector.h>
#include <wx/geometry.h>

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

    wxString versionS(app.GetVersion());
    unsigned long version;
    
    versionS.ToULong(&version);
    if ( version < 12 )
    {
        wxMessageBox("This sample requires Microsoft Excel 2007 or newer.", "Information");
        app.Quit();
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

    app.GetActiveWindow().SetDisplayGridlines(false);

    wxExcelShapes shapes = worksheet.GetShapes();
    if ( !shapes )
    {
        wxLogError(_("Failed to obtain shapes collection from worksheet number 1."));
        return;
    }

    wxExcelShape rect1, diamond;
    wxExcelColorFormat colorFormat;

    rect1 = shapes.AddShape(msoShapeRectangle, 5, 5, 50, 20);    
    rect1.GetFill().GetForeColor().SetRGB(*wxRED);
    diamond = shapes.AddShape(msoShapeDiamond, 100, 60, 50, 50);
    diamond.GetFill().GetForeColor().SetRGB(*wxGREEN);

    wxExcelShape connector = shapes.AddConnector(msoConnectorStraight, 0, 0, 5, 5);
    wxExcelConnectorFormat connectorFormat = connector.GetConnectorFormat();
    wxExcelLineFormat lineFormat = connector.GetLine();
    lineFormat.GetForeColor().SetRGB(*wxBLUE);
    lineFormat.SetWeight(3);
    lineFormat.SetDashStyle(msoLineRoundDot);
    lineFormat.SetBeginArrowheadStyle(msoArrowheadTriangle);
    lineFormat.SetEndArrowheadStyle(msoArrowheadTriangle);
    connectorFormat.BeginConnect(rect1, 1);
    connectorFormat.EndConnect(diamond, 1);
    
    wxVector<wxString> names;

    names.push_back(rect1.GetName());
    names.push_back(diamond.GetName());
    names.push_back(connector.GetName());

    wxExcelShapeRange range = shapes.GetRange(names);
    if ( range )
    {
        range.RerouteConnections();
        range.Group();        
    }
    
    wxVector<wxPoint2DDouble> points;    
    wxPoint2DDouble point;

    for (size_t i = 0; i < 21; i++)
    {
        point.m_x = 5 + (i * 8); 
        if ( i % 2 )
            point.m_y = 120 + i;
        else
            point.m_y = 150;
        points.push_back(point);
    }
    point.m_x = 5; point.m_y = 150;
    points.push_back(point);
    wxExcelShape s = shapes.AddPolyline(points);
    s.GetVertices();

    shapes.AddTextEffect(msoTextEffect6, "WordArt", "Arial", 50, msoTrue, msoTrue, 150, 10);

    wxExcelShape textBox = shapes.AddTextbox(msoTextOrientationHorizontal, 160, 80, 100, 50);
    wxExcelTextFrame textFrame = textBox.GetTextFrame();    
    textFrame.SetHorizontalAlignment(xlHAlignCenter);
    textFrame.SetVerticalAlignment(xlVAlignCenter);
    textFrame.Characters().SetText("This is\na Textbox");
    textFrame.Characters().GetFont().SetBold(true);
    textFrame.Characters().GetFont().SetColor(*wxBLUE);

    // wxExcelShape donut = shapes.AddShape(msoShapeDonut, 5, 5, 50, 50);
    // wxExcelShape callout = shapes.AddCallout(msoCalloutOne, 80, 80, 50, 50);    
    
    //wxVector<wxPoint2DDouble> points = rectangle.GetVertices();
    //
    //for ( size_t i = 0; i < points.size(); i++ )
    //{
    //    wxLogDebug("x = %f, y = %f", points[i].m_x, points[i].m_y);
    //}
    
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
