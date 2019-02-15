/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////


/**********************************************************

wxAutoExcel Shapes sample focuses on:
- Adding various MS Office Shapes to a worksheet.
- Customising the properties of addes MS Office Shape objects.

**********************************************************/


#include <wx/wx.h>
#include <wx/vector.h>
#include <wx/geometry.h>
#include <wx/iconbndl.h>

#include <wx/wxAutoExcel.h>

#if  !WXAUTOEXCEL_USE_SHAPES
    #error In order to compile this sample, wxAutoExcel has to be built with WXAUTOEXCEL_USE_SHAPES set to 1 in wxAutoExcel_setup.h
#endif

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
: wxFrame(NULL, wxID_ANY, _("wxAutoExcel Shapes sample"))
{
    SetIcons(wxIconBundle("appIcon", NULL));
    
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
        
    if ( !app.IsVersionAtLeast_(wxExcelApplication::evExcel2007) )
    {
        wxMessageBox(_("This sample requires Microsoft Excel 2007 or newer."), "Information");
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

    // Add a heart shape
    wxExcelShape heart = shapes.AddShape(msoShapeHeart, 300, 100, 50, 50);
    
    // set a simple two-color gradient
    wxExcelFillFormat fillFormat = heart.GetFill();     
    
    fillFormat.SetVisible(msoTrue);
    
    colorFormat = fillFormat.GetForeColor();
    colorFormat.SetRGB(*wxRED);
    colorFormat.SetTintAndShade(0.);

    colorFormat = fillFormat.GetBackColor();
    colorFormat.SetRGB(*wxBLACK);
    colorFormat.SetTintAndShade(0.);

    fillFormat.TwoColorGradient(msoGradientHorizontal, 1);        

    // set a red outline
    lineFormat = heart.GetLine();
    lineFormat.SetVisible(msoTrue);
    lineFormat.GetForeColor().SetRGB(*wxRED);
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
