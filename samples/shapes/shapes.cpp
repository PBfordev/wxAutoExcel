/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// License:     MIT license
/////////////////////////////////////////////////////////////////////////////


/**********************************************************

wxAutoExcel Shapes sample focuses on:
- Drawing a flowchart: adding shapes with text, aligning
  and distributing them, connecting them with connectors,
  and grouping them.
- Customising the properties of added MS Office Shape objects.

The sample requires Microsoft Excel 2007 or newer.

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

    Bind(wxEVT_MENU, &MyFrame::OnCreateWorksheet, this, wxID_NEW);
    Bind(wxEVT_MENU, &MyFrame::OnQuit, this, wxID_EXIT);
}

namespace
{

// The connection sites of the rectangular shapes and of the flowchart
// shapes used here, they are numbered counterclockwise from the top.
enum
{
    SiteTop = 1,
    SiteLeft,
    SiteBottom,
    SiteRight
};

// Adds a flowchart step: a shape of the given type with the text centered in it.
wxExcelShape AddStep(wxExcelShapes& shapes, MsoAutoShapeType type, const wxString& text,
                     const wxColour& color, double top, double width = 130, double height = 45)
{
    wxExcelShape shape = shapes.AddShape(type, 30, top, width, height);
    if ( !shape )
        return shape;

    shape.GetFill().GetForeColor().SetRGB(color);
    shape.GetLine().SetVisible(msoFalse);

    // TextFrame2 is the text API of Excel 2007 and newer,
    // it has more options than the older TextFrame
    wxExcelTextFrame2 textFrame = shape.GetTextFrame2();
    textFrame.SetVerticalAnchor(msoAnchorMiddle);
    textFrame.SetHorizontalAnchor(msoAnchorCenter);

    wxExcelTextRange2 textRange = textFrame.GetTextRange();
    textRange.SetText(text);

    wxExcelFont2 font = textRange.GetFont();
    font.SetBold(msoTrue);
    font.GetFill().GetForeColor().SetRGB(*wxWHITE);

    return shape;
}

// Connects the given sites of two shapes with an arrow,
// the connector moves together with the connected shapes.
wxExcelShape Connect(wxExcelShapes& shapes, wxExcelShape from, long fromSite,
                     wxExcelShape to, long toSite)
{
    wxExcelShape connector = shapes.AddConnector(msoConnectorElbow, 0, 0, 10, 10);
    if ( !connector )
        return connector;

    wxExcelConnectorFormat connectorFormat = connector.GetConnectorFormat();
    connectorFormat.BeginConnect(from, fromSite);
    connectorFormat.EndConnect(to, toSite);

    wxExcelLineFormat lineFormat = connector.GetLine();
    lineFormat.GetForeColor().SetRGB(wxColour(64, 64, 64));
    lineFormat.SetWeight(1.5);
    lineFormat.SetEndArrowheadStyle(msoArrowheadTriangle);

    return connector;
}

// Adds a text without a fill and an outline, with the size fitting the text.
wxExcelShape AddLabel(wxExcelShapes& shapes, const wxString& text, double left, double top)
{
    wxExcelShape label = shapes.AddTextbox(msoTextOrientationHorizontal, left, top, 50, 20);
    if ( !label )
        return label;

    label.GetFill().SetVisible(msoFalse);
    label.GetLine().SetVisible(msoFalse);

    wxExcelTextFrame2 textFrame = label.GetTextFrame2();
    textFrame.SetWordWrap(msoFalse);
    textFrame.SetAutoSize(msoAutoSizeShapeToFitText);
    textFrame.GetTextRange().SetText(text);

    return label;
}

void AddFlowchart(wxExcelWorksheet& worksheet)
{
    wxExcelShapes shapes = worksheet.GetShapes();
    if ( !shapes )
    {
        wxLogError(_("Failed to obtain shapes collection from the worksheet."));
        return;
    }

    const wxColour terminatorColor(84, 130, 53);
    const wxColour processColor(46, 117, 182);
    const wxColour decisionColor(197, 90, 17);

    // the steps are only roughly placed in a column, they are aligned below
    wxExcelShape start   = AddStep(shapes, msoShapeFlowchartTerminator, "Start", terminatorColor, 20, 100);
    wxExcelShape receive = AddStep(shapes, msoShapeFlowchartProcess, "Receive the order", processColor, 90);
    wxExcelShape inStock = AddStep(shapes, msoShapeFlowchartDecision, "In stock?", decisionColor, 160, 130, 70);
    wxExcelShape ship    = AddStep(shapes, msoShapeFlowchartProcess, "Ship the goods", processColor, 270);
    wxExcelShape end     = AddStep(shapes, msoShapeFlowchartTerminator, "End", terminatorColor, 350, 100);

    wxVector<wxString> names;

    names.push_back(start.GetName());
    names.push_back(receive.GetName());
    names.push_back(inStock.GetName());
    names.push_back(ship.GetName());
    names.push_back(end.GetName());

    // center the steps in the column and make the spaces between them equal
    wxExcelShapeRange column = shapes.GetRange(names);
    if ( column )
    {
        column.Align(msoAlignCenters);
        column.Distribute(msoDistributeVertically);
    }

    // the step on the side branch is placed next to the decision
    wxExcelShape order = AddStep(shapes, msoShapeFlowchartProcess, "Order from the supplier", processColor, 0);
    order.SetLeft(inStock.GetLeft() + inStock.GetWidth() + 60);
    order.SetTop(inStock.GetTop() + (inStock.GetHeight() - order.GetHeight()) / 2);
    names.push_back(order.GetName());

    names.push_back(Connect(shapes, start, SiteBottom, receive, SiteTop).GetName());
    names.push_back(Connect(shapes, receive, SiteBottom, inStock, SiteTop).GetName());
    names.push_back(Connect(shapes, inStock, SiteBottom, ship, SiteTop).GetName());
    names.push_back(Connect(shapes, inStock, SiteRight, order, SiteLeft).GetName());
    names.push_back(Connect(shapes, order, SiteBottom, ship, SiteRight).GetName());
    names.push_back(Connect(shapes, ship, SiteBottom, end, SiteTop).GetName());

    names.push_back(AddLabel(shapes, "Yes", inStock.GetLeft() + inStock.GetWidth() / 2,
                             inStock.GetTop() + inStock.GetHeight()).GetName());
    names.push_back(AddLabel(shapes, "No", inStock.GetLeft() + inStock.GetWidth(),
                             inStock.GetTop() + inStock.GetHeight() / 2 - 22).GetName());

    // group all the shapes, so they can be moved or copied together
    wxExcelShapeRange flowchart = shapes.GetRange(names);
    if ( flowchart )
        flowchart.Group().SetName("Flowchart");
}

void AddOtherShapes(wxExcelWorksheet& worksheet)
{
    wxExcelShapes shapes = worksheet.GetShapes();
    if ( !shapes )
    {
        wxLogError(_("Failed to obtain shapes collection from the worksheet."));
        return;
    }

    wxExcelColorFormat colorFormat;
    wxExcelLineFormat lineFormat;

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
    shapes.AddPolyline(points);

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

} // unnamed namespace

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
        wxMessageBox(_("This sample requires Microsoft Excel 2007 or newer."), _("Information"));
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

    // the worksheet is active, so the gridlines are hidden for it
    worksheet.SetName("Flowchart");
    app.GetActiveWindow().SetDisplayGridlines(false);
    AddFlowchart(worksheet);

    // the newly added worksheet becomes the active one
    worksheet = workbook.GetWorksheets().AddAfterOrBefore(worksheet, true);
    if ( !worksheet )
    {
        wxLogError(_("Failed to add a worksheet."));
        return;
    }
    worksheet.SetName("Other shapes");
    app.GetActiveWindow().SetDisplayGridlines(false);
    AddOtherShapes(worksheet);
}

void MyFrame::OnQuit(wxCommandEvent& WXUNUSED(event))
{
    Close(true);
}


bool MyApp::OnInit()
{
    if ( !wxApp::OnInit() )
        return false;

    // display wxAutoExcel-related traces in debug output
    wxLog::AddTraceMask(wxTRACE_AutoExcel);

    MyFrame* frame = new MyFrame();
    frame->Show();

    return true;
}

wxIMPLEMENT_APP(MyApp);
