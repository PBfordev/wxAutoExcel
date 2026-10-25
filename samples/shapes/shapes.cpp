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
- Applying effects to shapes: shadow, glow, soft edges,
  reflection, bevel, 3D rotation, gradient, flipping,
  and rotating.
- Adding form controls linked to cells.
- Placing shapes so that they move and size with cells.
- Drawing freeform shapes with curves and lines, and moving
  their nodes.
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

// Adds a shape of the given type with no outline and with the text centered in it.
wxExcelShape AddShapeWithText(wxExcelShapes& shapes, MsoAutoShapeType type, const wxString& text,
                              const wxColour& color, double left, double top,
                              double width, double height)
{
    wxExcelShape shape = shapes.AddShape(type, left, top, width, height);
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
    // the anchor centers the text as a whole, the alignment its lines
    textRange.GetParagraphFormat().SetAlignment(msoAlignCenter);

    wxExcelFont2 font = textRange.GetFont();
    font.SetBold(msoTrue);
    font.GetFill().GetForeColor().SetRGB(*wxWHITE);

    return shape;
}

// Adds a flowchart step, the steps are placed in a column.
wxExcelShape AddStep(wxExcelShapes& shapes, MsoAutoShapeType type, const wxString& text,
                     const wxColour& color, double top, double width = 130, double height = 45)
{
    return AddShapeWithText(shapes, type, text, color, 30, top, width, height);
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

// Adds a shape for showing an effect, with the name of the effect in it;
// the shapes are placed in rows of four.
wxExcelShape AddEffectShape(wxExcelShapes& shapes, MsoAutoShapeType type,
                            const wxString& name, long index)
{
    const double left = 30 + (index % 4) * 130;
    const double top = 30 + (index / 4) * 150;

    return AddShapeWithText(shapes, type, name, wxColour(46, 117, 182), left, top, 90, 90);
}

void AddEffects(wxExcelWorksheet& worksheet)
{
    wxExcelShapes shapes = worksheet.GetShapes();
    if ( !shapes )
    {
        wxLogError(_("Failed to obtain shapes collection from the worksheet."));
        return;
    }

    long index = 0;
    wxExcelShape shape;

    shape = AddEffectShape(shapes, msoShapeRoundedRectangle, "Shadow", index++);
    wxExcelShadowFormat shadow = shape.GetShadow();
    shadow.SetVisible(msoTrue);
    shadow.SetStyle(msoShadowStyleOuterShadow);
    shadow.GetForeColor().SetRGB(*wxBLACK);
    shadow.SetOffsetX(6);
    shadow.SetOffsetY(6);
    shadow.SetBlur(8);
    shadow.SetTransparency(0.5);

    shape = AddEffectShape(shapes, msoShapeRoundedRectangle, "Glow", index++);
    wxExcelGlowFormat glow = shape.GetGlow();
    glow.GetColor().SetRGB(wxColour(255, 192, 0));
    glow.SetRadius(12);

    shape = AddEffectShape(shapes, msoShapeRoundedRectangle, "Soft edges", index++);
    shape.GetSoftEdge().SetType(msoSoftEdgeType4);

    shape = AddEffectShape(shapes, msoShapeRoundedRectangle, "Reflection", index++);
    shape.GetReflection().SetType(msoReflectionType2);

    shape = AddEffectShape(shapes, msoShapeRoundedRectangle, "Bevel", index++);
    wxExcelThreeDFormat threeD = shape.GetThreeD();
    threeD.SetBevelTopType(msoBevelCircle);
    threeD.SetBevelTopInset(15);
    threeD.SetBevelTopDepth(8);

    shape = AddEffectShape(shapes, msoShapeRoundedRectangle, "3D rotation", index++);
    threeD = shape.GetThreeD();
    threeD.SetPresetCamera(msoCameraIsometricOffAxis1Right);
    threeD.SetDepth(30);

    // a gradient with three colors: start with a gradient of two colors,
    // change the colors of its two stops, and add a third stop between them
    shape = AddEffectShape(shapes, msoShapeHeart, "Gradient", index++);
    wxExcelFillFormat fill = shape.GetFill();
    fill.TwoColorGradient(msoGradientHorizontal, 1);
    wxExcelGradientStops stops = fill.GetGradientStops();
    if ( stops )
    {
        stops[1].GetColor().SetRGB(wxColour(255, 128, 128));
        stops[2].GetColor().SetRGB(wxColour(128, 0, 0));
        stops.Insert(*wxRED, 0.5);
    }

    // the arrow points to the right, flipping it makes it point to the left;
    // the text is rotated with the shape but it is not flipped
    shape = AddEffectShape(shapes, msoShapeRightArrow, "Flipped and rotated", index++);
    shape.Flip(msoFlipHorizontal);
    shape.SetRotation(30);
}

// Adds a shape with the text covering the given range.
wxExcelShape AddShapeOverRange(wxExcelShapes& shapes, wxExcelRange range, const wxString& text)
{
    return AddShapeWithText(shapes, msoShapeRectangle, text, wxColour(46, 117, 182),
                            range.GetLeft(), range.GetTop(), range.GetWidth(), range.GetHeight());
}

void AddFormControls(wxExcelWorksheet& worksheet)
{
    wxExcelShapes shapes = worksheet.GetShapes();
    if ( !shapes )
    {
        wxLogError(_("Failed to obtain shapes collection from the worksheet."));
        return;
    }

    // an order with the quantity set with a spinner and
    // the discount set with a check box
    worksheet.GetRange("1:4").SetRowHeight(22);
    worksheet.GetRange("A1").SetValue("Quantity");
    worksheet.GetRange("A2").SetValue("Unit price");
    worksheet.GetRange("B2").SetValue(12.5);
    worksheet.GetRange("A3").SetValue("10% discount");
    worksheet.GetRange("B3").SetValue(false);
    worksheet.GetRange("A4").SetValue("Total");
    worksheet.GetRange("B4").SetFormula("=B1*B2*IF(B3,0.9,1)");
    worksheet.GetRange("A4:B4").GetFont().SetBold(true);
    worksheet.GetRange("B2:B4").SetNumberformat("0.00");

    // the controls are placed in the cells next to the cells they are linked to,
    // the values of the linked cells change when the controls are used
    wxExcelRange cell = worksheet.GetRange("C1");
    wxExcelShape spinner = shapes.AddFormControl(xlSpinner, cell.GetLeft() + 2, cell.GetTop() + 1,
                                                 20, cell.GetHeight() - 2);
    if ( spinner )
    {
        wxExcelControlFormat control = spinner.GetControlFormat();
        control.SetMin(1);
        control.SetMax(100);
        control.SetSmallChange(1);
        control.SetLinkedCell("$B$1");
        control.SetValue(5);
    }

    cell = worksheet.GetRange("C3");
    wxExcelShape checkBox = shapes.AddFormControl(xlCheckBox, cell.GetLeft() + 2, cell.GetTop() + 1,
                                                  60, cell.GetHeight() - 2);
    if ( checkBox )
    {
        checkBox.GetTextFrame().Characters().SetText("Apply");

        wxExcelControlFormat control = checkBox.GetControlFormat();
        control.SetLinkedCell("$B$3");
        control.SetValue(xlOff);
    }

    // shapes can move and size with the cells under them or stay where they are;
    // after widening column B, only the first shape below is wider
    wxExcelShape shape;

    shape = AddShapeOverRange(shapes, worksheet.GetRange("B7:C9"), "Moves and sizes with cells");
    shape.SetPlacement(xlMoveAndSize);

    shape = AddShapeOverRange(shapes, worksheet.GetRange("B11:C13"), "Free floating");
    shape.SetPlacement(xlFreeFloating);

    worksheet.GetRange("B1").GetEntireColumn().SetColumnWidth(25);
}

void AddFreeforms(wxExcelWorksheet& worksheet)
{
    wxExcelShapes shapes = worksheet.GetShapes();
    if ( !shapes )
    {
        wxLogError(_("Failed to obtain shapes collection from the worksheet."));
        return;
    }

    wxExcelFreeformBuilder builder;

    // a drop drawn with Bezier curves, each given by two control points and its end
    builder = shapes.BuildFreeform(msoEditingCorner, 80, 30);
    if ( !builder )
        return;

    double x2 = 125, y2 = 85, x3 = 125, y3 = 115;
    builder.AddNodes(msoSegmentCurve, msoEditingCorner, 95, 60, &x2, &y2, &x3, &y3);
    x2 = 35; y2 = 145; x3 = 35; y3 = 115;
    builder.AddNodes(msoSegmentCurve, msoEditingCorner, 125, 145, &x2, &y2, &x3, &y3);
    x2 = 65; y2 = 60; x3 = 80; y3 = 30;
    builder.AddNodes(msoSegmentCurve, msoEditingCorner, 35, 85, &x2, &y2, &x3, &y3);

    wxExcelShape drop = builder.ConvertToShape();
    if ( !drop )
        return;
    drop.GetFill().GetForeColor().SetRGB(wxColour(46, 117, 182));
    drop.GetLine().SetVisible(msoFalse);
    AddLabel(shapes, "Curves", 55, 180);

    // a lightning bolt drawn with straight lines
    static const double boltPoints[][2] =
    {
        { 260, 30 }, { 235, 85 }, { 270, 85 }, { 195, 170 },
        { 215, 105 }, { 180, 105 }, { 200, 30 }
    };

    builder = shapes.BuildFreeform(msoEditingAuto, 200, 30);
    if ( !builder )
        return;

    for ( size_t i = 0; i < WXSIZEOF(boltPoints); ++i )
        builder.AddNodes(msoSegmentLine, msoEditingAuto, boltPoints[i][0], boltPoints[i][1]);

    wxExcelShape bolt = builder.ConvertToShape();
    if ( !bolt )
        return;
    bolt.GetFill().GetForeColor().SetRGB(wxColour(255, 192, 0));
    bolt.GetLine().GetForeColor().SetRGB(wxColour(191, 144, 0));
    AddLabel(shapes, "Lines", 205, 180);

    // a copy of the drop with a longer tip: the tip of the copy is moved up; the shape
    // is closed, so the tip is both its first and last node, and both must be moved
    wxExcelShape copy = drop.Duplicate();
    if ( !copy )
        return;
    copy.SetLeft(330);
    copy.SetTop(drop.GetTop());

    const wxVector<wxPoint2DDouble> vertices = copy.GetVertices();
    wxExcelShapeNodes nodes = copy.GetNodes();
    if ( !vertices.empty() && nodes )
    {
        const wxPoint2DDouble tip(vertices[0].m_x, vertices[0].m_y - 30);

        nodes.SetPosition(1, tip.m_x, tip.m_y);
        nodes.SetPosition(nodes.GetCount(), tip.m_x, tip.m_y);
    }
    AddLabel(shapes, "A moved node", 335, 180);
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
    // use English formulas and number formats regardless of the Excel language
    workbook.SetAutomationLCID_(wxExcelObject::lcidEnglishUS);

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
    worksheet.SetName("Effects");
    app.GetActiveWindow().SetDisplayGridlines(false);
    AddEffects(worksheet);

    worksheet = workbook.GetWorksheets().AddAfterOrBefore(worksheet, true);
    if ( !worksheet )
    {
        wxLogError(_("Failed to add a worksheet."));
        return;
    }
    worksheet.SetName("Form controls");
    app.GetActiveWindow().SetDisplayGridlines(false);
    AddFormControls(worksheet);

    worksheet = workbook.GetWorksheets().AddAfterOrBefore(worksheet, true);
    if ( !worksheet )
    {
        wxLogError(_("Failed to add a worksheet."));
        return;
    }
    worksheet.SetName("Freeforms");
    app.GetActiveWindow().SetDisplayGridlines(false);
    AddFreeforms(worksheet);
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
