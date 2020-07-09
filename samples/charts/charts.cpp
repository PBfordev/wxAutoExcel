/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////


/**********************************************************

wxAutoExcel Charts sample focuses on:
- Checking if the MS Excel is 2007 (version 12) or newer.
- Adding both an embedded chart and a chart sheet.
- Adding a sparkline (only in Excel 2010 and newer)
- Creating charts of various types.
- Customising appearance of a chart.

**********************************************************/


#include <wx/wx.h>
#include <wx/msw/ole/oleutils.h>
#include <wx/iconbndl.h>

#include <wx/wxAutoExcel.h>


#if  !WXAUTOEXCEL_USE_CHARTS
    #error In order to compile this sample, wxAutoExcel has to be built with WXAUTOEXCEL_USE_CHARTS set to 1 in wxAutoExcel_setup.h
#endif

class MyFrame : public wxFrame
{
public:
    MyFrame();
private:
    void OnShowSample(wxCommandEvent& event);
    void OnQuit(wxCommandEvent& event);
};


class MyApp : public wxApp
{
public:	
    virtual bool OnInit();
};

using namespace wxAutoExcel;


MyFrame::MyFrame()
: wxFrame(NULL, wxID_ANY, _("wxAutoExcel Charts sample"))
{
    SetIcons(wxIconBundle("appIcon", NULL));

    wxMenu *menu = new wxMenu;
    menu->Append(wxID_NEW, _("&Show me!"));
    menu->Append(wxID_EXIT, _("E&xit"));

    wxMenuBar *menuBar = new wxMenuBar();
    menuBar->Append(menu, _("&Sample"));
    SetMenuBar(menuBar);

    Bind(wxEVT_COMMAND_MENU_SELECTED, &MyFrame::OnShowSample, this, wxID_NEW);
    Bind(wxEVT_COMMAND_MENU_SELECTED, &MyFrame::OnQuit, this, wxID_EXIT);
}

class ChartSample
{
public:
    bool Init();

    bool AddChartStacked();
    bool AddChartClusteredWithLine();
    bool AddChart3D();
    bool AddSparkline();
private:
    wxExcelApplication  m_app;
    wxExcelWorkbook     m_workbook;
    wxExcelWorksheet    m_dataWorksheet;
    wxExcelWorksheet    m_embeddedChartsWorksheet;

    bool WriteData(wxExcelRange& range);
};

bool ChartSample::Init()
{
    m_app = wxExcelApplication::CreateInstance();
    if ( !m_app )
    {
        wxLogError(_("Failed to create an instance of MS Excel application."));
        return false;
    }

    if ( !m_app.IsVersionAtLeast_(wxExcelApplication::evExcel2007) )
    {
        wxMessageBox(_("This sample requires Microsoft Excel 2007 or newer."), "Information");
        m_app.Quit();
        return false;
    }

    m_app.SetVisible(true);
    m_app.SetDisplayAlerts(false);

    m_workbook = m_app.GetWorkbooks().Add();
    if ( !m_workbook )
    {
        wxLogError(_("Failed to create a new workbook."));
        return false;
    }
    m_workbook.SetAutomationLCID_(wxExcelObject::lcidEnglishUS);

    m_dataWorksheet = m_workbook.GetWorksheets()[1];
    if ( !m_dataWorksheet )
    {
        wxLogError(_("Failed to obtain worksheet number 1."));
        return false;
    }
    m_dataWorksheet.SetName("Data (millions EUR)");

    wxExcelRange range = m_dataWorksheet.GetRange("A1:F1");
    if ( !WriteData(range) )
        return false;

    m_embeddedChartsWorksheet = m_workbook.GetWorksheets().AddAfterOrBefore(m_dataWorksheet, true);
    if ( !m_embeddedChartsWorksheet )
    {
        wxLogError(_("Failed to add a worksheet."));
        return false;
    }
    m_embeddedChartsWorksheet.SetName("Embedded charts");

    return true;
}


bool ChartSample::WriteData(wxExcelRange& range)
{
    wxVariant variant;

    // write sheet headers
    variant.ClearList();
    variant.Append("Branch");
    variant.Append("Q1");
    variant.Append("Q2");
    variant.Append("Q3");
    variant.Append("Q4");
    variant.Append("Total");
    range = variant;
    range.GetFont().SetBold(true);

    range = range.GetOffset(1);
    variant.ClearList();
    variant.Append("North");
    variant.Append(10);
    variant.Append(12);
    variant.Append(11);
    variant.Append(14);
    variant.Append("=SUM(RC[-4]:RC[-1])");
    range = variant;

    range = range.GetOffset(1);
    variant.ClearList();
    variant.Append("East");
    variant.Append(15);
    variant.Append(17);
    variant.Append(18);
    variant.Append(22);
    variant.Append("=SUM(RC[-4]:RC[-1])");
    range = variant;

    range = range.GetOffset(1);
    variant.ClearList();
    variant.Append("South");
    variant.Append(12);
    variant.Append(14);
    variant.Append(18);
    variant.Append(20);
    variant.Append("=SUM(RC[-4]:RC[-1])");
    range = variant;

    range = range.GetOffset(1);
    variant.ClearList();
    variant.Append("West");
    variant.Append(11);
    variant.Append(11);
    variant.Append(12);
    variant.Append(15);
    variant.Append("=SUM(RC[-4]:RC[-1])");
    range = variant;

    range = range.GetOffset(1);
    variant.ClearList();
    variant.Append("Total");
    variant.Append("=SUM(R[-4]C:R[-1]C)");
    variant.Append("=SUM(R[-4]C:R[-1]C)");
    variant.Append("=SUM(R[-4]C:R[-1]C)");
    variant.Append("=SUM(R[-4]C:R[-1]C)");
    variant.Append("=SUM(R[-4]C:R[-1]C)");
    range = variant;

    range = range.GetOffset(1);
    variant.ClearList();
    variant.Append("Average");
    variant.Append("=R[-1]C/4");
    variant.Append("=R[-1]C/4");
    variant.Append("=R[-1]C/4");
    variant.Append("=R[-1]C/4");
    variant.Append("=R[-1]C/4");
    range = variant;

    return range;
}

bool ChartSample::AddChartStacked()
{
    wxExcelChart chart;

    chart = m_embeddedChartsWorksheet.GetShapes().AddChart(xlColumnStacked, 1, 1, 250, 250).GetChart();
    chart.SetHasTitle(true);
    chart.GetChartTitle().SetText("Stacked chart");
    m_embeddedChartsWorksheet.ChartObjects()[1].SetRoundedCorners(true);

    if ( !chart )
        return false;

    wxExcelRange sourceRange = m_dataWorksheet.GetRange("A1:E5");
    if ( !sourceRange )
        return false;

    chart.SetSourceData(sourceRange);

    wxExcelAxis axis = chart.Axes(xlValue);
    wxExcelAxisTitle axisTitle;

    axis.SetHasTitle(true);
    axisTitle = axis.GetAxisTitle();
    axisTitle.SetCaption("Sales");


    axis = chart.Axes(xlCategory);

    // customize category axis
    wxArrayString categories;
    categories.push_back("1");
    categories.push_back("2");
    categories.push_back("3");
    categories.push_back("4");
    axis.SetCategoryNames(categories);

    axis.SetHasTitle(true);
    axisTitle = axis.GetAxisTitle();
    axisTitle.SetCaption("Quarter");


    return chart;
}

bool ChartSample::AddChartClusteredWithLine()
{
    wxExcelChart chart;

    chart = m_embeddedChartsWorksheet.GetShapes().AddChart(xlColumnClustered, 260, 1, 250, 250).GetChart();
    chart.SetHasTitle(true);
    chart.GetChartTitle().SetText("Clustered chart with a line");

    if ( !chart )
        return false;

    wxExcelRange sourceRange = m_dataWorksheet.GetRange("A1:E5");
    if ( !sourceRange )
        return false;

    chart.SetSourceData(sourceRange);

    wxExcelSeries lineSeries = chart.SeriesCollection().NewSeries();
    lineSeries.SetName("Average");
    lineSeries.SetValues(m_dataWorksheet.GetRange("B7:E7"));
    lineSeries.SetChartType(xlLineMarkers);


    wxExcelAxis axis = chart.Axes(xlValue);
    wxExcelAxisTitle axisTitle;

    axis.SetHasTitle(true);
    axisTitle = axis.GetAxisTitle();
    axisTitle.SetCaption("Sales");

    return chart;
}

bool ChartSample::AddChart3D()
{
    wxExcelChart chart;

    chart = m_workbook.GetCharts().Add();
    chart.MoveAfterOrBefore(m_embeddedChartsWorksheet, true);
    chart.SetChartType(xl3DColumnClustered);
    chart.SetName("Customised 3D clustered chart");

    if ( !chart )
        return false;

    wxExcelRange sourceRange = m_dataWorksheet.GetRange("A1:E5");
    if ( !sourceRange )
        return false;

    chart.SetSourceData(sourceRange);

    long seriesCount = chart.SeriesCollection().GetCount();
    for ( long l = 1; l <= seriesCount; l++ )
    {
        chart.SeriesCollection()[l].ApplyDataLabels();
    }

    chart.SetHasLegend(true);
    chart.GetLegend().SetPosition(xlLegendPositionTop);

    wxExcelAxis axis = chart.Axes(xlValue);

    axis.SetHasTitle(true);
    axis.GetAxisTitle().SetCaption("Sales");

    chart.GetBackWall().GetFormat().GetFill().PresetGradient(msoGradientDiagonalDown, 1, msoGradientGold);
    chart.GetSideWall().GetFormat().GetFill().PresetGradient(msoGradientDiagonalDown, 1, msoGradientGold);

    wxExcelFillFormat fill = chart.GetFloor().GetFormat().GetFill();
    fill.Solid();
    fill.GetBackColor().SetRGB(*wxBLACK);

    return chart;
}


bool ChartSample::AddSparkline()
{
    if ( !m_app.IsVersionAtLeast_(wxExcelApplication::evExcel2010) )
    {
        wxMessageBox(_("Sparklines not added, this feature requires Microsoft Excel 2010 or newer."), "Information");
        return false;
    }

    m_dataWorksheet.Select();

    // add sparklines for branches by quarters
    // minimum value for each chart is not set here, so they can look quite deceptive
    m_dataWorksheet.GetRange("G2").GetSparklineGroups().Add(xlSparkLine, "B2:E2");
    m_dataWorksheet.GetRange("G3").GetSparklineGroups().Add(xlSparkLine, "B3:E3");
    m_dataWorksheet.GetRange("G4").GetSparklineGroups().Add(xlSparkLine, "B4:E4");
    m_dataWorksheet.GetRange("G5").GetSparklineGroups().Add(xlSparkLine, "B5:E5");

    // add bars for total sales by branches
    wxExcelSparklineGroups groups;
    wxExcelSparkVerticalAxis axis;
    wxExcelSparkPoints points;

    groups = m_dataWorksheet.GetRange("F8").GetSparklineGroups();
    groups.Add(xlSparkColumn, "F2:F5");

    // set minimal axis value to 0
    axis = groups[1].GetAxes().GetVertical();
    axis.SetMinScaleType(xlSparkScaleCustom);
    axis.SetCustomMinScaleValue(0.);

    // set the color of highest bar to green and lowest to red
    points = groups[1].GetPoints();
    points.GetHighpoint().SetVisible(true);
    points.GetHighpoint().GetColor().SetColor(*wxGREEN);
    points.GetLowpoint().SetVisible(true);
    points.GetLowpoint().GetColor().SetColor(*wxRED);

    return true;
}


void MyFrame::OnShowSample(wxCommandEvent& WXUNUSED(event))
{
    ChartSample sample;

    if ( !sample.Init() )
        return;

    sample.AddSparkline();
    sample.AddChartStacked();
    sample.AddChartClusteredWithLine();
    sample.AddChart3D();
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
