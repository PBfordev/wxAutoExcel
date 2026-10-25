/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// License:     MIT license
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

    Bind(wxEVT_MENU, &MyFrame::OnShowSample, this, wxID_NEW);
    Bind(wxEVT_MENU, &MyFrame::OnQuit, this, wxID_EXIT);
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
    static void SetWallGradient(wxExcelFillFormat fill);
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
        wxMessageBox(_("This sample requires Microsoft Excel 2007 or newer."), _("Information"));
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
    static const char* const headers[] = { "Branch", "Q1", "Q2", "Q3", "Q4", "Total" };
    static const struct
    {
        const char* name;
        long        sales[4];
    } branches[] =
    {
        { "North", { 10, 12, 11, 14 } },
        { "East",  { 15, 17, 18, 22 } },
        { "South", { 12, 14, 18, 20 } },
        { "West",  { 11, 11, 12, 15 } },
    };
    static const struct
    {
        const char* name;
        const char* formula;
    } summaries[] =
    {
        { "Total",   "=SUM(R[-4]C:R[-1]C)" },
        { "Average", "=R[-1]C/4" },
    };

    // each row is written at once, as a list with a value for each cell
    wxVariant row;

    row.ClearList();
    for ( size_t i = 0; i < WXSIZEOF(headers); ++i )
        row.Append(headers[i]);
    range = row;
    range.GetFont().SetBold(true);

    for ( size_t i = 0; i < WXSIZEOF(branches); ++i )
    {
        range = range.GetOffset(1);
        row.ClearList();
        row.Append(branches[i].name);
        for ( size_t j = 0; j < WXSIZEOF(branches[i].sales); ++j )
            row.Append(branches[i].sales[j]);
        row.Append("=SUM(RC[-4]:RC[-1])");
        range = row;
    }

    for ( size_t i = 0; i < WXSIZEOF(summaries); ++i )
    {
        range = range.GetOffset(1);
        row.ClearList();
        row.Append(summaries[i].name);
        for ( size_t j = 1; j < WXSIZEOF(headers); ++j )
            row.Append(summaries[i].formula);
        range = row;
    }

    return range;
}

bool ChartSample::AddChartStacked()
{
    wxExcelChart chart;

    chart = m_embeddedChartsWorksheet.GetShapes().AddChart(xlColumnStacked, 1, 1, 250, 250).GetChart();
    if ( !chart )
        return false;

    chart.SetHasTitle(true);
    chart.GetChartTitle().SetText("Stacked chart");
    m_embeddedChartsWorksheet.ChartObjects()[1].SetRoundedCorners(true);

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

    // replace the category names taken from the source range (Q1-Q4)
    wxArrayString categories;
    categories.push_back("Jan-Mar");
    categories.push_back("Apr-Jun");
    categories.push_back("Jul-Sep");
    categories.push_back("Oct-Dec");
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
    if ( !chart )
        return false;

    chart.SetHasTitle(true);
    chart.GetChartTitle().SetText("Clustered chart with a line");

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

void ChartSample::SetWallGradient(wxExcelFillFormat fill)
{
    // start with a gradient of two colors
    fill.TwoColorGradient(msoGradientHorizontal, 1);

    // change the colors of its two stops and add a third stop between them
    wxExcelGradientStops stops = fill.GetGradientStops();
    if ( !stops )
        return;

    stops[1].GetColor().SetRGB(*wxWHITE);
    stops[2].GetColor().SetRGB(wxColour(31, 78, 121));
    stops.Insert(wxColour(189, 215, 238), 0.6);
}

bool ChartSample::AddChart3D()
{
    wxExcelChart chart;

    chart = m_workbook.GetCharts().Add();
    if ( !chart )
        return false;

    chart.MoveAfterOrBefore(m_embeddedChartsWorksheet, true);
    chart.SetChartType(xl3DColumnClustered);
    chart.SetName("Customised 3D clustered chart");

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

    SetWallGradient(chart.GetBackWall().GetFormat().GetFill());
    SetWallGradient(chart.GetSideWall().GetFormat().GetFill());

    wxExcelFillFormat fill = chart.GetFloor().GetFormat().GetFill();
    fill.Solid();
    fill.GetForeColor().SetRGB(wxColour(31, 78, 121));

    return chart;
}


bool ChartSample::AddSparkline()
{
    if ( !m_app.IsVersionAtLeast_(wxExcelApplication::evExcel2010) )
    {
        wxMessageBox(_("Sparklines not added, this feature requires Microsoft Excel 2010 or newer."), _("Information"));
        return false;
    }

    m_dataWorksheet.Select();

    wxExcelSparklineGroups groups;
    wxExcelSparkVerticalAxis axis;
    wxExcelSparkPoints points;

    // add sparklines for branches by quarters, one for each row of the source data;
    // they are added at once, so they are in the same group and can share the scale
    groups = m_dataWorksheet.GetRange("G2:G5").GetSparklineGroups();
    groups.Add(xlSparkLine, "B2:E5");

    // without the common scale, each sparkline would be scaled to its own values
    // and the branches with very different sales could look the same
    axis = groups[1].GetAxes().GetVertical();
    axis.SetMinScaleType(xlSparkScaleGroup);
    axis.SetMaxScaleType(xlSparkScaleGroup);

    // add bars for total sales by branches
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
    if ( !wxApp::OnInit() )
        return false;

    // display wxAutoExcel-related traces in debug output
    wxLog::AddTraceMask(wxTRACE_AutoExcel);

    MyFrame* frame = new MyFrame();
    frame->Show();

    return true;
}

wxIMPLEMENT_APP(MyApp);
