/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Modified by:
// RCS-ID:      $Id: $
// Copyright:   (c) 2012 PB <pb4dev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////

#include <wx/stdpaths.h>

#include "getdata.h"

#include "xlspy.h"


class MyTreeItemData : public wxTreeItemData
{
public:
    wxStringPairVector m_xlData;
};


MyFrame::MyFrame()
: wxFrame(NULL, wxID_ANY, _("wxAutoExcel XLSpy sample"), 
          wxDefaultPosition, wxSize(800, 500))
{
    wxMenu *menu = new wxMenu;
    menu->Append(wxID_OPEN, _("&Open file...\tCtrl+O"));
    menu->Append(wxID_EXIT, _("E&xit"));

    wxMenuBar *menuBar = new wxMenuBar();
    menuBar->Append(menu, _("&Sample"));
    SetMenuBar(menuBar);
            
    m_treeCtrl = new wxTreeCtrl(this);    

    m_listCtrl = new wxListCtrl(this, wxID_ANY, wxDefaultPosition, wxDefaultSize, 
        wxLC_REPORT | wxLC_SINGLE_SEL);
    m_listCtrl->AppendColumn(_("Name"));
    m_listCtrl->AppendColumn(_("Value"));

    wxBoxSizer* sizer = new wxBoxSizer(wxHORIZONTAL);
    sizer->Add(m_treeCtrl, 2, wxALL|wxEXPAND, 0);
    sizer->Add(m_listCtrl, 3, wxALL|wxEXPAND, 0);
    SetSizer(sizer);
    Layout();

    Bind(wxEVT_COMMAND_MENU_SELECTED, &MyFrame::OnOpenFile, this, wxID_OPEN);
    Bind(wxEVT_COMMAND_MENU_SELECTED, &MyFrame::OnQuit, this, wxID_EXIT);
    Bind(wxEVT_CLOSE_WINDOW, &MyFrame::OnClose, this);

    m_treeCtrl->Bind(wxEVT_TREE_SEL_CHANGED, &MyFrame::OnTreeCtrlSelChanged, this);   

    wxBusyCursor busy;
    if ( CreateExcelInstance() )
    {
        m_treeCtrl->AddRoot(_("Microsoft Excel"));
        m_treeCtrl->SelectItem(AppendApplicationData(m_treeCtrl->GetRootItem()));
        m_treeCtrl->ExpandAll();  
    }
}



void MyFrame::OnOpenFile(wxCommandEvent& WXUNUSED(event))
{        
    if ( !m_app )
        return;
    
    if ( !m_app.Is2007OrNewer() )
    {
        // only so the sample is not cluttered with version checks for 2007+ features
        wxMessageBox(_("This sample requires Microsoft Excel 2007 or newer."), "Information");        
        return;
    }

    static wxString defaultDir;

    if ( defaultDir.empty() )
    {
        defaultDir = wxStandardPaths::Get().GetDataDir();
    }

    wxFileDialog fd(this, _("Select Excel File"), defaultDir, "",
                    _("MS Excel files (*.xls?)|*.xls?|All files (*.*)|*.*"), wxFD_OPEN | wxFD_FILE_MUST_EXIST);
        
    if ( fd.ShowModal() == wxID_CANCEL )
        return;

    defaultDir = fd.GetDirectory();
    
    wxBusyCursor busy;
    
    m_workbook = m_app.GetWorkbooks().Open(fd.GetPath(), WXAEEP(0L), true);
     
    if ( !m_workbook ) 
    {
        wxLogError(_("Failed to open file %s."), fd.GetPath());
        return;
    }                

    m_treeCtrl->DeleteAllItems(); 
    m_listCtrl->DeleteAllItems();
    
    m_treeCtrl->AddRoot(_("Microsoft Excel"));
    wxTreeItemId appId = AppendApplicationData(m_treeCtrl->GetRootItem());
    
    AddWorkbookData(m_treeCtrl->GetRootItem());   
    m_workbook.Close();
            
    m_treeCtrl->SelectItem(appId);
    m_treeCtrl->ExpandAll();    
}

void MyFrame::OnQuit(wxCommandEvent&)
{    
    if ( m_app )
        m_app.Quit();

    Close(true);
}

void MyFrame::OnClose(wxCloseEvent&)
{
    Destroy();
}


void MyFrame::OnTreeCtrlSelChanged(wxTreeEvent& evt)
{         
    m_listCtrl->DeleteAllItems();
    
    MyTreeItemData* tiData = dynamic_cast<MyTreeItemData*>( m_treeCtrl->GetItemData(evt.GetItem()) );
    if ( !tiData )
        return;
    
    wxListItem li;
    const wxStringPairVector& xlData = tiData->m_xlData;
        
    for ( size_t i = 0; i < xlData.size(); i++ )
    {
        li.SetId((long)i);
        li.SetStateMask(wxLIST_MASK_TEXT);
        
        li.SetColumn(0);
        li.SetText(xlData[i].first);
        m_listCtrl->InsertItem(li);
        
        li.SetColumn(1);
        li.SetText(xlData[i].second);
        m_listCtrl->SetItem(li);
    }
    m_listCtrl->SetColumnWidth(0, wxLIST_AUTOSIZE);
    m_listCtrl->SetColumnWidth(1, wxLIST_AUTOSIZE);
}


bool MyFrame::CreateExcelInstance()
{
    m_app = wxExcelApplication::CreateInstance();
    
    if ( !m_app ) 
    {
        wxLogError(_("Failed to create an instance of MS Excel application."));
        return false;
    }
    
    m_app.SetAutomationLCID_(1033);
    m_app.SetDisplayAlerts(false);
    m_app.SetVisible(false);        

    return true;
}

void MyFrame::AddWorkbookData(const wxTreeItemId& id)
{                                    
    MyTreeItemData* data = new MyTreeItemData();
    ExcelSpy::GetWorkbookData(m_app, m_workbook, data->m_xlData);            
    wxTreeItemId wkbId = m_treeCtrl->AppendItem(id, m_workbook.GetName(), -1, -1, data);    
        
    // Built-in document properties
    data = new MyTreeItemData();
    ExcelSpy::GetDocumentPropertiesData(m_workbook.GetBuiltinDocumentProperties(), data->m_xlData);
    m_treeCtrl->AppendItem(wkbId, _("Built-in document properties"), -1, -1, data);

    // Styles
    data = new MyTreeItemData();
    ExcelSpy::GetStylesData(m_workbook, data->m_xlData);
    m_treeCtrl->AppendItem(wkbId, _("Styles"), -1, -1, data);

    AddSheetsData(wkbId);
    AddWorksheetsData(wkbId);
    AddChartsData(wkbId);
}

void MyFrame::AddSheetsData(const wxTreeItemId& id)
{
    wxExcelSheets sheets = m_workbook.GetSheets();

    if ( !sheets ) 
    {
        wxLogError(_("Failed to obtain Sheets object."));
        return;
    }        
    
    MyTreeItemData* data = new MyTreeItemData();
    ExcelSpy::GetSheetsData(sheets, data->m_xlData);
    
    wxTreeItemId sheetsId;
    sheetsId = m_treeCtrl->AppendItem(id, _("Sheets"), -1, -1, data);

    wxExcelSheet sheet;
    long count = sheets.GetCount();

    for ( long l = 1; l <= count; l++ )
    {
        sheet = sheets[l];
        if ( !sheet )
        {
            wxLogError(_("Failed to obtain Sheet No. %ld."), count);
            continue;
        }
        data = new MyTreeItemData();
        ExcelSpy::GetSheetData(sheet, data->m_xlData);
        m_treeCtrl->AppendItem(sheetsId, sheet.GetName(), -1, -1, data);
    }        
}    

void MyFrame::AddWorksheetsData(const wxTreeItemId& id)
{
    wxExcelWorksheets sheets = m_workbook.GetWorksheets();

    if ( !sheets ) 
    {
        wxLogError(_("Failed to obtain Worksheets object."));
        return;
    }        
    
    MyTreeItemData* data = new MyTreeItemData();
    ExcelSpy::GetWorksheetsData(sheets, data->m_xlData);
    
    wxTreeItemId sheetsId;
    sheetsId = m_treeCtrl->AppendItem(id, _("Workheets"), -1, -1, data);

    wxExcelWorksheet sheet;
    long count = sheets.GetCount();

    for ( long l = 1; l <= count; l++ )
    {
        sheet = sheets[l];
        if ( !sheet )
        {
            wxLogError(_("Failed to obtain Worksheet No. %ld."), count);
            continue;
        }
        data = new MyTreeItemData();
        ExcelSpy::GetWorksheetData(sheet, data->m_xlData);
        wxTreeItemId sheetId = m_treeCtrl->AppendItem(sheetsId, sheet.GetName(), -1, -1, data);

        // PageSetup   
        data = new MyTreeItemData();
        wxExcelPageSetup pageSetup = sheet.GetPageSetup();
        ExcelSpy::GetPageSetupData(pageSetup, data->m_xlData);
        m_treeCtrl->AppendItem(sheetId, _("PageSetup"), -1, -1, data);
        
        // UsedRange
        wxExcelRange range = sheet.GetUsedRange();
        if ( range )
        {
            data = new MyTreeItemData();
            ExcelSpy::GetRangeData(range, data->m_xlData);
            m_treeCtrl->AppendItem(sheetId, _("UsedRange"), -1, -1, data);
        }

        // Comments        
        data = new MyTreeItemData();
        ExcelSpy::GetCommentsData(sheet, data->m_xlData);
        m_treeCtrl->AppendItem(sheetId, _("Comments"), -1, -1, data);
                

        AddOLEObjectsData(sheet, sheetId);
        AddShapesData(sheet, sheetId);
        AddChartObjectsData(sheet, sheetId);        
        AddHyperlinksData(sheet, sheetId);
    }        
}    

void MyFrame::AddChartsData(const wxTreeItemId& id)
{
#if WXAUTOEXCEL_USE_CHARTS
    wxExcelCharts charts = m_workbook.GetCharts();

    if ( !charts ) 
    {
        wxLogError(_("Failed to obtain Charts object."));
        return;
    }        
    
    MyTreeItemData* data = new MyTreeItemData();
    ExcelSpy::GetChartsData(charts, data->m_xlData);
    
    wxTreeItemId chartsId;
    chartsId = m_treeCtrl->AppendItem(id, _("Charts"), -1, -1, data);

    wxExcelChart chart;
    long count = charts.GetCount();

    for ( long l = 1; l <= count; l++ )
    {
        chart = charts[l];
        if ( !chart )
        {
            wxLogError(_("Failed to obtain chart No. %ld."), count);
            continue;
        }
        data = new MyTreeItemData();
        ExcelSpy::GetChartData(chart, data->m_xlData);
        wxTreeItemId chartId = m_treeCtrl->AppendItem(chartsId, chart.GetName(), -1, -1, data);

        // PageSetup   
        data = new MyTreeItemData();
        wxExcelPageSetup pageSetup = chart.GetPageSetup();
        ExcelSpy::GetPageSetupData(pageSetup, data->m_xlData);
        m_treeCtrl->AppendItem(chartId, _("PageSetup"), -1, -1, data);                        
    }            
#endif // #if WXAUTOEXCEL_USE_CHARTS
}

void MyFrame::AddOLEObjectsData(wxExcelWorksheet& sheet, const wxTreeItemId& sheetId)
{
    wxExcelOLEObjects objects = sheet.OLEObjects();

    if ( !objects )
        return;

    wxTreeItemId objectsId;
    MyTreeItemData* data = new MyTreeItemData();    
    long count = objects.GetCount();

    ExcelSpy::GetOLEObjectsData(objects, data->m_xlData);
    objectsId = m_treeCtrl->AppendItem(sheetId, _("OLE objects"), -1, -1, data);

    for (long l = 1; l <= count; l++ )
    {
        wxExcelOLEObject obj = objects[l];

        data = new MyTreeItemData();
        ExcelSpy::GetOLEObjectData(obj, data->m_xlData);
        m_treeCtrl->AppendItem(objectsId, obj.GetName(), -1, -1, data);
    }
}

void MyFrame::AddShapesData(wxExcelWorksheet& sheet, const wxTreeItemId& sheetId)
{
#if WXAUTOEXCEL_USE_SHAPES    
    wxExcelShapes shapes = sheet.GetShapes();

    if ( !shapes )
        return;

    wxTreeItemId shapesId;
    MyTreeItemData* data = new MyTreeItemData();        

    ExcelSpy::GetShapesData(shapes, data->m_xlData);
    shapesId = m_treeCtrl->AppendItem(sheetId, _("Shapes"), -1, -1, data);
    long count = shapes.GetCount();

    for (long l = 1; l <= count; l++ )
    {
        wxExcelShape shape = shapes[l];

        data = new MyTreeItemData();
        ExcelSpy::GetShapeData(shape, data->m_xlData);
        m_treeCtrl->AppendItem(shapesId, shape.GetName(), -1, -1, data);
    }
#endif // #if WXAUTOEXCEL_USE_SHAPES
}

void MyFrame::AddChartObjectsData(wxExcelWorksheet& sheet, const wxTreeItemId& sheetId)
{
#if WXAUTOEXCEL_USE_CHARTS
    wxExcelChartObjects chartObjects = sheet.ChartObjects();

    if ( !chartObjects )
        return;

    wxTreeItemId chartsId;

    MyTreeItemData* data = new MyTreeItemData();        

    ExcelSpy::GetChartObjectsData(chartObjects, data->m_xlData);
    chartsId = m_treeCtrl->AppendItem(sheetId, _("ChartObjects"), -1, -1, data);
        
    long count = chartObjects.GetCount();

    for (long l = 1; l <= count; l++ )
    {
        wxExcelChartObject chartObject = chartObjects[l];

        data = new MyTreeItemData();
        ExcelSpy::GetChartObjectData(chartObject, data->m_xlData);
        m_treeCtrl->AppendItem(chartsId, chartObject.GetName(), -1, -1, data);
    }
#endif // #if WXAUTOEXCEL_USE_CHARTS
}

void MyFrame::AddHyperlinksData(wxExcelWorksheet& sheet, const wxTreeItemId& sheetId)
{
    wxExcelHyperlinks links = sheet.GetHyperlinks();

    if ( !links ) 
    {
        wxLogError(_("Failed to obtain Hyperlinks object."));
        return;
    }        
    
    wxTreeItemId linksId;
    MyTreeItemData* data = new MyTreeItemData();        

    ExcelSpy::GetHyperlinksData(links, data->m_xlData);
    linksId = m_treeCtrl->AppendItem(sheetId, _("Hyperlinks"), -1, -1, data);
    long count = links.GetCount();

    for (long l = 1; l <= count; l++ )
    {
        wxExcelHyperlink link = links[l];

        data = new MyTreeItemData();
        ExcelSpy::GetHyperlinkData(link, data->m_xlData);
        m_treeCtrl->AppendItem(linksId, link.GetName(), -1, -1, data);
    }
}


wxTreeItemId MyFrame::AppendApplicationData(const wxTreeItemId& id)
{
    if ( !m_app ) 
        return wxTreeItemId();

    MyTreeItemData* data = new MyTreeItemData();        
    ExcelSpy::GetApplicationData(m_app, data->m_xlData);    
    wxTreeItemId appId = m_treeCtrl->AppendItem(id, _("Application"), -1, -1, data);       

    data = new MyTreeItemData();        
    ExcelSpy::GetInternationalData(m_app, data->m_xlData);    
    m_treeCtrl->AppendItem(appId, _("International"), -1, -1, data);       
    
    data = new MyTreeItemData();        
    ExcelSpy::GetRecentFilesData(m_app, data->m_xlData);    
    m_treeCtrl->AppendItem(appId, _("RecentFiles"), -1, -1, data);       

    return appId;
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
