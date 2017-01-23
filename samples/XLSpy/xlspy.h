#include <wx/wx.h>
#include <wx/filedlg.h>
#include <wx/treectrl.h>
#include <wx/listctrl.h>

#include <wx/msw/ole/oleutils.h> 

#include <wx/wxAutoExcel.h>

using namespace wxAutoExcel;

class MyFrame : public wxFrame
{
public:
    MyFrame();    
private:    
     enum 
     { 
        ID_OPEN_SAMPLE_FILE = wxID_HIGHEST + 1
     };

    wxTreeCtrl* m_treeCtrl;
    wxListCtrl* m_listCtrl;
    
    wxExcelApplication m_app;
    wxExcelWorkbook m_workbook;
    
    void OnOpenFile(wxCommandEvent& evt);
    void OnOpenSampleFile(wxCommandEvent& evt);
    void OnQuit(wxCommandEvent& evt);   
    void OnClose(wxCloseEvent& evt);

    void OnTreeCtrlSelChanged(wxTreeEvent& evt);        

    bool CreateExcelInstance();

    void OpenFile(const wxString& name);
    
    void AddWorkbookData(const wxTreeItemId& id);
    
    void AddSheetsData(const wxTreeItemId& id);    
    void AddWorksheetsData(const wxTreeItemId& id);    
    void AddChartsData(const wxTreeItemId& id);
    
    void AddOLEObjectsData(wxExcelWorksheet& sheet, const wxTreeItemId& sheetId);     
    void AddShapesData(wxExcelWorksheet& sheet, const wxTreeItemId& sheetId);    
    void AddChartObjectsData(wxExcelWorksheet& sheet, const wxTreeItemId& sheetId);
    void AddHyperlinksData(wxExcelWorksheet& sheet, const wxTreeItemId& sheetId);
        
    wxTreeItemId AppendApplicationData(const wxTreeItemId& id);
};

class MyApp : public wxApp
{
public:	
    virtual bool OnInit();    
};
