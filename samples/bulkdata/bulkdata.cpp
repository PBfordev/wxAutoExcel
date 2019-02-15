/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2014 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////


/**********************************************************

wxAutoExcel Bulkdata sample demonstrates the efficient way for transferring
huge amounts of data between MS Excel and your application
based on using SAFEARRAYs. See MyFrame::OnWriteSafeArray()
and MyFrame::ReadSafeArray() methods.
It also manifests how to write values the default way,
using wxVariant with list type (MyFrame::OnWriteVariantList())
and allows you to see how much less efficient it is for larger data sets.

**********************************************************/


#include <wx/wx.h>
#include <wx/numdlg.h>
#include <wx/busyinfo.h>
#include <wx/stopwatch.h>
#include <wx/intl.h>
#include <wx/numformatter.h>
#include <wx/iconbndl.h>

#include <wx/msw/ole/oleutils.h> 
#include <wx/msw/ole/safearray.h> 

#include <wx/wxAutoExcel.h>

using namespace wxAutoExcel;

class MyFrame : public wxFrame
{
public:
    MyFrame();    
private:
    enum
    {
        ID_GetNumCols = wxID_HIGHEST + 1,
        ID_GetNumRows,         
        ID_WriteSafeArray,
        ID_WriteVariantList,
    };
    static const long maxColsSmall = 255;
    static const long maxRowsSmall = 65536;
    static const long maxColsLarge = 16384;    
    static const long maxRowsLarge = 1048576;
    

    long m_numCols, m_numRows;
    bool m_supportsLargeWorksheets;    
    
    void OnGetNumCols(wxCommandEvent& event);    
    void OnGetNumRows(wxCommandEvent& event);    

    void CheckAndAdjustLimits();
    
    void OnWriteSafeArray(wxCommandEvent& event);
    void OnWriteVariantList(wxCommandEvent& event);

    void OnQuit(wxCommandEvent& event);    
    
    bool SetupExcel(wxExcelApplication& app, wxExcelWorkbook& workbook, wxExcelWorksheet& worksheet);
    wxExcelRange WriteHeader(wxExcelWorksheet& worksheet);

    void ReadSafeArray(wxExcelApplication& app, wxExcelRange& dataRange);
};


MyFrame::MyFrame()
: wxFrame(NULL, wxID_ANY, _("wxAutoExcel bulkdata sample"))
{
    SetIcons(wxIconBundle("appIcon", NULL));
    
    wxMenu *menu = new wxMenu;
    menu->Append(ID_GetNumCols, _("Select number of &columns to write...\tCtrl+C"));
    menu->Append(ID_GetNumRows, _("Select number of &rows to write...\tCtrl+R"));
    menu->AppendSeparator();
    menu->Append(ID_WriteSafeArray, _("Write data to MS Excel using &SAFEARRAY!\tCtrl+S"));
    menu->Append(ID_WriteVariantList, _("&Write data to MS Excel using wxVariant&List!\tCtrl+L"));
    menu->AppendSeparator();
    menu->Append(wxID_EXIT, _("E&xit"));

    wxMenuBar *menuBar = new wxMenuBar();
    menuBar->Append(menu, _("&Sample"));
    SetMenuBar(menuBar);

    Bind(wxEVT_COMMAND_MENU_SELECTED, &MyFrame::OnGetNumCols, this, ID_GetNumCols);
    Bind(wxEVT_COMMAND_MENU_SELECTED, &MyFrame::OnGetNumRows, this, ID_GetNumRows);
    Bind(wxEVT_COMMAND_MENU_SELECTED, &MyFrame::OnWriteSafeArray, this, ID_WriteSafeArray);
    Bind(wxEVT_COMMAND_MENU_SELECTED, &MyFrame::OnWriteVariantList, this, ID_WriteVariantList);
    Bind(wxEVT_COMMAND_MENU_SELECTED, &MyFrame::OnQuit, this, wxID_EXIT);

    m_numCols =  200;
    m_numRows = 5000;

    wxExcelApplication app;
    wxExcelWorkbook workbook;
    wxExcelWorksheet worksheet;

    m_supportsLargeWorksheets = false;

    if ( SetupExcel(app, workbook, worksheet) )
    {        
        if ( app.IsOk_() )
            app.Quit();        
    }    
}


bool MyFrame::SetupExcel(wxExcelApplication& app, wxExcelWorkbook& workbook, wxExcelWorksheet& worksheet)
{
   // first create an instance of MS Excel
    app = wxExcelApplication::CreateInstance();
    if ( !app ) 
    {
        wxLogError(_("Failed to create an instance of MS Excel application."));
        return false;
    }
    app.SetVisible(false); // hide MS Excel window       

    // add a new workbook
    workbook = app.GetWorkbooks().Add();    
    if ( !workbook ) 
    {
        wxLogError(_("Failed to create a new workbook."));
        return  false;
    }

    // Set the workbook automation object locale to US English, so we can use
    // English names for its formulas, styles etc. in the automation calls,
    // regardless of the language Excel may be localized into.
    // The end user will still see the localized ones in Excel.
    workbook.SetAutomationLCID_(wxExcelObject::lcidEnglishUS);

    m_supportsLargeWorksheets = app.IsVersionAtLeast_(wxExcelApplication::evExcel2007) && workbook.GetExcel8CompatibilityMode() == false;            

    // get the first worksheet in the newly added workbook
    worksheet = workbook.GetWorksheets()[1];
    if ( !worksheet )
    {
        wxLogError(_("Failed to obtain worksheet number 1."));
        return false;
    }

    return true;
}

wxExcelRange MyFrame::WriteHeader(wxExcelWorksheet& worksheet)
{
    wxVariant header;
    wxExcelRange range;

    // write column headers
    range = worksheet.GetRange("A1").GetResize(NULL, WXAEEP(m_numCols));
    
    header.ClearList();
    for ( long i = 0; i < m_numCols; i++ )
    {
        header.Append(wxString::Format("Column %ld", i+1));
    }            
    range = header;
    
    range.GetFont().SetBold(true);

    return range;
}

void MyFrame::OnGetNumCols(wxCommandEvent& WXUNUSED(event))
{
    long maxCols = maxColsSmall;
    long cols;
    
    if ( m_supportsLargeWorksheets )
        maxCols =  maxColsLarge;
    
    cols = wxGetNumberFromUser(wxEmptyString, "columns", _("Number of columns to write"), 
        m_numCols, 1, maxCols, this);

    if ( cols != -1 )
        m_numCols = cols;
}

void MyFrame::OnGetNumRows(wxCommandEvent& WXUNUSED(event))
{
    long maxRows = maxRowsSmall;
    long rows;
    
    if ( m_supportsLargeWorksheets )
        maxRows =  maxRowsLarge;
    
    rows = wxGetNumberFromUser(wxEmptyString, "rows", _("Number of rows to write"), 
        m_numRows, 1, maxRows, this);

    if ( rows != -1 )
        m_numRows = rows;
}

// check and adjust for the very unlikely event
// that limits changed since the last check at application start
// e.g. Compatibility mode turned off
void MyFrame::CheckAndAdjustLimits()
{
    long cols = m_numCols, rows = m_numRows;

    if ( !m_supportsLargeWorksheets )
        cols =  wxMin(m_numCols, maxColsSmall);
    if ( !m_supportsLargeWorksheets )
        rows =  wxMin(m_numRows, maxRowsSmall);

    wxString msg;
    if ( cols != m_numCols )
    {
        m_numCols = cols;
        msg.Printf(_("Number of columns adjusted to %ld\n"), m_numCols);
    }
    if ( rows != m_numRows )
    {
        m_numRows = rows;
        msg.Printf(_("Number of rows adjusted to %ld\n"), m_numRows);
    }
    if ( !msg.empty() )
    {
        wxMessageBox(msg);
    }
}

void MyFrame::OnWriteSafeArray(wxCommandEvent& WXUNUSED(event))
{
    wxExcelApplication app;
    wxExcelWorkbook workbook;
    wxExcelWorksheet worksheet;

    if ( !SetupExcel(app, workbook, worksheet) )
        return;
    CheckAndAdjustLimits();
    
    wxExcelRange headerRange = WriteHeader(worksheet);    
    wxExcelRange dataRange;
               
    wxStopWatch sw;
    long timeGenerating = 0, timeTotal = 0;    

    wxMessageBox(_("After the data are written to MS Excel, switch back to the bulkdata sample application."));

    { // new scope for wxBusyInfo
        wxBusyInfo wait(wxString::Format(_("Attempting to write %s values (%s columns x %s rows)..."),
            wxNumberFormatter::ToString(m_numCols * m_numRows),
            wxNumberFormatter::ToString(m_numCols),
            wxNumberFormatter::ToString(m_numRows)));

        // generate data
        
        double dVal = 1.;         
        SAFEARRAYBOUND bounds[2]; // 2 dimensions
        long indices[2];

        wxSafeArray<VT_R8> safeArray; // copy values as doubles
        
        bounds[0].lLbound = 0; // elements start at 0
        bounds[0].cElements = m_numRows;
        bounds[1].lLbound = 0; // elements start at 0
        bounds[1].cElements = m_numCols;
        
        if ( !safeArray.Create(bounds, 2) )
         {
            wxLogError(_("Failed to create SAFEARRAY."));
            return;
        }
        
            
        sw.Start();
        for ( long row = 0; row < m_numRows; row++ )
        {
            indices[0] = row;
            for ( long col = 0; col < m_numCols; col++ )
            {
                indices[1] = col;
                if ( !safeArray.SetElement(indices, dVal++) )
                {
                    wxLogError(_("Failed to set a SAFEARRAY element."));
                    return;
                }
            }
        }
        timeGenerating = sw.Time();
        // write data to Excel
                
        // create a range with m_numCols columns and m_numRows rows
        dataRange = headerRange.GetOffset(1, 0).GetResize(&m_numRows);
                
        dataRange = wxVariant(new wxVariantDataSafeArray(safeArray.Detach()));        
        if ( dataRange ) // we succeeded to write the data
        {        
            dataRange.SetNumberFormat("#,##0");
            timeTotal = sw.Time();
            worksheet.GetUsedRange().GetColumns().AutoFit();        
        }    
    }
           
    app.SetVisible(true); // display MS Excel window           

    if ( !dataRange)
    {
        wxLogError(_("Error writing the data to the sheet."));
    } 
    else
    if ( dataRange.GetCount() !=  m_numCols * m_numRows)
    {
        wxLogError(_("Failed to write all the data to the sheet."));
        return;
    }
    else
    {
        wxMessageBox( wxString::Format(_("Range.SetValue() using SAFEARRAY\n---\nTime: total %s ms, creating array %s ms\n(%s values: %s columns, %s rows, address %s)"),
            wxNumberFormatter::ToString(timeTotal), wxNumberFormatter::ToString(timeGenerating),
            wxNumberFormatter::ToString(m_numCols * m_numRows),
            wxNumberFormatter::ToString(m_numCols),
            wxNumberFormatter::ToString(m_numRows),
            dataRange.GetAddress() 
            ) );        

        if ( (m_numCols * m_numRows) > 1 // if there is just one value, it won't be returned as an array but as a simple wxVariant
              && wxMessageBox(_("Attempt to obtain the copied data back from MS Excel?"),
              _("Confirm"), wxYES_NO) == wxYES )
        {
            ReadSafeArray(app, dataRange);
        }      
    }
}

void MyFrame::ReadSafeArray(wxExcelApplication& app, wxExcelRange& dataRange)
{
    wxStopWatch sw;
    long timeTotal;  
    wxVariant data;        

    wxMessageBox(_("After the data are read from MS Excel, switch back to the bulkdata sample application."));
    
    app.SetVisible(false); // hide MS Excel window
    
    { // new scope for wxBusyInfo
        wxBusyInfo wait(wxString::Format("Attempting to read %s values (%s columns x %s rows)...",
            wxNumberFormatter::ToString(m_numCols * m_numRows),
            wxNumberFormatter::ToString(m_numCols),
            wxNumberFormatter::ToString(m_numRows)));

        
        dataRange.SetConvertVariantFlags_(wxOleConvertVariant_ReturnSafeArrays);                
                            
        sw.Start();
        data = dataRange.GetValue();
        timeTotal = sw.Time();        
    }           
    app.SetVisible(true); // display MS Excel window           

    if ( !dataRange )
    {
        wxLogError(_("Error during reading the data from the sheet."));
    }
    
    if ( data.GetType() != "safearray" )
    {
        wxLogError(_("Failed to read the data from the sheet as a SAFEARRAY."));
    }
    else
    {
        wxSafeArray<VT_VARIANT> safeArray;
        wxVariantDataSafeArray* const  sa = wxStaticCastVariantData(data.GetData(), wxVariantDataSafeArray);
        
        if ( !safeArray.Attach(sa->GetValue()) ) // shouldn't really ever happen here
        {
            if ( !safeArray.HasArray() )
            {
                SafeArrayDestroy(sa->GetValue()); // we have to dispose the SAFEARRAY ourselves
            }
            
            wxLogError(_("Failed to get the data from the SAFEARRAY."));
            return;
        }
     
        // just to be sure, verify the SAFEARRAY item count is the same as dataRange.Count
        // but it should not be necessary as long dataRange() evaluates to true...
        size_t dims = safeArray.GetDim();
        long count = 1;

        for ( size_t i = 1; i <= dims; i++ )
            count *= safeArray.GetCount(i);

        if ( count != dataRange.GetCount() )
        {
            wxLogError(_("Failed to read all the values from MS Excel."));
            return;
        }

        wxVariant valueFirst, valueLast;
        long indices[2];

        indices[0] = 1;
        safeArray.GetLBound(1, indices[1]);
        safeArray.GetElement(indices, valueFirst);
                
        safeArray.GetUBound(1, indices[0]);
        safeArray.GetUBound(dims, indices[1]);
        safeArray.GetElement(indices, valueLast);       

        wxMessageBox( wxString::Format(_("Range.GetValue() using SAFEARRAY\n---\nTime = %s ms\n(%s values: %s columns, %s rows, "
            "First value = %s, last value = %s\n Range.Address %s)"),
            wxNumberFormatter::ToString(timeTotal),
            wxNumberFormatter::ToString(m_numCols * m_numRows),
            wxNumberFormatter::ToString(m_numCols),
            wxNumberFormatter::ToString(m_numRows),
            valueFirst.GetString(), valueLast.GetString(),
            dataRange.GetAddress() 
            ) );     
    }    

}


void MyFrame::OnWriteVariantList(wxCommandEvent& WXUNUSED(event))
{
    wxExcelApplication app;
    wxExcelWorkbook workbook;
    wxExcelWorksheet worksheet;

    if ( !SetupExcel(app, workbook, worksheet) )
        return;
    CheckAndAdjustLimits();

    wxExcelRange headerRange = WriteHeader(worksheet);

    wxExcelRange dataRange;    
    
    // generate data
    wxVariant data;    
    double dVal = 1.;    
    wxStopWatch sw;
    long timeTotal;
    wxString msg;
               
    { // new scope for wxBusyInfo
        wxBusyInfo wait(wxString::Format("Attempting to write %s values (%s columns x %s rows)...",
                wxNumberFormatter::ToString(m_numCols * m_numRows),
                wxNumberFormatter::ToString(m_numCols),
                wxNumberFormatter::ToString(m_numRows)));

        // wxVariant list doesn't support 2-dimensional array
        // so let's copy it row by row. Ddepending on how 
        // the data are laid out, it may be more efficient 
        // to copy them by columns instead (not shown here)
            
        dataRange = headerRange.GetOffset(1, 0);    
        
        sw.Start();
        for ( long row = 0; row < m_numRows; row++ )
        {        
            data.ClearList();
            for ( long col = 0; col < m_numCols; col++ )
            {
                data.Append(dVal++);
            }
            dataRange = data;
            dataRange.SetNumberFormat("#,##0");
            dataRange = dataRange.GetOffset(1, 0);            
        }                    
        timeTotal = sw.Time();    
        
        worksheet.GetUsedRange().GetColumns().AutoFit();
    }

    app.SetVisible(true); // display MS Excel window       

    msg.Printf("Range.SetValue() using wxVariant list\n---\nTime = %s ms\n(%s values: %s columns, %s rows)",
        wxNumberFormatter::ToString(timeTotal),
        wxNumberFormatter::ToString(m_numCols * m_numRows),
        wxNumberFormatter::ToString(m_numCols),
        wxNumberFormatter::ToString(m_numRows) );    
    wxMessageBox(msg);        
}


void MyFrame::OnQuit(wxCommandEvent& WXUNUSED(event))
{
    Close(true);
}



class MyApp : public wxApp
{
public:	
    virtual bool OnInit();    
private:
    wxLocale m_locale;
};


bool MyApp::OnInit()
{
    if (!wxApp::OnInit())
        return false;       	

#ifndef NDEBUG
    wxMessageBox(_("It appears you are running the debg build of the sample.\n"
        "Data operations are very likely to be (much) slower than in the release build."),
        _("Warning"), wxICON_EXCLAMATION | wxOK);
#endif

    m_locale.Init(); // needed only because of getting thousand separator with wxNumberFormatter

    MyFrame* frame = new MyFrame();
    frame->Show();

    wxLog::AddTraceMask(wxTRACE_AutoExcel);                                  

    return true;
}

IMPLEMENT_APP(MyApp)
