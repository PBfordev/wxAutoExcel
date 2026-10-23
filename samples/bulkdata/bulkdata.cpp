/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2014 PB <pbfordev@gmail.com>
// License:     MIT license
/////////////////////////////////////////////////////////////////////////////


/**********************************************************

wxAutoExcel Bulkdata sample demonstrates the efficient ways for transferring
huge amounts of data between MS Excel and your application:
using wxExcelRangeData, see MyFrame::OnWriteRangeData() and
MyFrame::ReadRangeData(), or using SAFEARRAYs directly, see
MyFrame::OnWriteSafeArray() and MyFrame::ReadSafeArray().
It also manifests how to write values the default way,
using wxVariant with list type (MyFrame::OnWriteVariantList())
and allows you to see how much less efficient it is for larger data sets.

The values written can be of several types, selected in the Values menu.
They start at 1 and increase in each following cell, row by row.

**********************************************************/


#include <wx/wx.h>
#include <wx/numdlg.h>
#include <wx/busyinfo.h>
#include <wx/stopwatch.h>
#include <wx/intl.h>
#include <wx/numformatter.h>
#include <wx/iconbndl.h>
#include <wx/longlong.h>

#include <wx/msw/ole/oleutils.h>
#include <wx/msw/ole/safearray.h>

#include <wx/wxAutoExcel.h>

using namespace wxAutoExcel;

// The dates start at 2000-01-01, its Julian day number at noon UTC is 2451545.
// Each date is one day after the previous one, after 9999-12-31, the last
// date Excel supports, they start again at 2000-01-01. The serial number of
// 9999-12-31 is 2958465, that of 2000-01-01 is 36526.
static const double firstDateJDN = 2451545.;
static const long maxDates = 2958465 - 36526 + 1;

class MyFrame : public wxFrame
{
public:
    MyFrame();
private:
    // the types of the values written
    enum ValueType
    {
        Value_String,
        Value_Double,
        Value_Long,
        Value_Date,
        Value_Currency,
        Value_Bool,
        // the columns cycle through the types above
        Value_Mixed,
        Value_Count
    };

    enum
    {
        ID_GetNumCols = wxID_HIGHEST + 1,
        ID_GetNumRows,
        ID_WriteRangeData,
        ID_WriteSafeArray,
        ID_WriteVariantList,
        ID_ValueFirst,
        ID_ValueLast = ID_ValueFirst + Value_Count - 1
    };

    // returns the identifier of the menu item selecting the value type
    static int GetValueMenuId(ValueType valueType) { return ID_ValueFirst + static_cast<int>(valueType); }

    static const long maxColsSmall = 255;
    static const long maxRowsSmall = 65536;
    static const long maxColsLarge = 16384;
    static const long maxRowsLarge = 1048576;


    long m_numCols, m_numRows;
    bool m_supportsLargeWorksheets;

    ValueType m_valueType;
    // the length of the strings, all have the same, e.g. "0001" to "1000"
    int m_stringWidth;

    void OnGetNumCols(wxCommandEvent& event);
    void OnGetNumRows(wxCommandEvent& event);
    void OnValueType(wxCommandEvent& event);

    void CheckAndAdjustLimits();

    void OnWriteRangeData(wxCommandEvent& event);
    void OnWriteSafeArray(wxCommandEvent& event);
    void OnWriteVariantList(wxCommandEvent& event);

    void OnQuit(wxCommandEvent& event);

    bool SetupExcel(wxExcelApplication& app, wxExcelWorkbook& workbook, wxExcelWorksheet& worksheet);
    wxExcelRange WriteHeader(wxExcelWorksheet& worksheet);

    void ReadRangeData(wxExcelApplication& app, wxExcelRange& dataRange);
    void ReadSafeArray(wxExcelApplication& app, wxExcelRange& dataRange);

    wxLongLong_t GetNumValues() const { return static_cast<wxLongLong_t>(m_numCols) * m_numRows; }
    wxString GetValueTypeName() const;
    ValueType GetColumnValueType(long column) const;
    void PrepareValues();
    wxVariant MakeValue(long row, long column) const;
    void FormatColumns(wxExcelRange& dataRange);
    static wxString GetNumberFormat(ValueType valueType);
    static wxString DescribeValue(const wxVariant& value);
};


MyFrame::MyFrame()
: wxFrame(NULL, wxID_ANY, _("wxAutoExcel bulkdata sample"))
{
    SetIcons(wxIconBundle("appIcon", NULL));

    wxMenu *menu = new wxMenu;
    menu->Append(ID_GetNumCols, _("Select number of &columns to write...\tCtrl+C"));
    menu->Append(ID_GetNumRows, _("Select number of &rows to write...\tCtrl+R"));
    menu->AppendSeparator();
    menu->Append(ID_WriteRangeData, _("Write data to MS Excel using wx&ExcelRangeData!\tCtrl+E"));
    menu->Append(ID_WriteSafeArray, _("Write data to MS Excel using &SAFEARRAY!\tCtrl+S"));
    menu->Append(ID_WriteVariantList, _("&Write data to MS Excel using wxVariant&List!\tCtrl+L"));
    menu->AppendSeparator();
    menu->Append(wxID_EXIT, _("E&xit"));

    wxMenu *valueMenu = new wxMenu;
    valueMenu->AppendRadioItem(GetValueMenuId(Value_String), _("&Strings (\"0001\", \"0002\", ...)"));
    valueMenu->AppendRadioItem(GetValueMenuId(Value_Double), _("&Doubles (1.5, 2.5, ...)"));
    valueMenu->AppendRadioItem(GetValueMenuId(Value_Long), _("&Longs (1, 2, ...)"));
    valueMenu->AppendRadioItem(GetValueMenuId(Value_Date), _("D&ates (2000-01-01, 2000-01-02, ...)"));
    valueMenu->AppendRadioItem(GetValueMenuId(Value_Currency), _("&Currency (1.25, 2.25, ...)"));
    valueMenu->AppendRadioItem(GetValueMenuId(Value_Bool), _("&Booleans (true, false, ...)"));
    valueMenu->AppendRadioItem(GetValueMenuId(Value_Mixed), _("&Mixed (each column of another type)"));

    wxMenuBar *menuBar = new wxMenuBar();
    menuBar->Append(menu, _("&Sample"));
    menuBar->Append(valueMenu, _("&Values"));
    SetMenuBar(menuBar);

    Bind(wxEVT_COMMAND_MENU_SELECTED, &MyFrame::OnGetNumCols, this, ID_GetNumCols);
    Bind(wxEVT_COMMAND_MENU_SELECTED, &MyFrame::OnGetNumRows, this, ID_GetNumRows);
    Bind(wxEVT_COMMAND_MENU_SELECTED, &MyFrame::OnWriteRangeData, this, ID_WriteRangeData);
    Bind(wxEVT_COMMAND_MENU_SELECTED, &MyFrame::OnWriteSafeArray, this, ID_WriteSafeArray);
    Bind(wxEVT_COMMAND_MENU_SELECTED, &MyFrame::OnWriteVariantList, this, ID_WriteVariantList);
    Bind(wxEVT_COMMAND_MENU_SELECTED, &MyFrame::OnValueType, this, ID_ValueFirst, ID_ValueLast);
    Bind(wxEVT_COMMAND_MENU_SELECTED, &MyFrame::OnQuit, this, wxID_EXIT);

    m_numCols =  200;
    m_numRows = 5000;

    m_valueType = Value_Double;
    valueMenu->Check(GetValueMenuId(m_valueType), true);
    m_stringWidth = 1;

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

void MyFrame::OnValueType(wxCommandEvent& event)
{
    m_valueType = static_cast<ValueType>(event.GetId() - ID_ValueFirst);
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
        msg += wxString::Format(_("Number of columns adjusted to %ld\n"), m_numCols);
    }
    if ( rows != m_numRows )
    {
        m_numRows = rows;
        msg += wxString::Format(_("Number of rows adjusted to %ld\n"), m_numRows);
    }
    if ( !msg.empty() )
    {
        wxMessageBox(msg);
    }
}

wxString MyFrame::GetValueTypeName() const
{
    static const char* const names[] =
        { "strings", "doubles", "longs", "dates", "currency values", "booleans", "mixed values" };

    return names[m_valueType];
}

MyFrame::ValueType MyFrame::GetColumnValueType(long column) const
{
    if ( m_valueType == Value_Mixed )
        return static_cast<ValueType>(column % Value_Mixed);
    return m_valueType;
}

// Must be called before MakeValue().
void MyFrame::PrepareValues()
{
    // all strings have the length of the largest number
    m_stringWidth = static_cast<int>(wxString::Format("%" wxLongLongFmtSpec "d", GetNumValues()).length());
}

// Returns the value of the cell in the given row and column, both starting at 0.
// The values start at 1 and increase in each following cell, row by row.
wxVariant MyFrame::MakeValue(long row, long column) const
{
    const wxLongLong_t index = static_cast<wxLongLong_t>(row) * m_numCols + column;

    switch ( GetColumnValueType(column) )
    {
        case Value_String:
            return wxString::Format("%0*" wxLongLongFmtSpec "d", m_stringWidth, index + 1);
        case Value_Double:
            return static_cast<double>(index) + 1.5;
        case Value_Long:
            // Excel stores all numbers as doubles, so these are read back as doubles
            return static_cast<long>(index + 1);
        case Value_Date:
            {
                // the date of the day at noon UTC, at midnight local time,
                // which daylight saving time does not shift to another hour
                const wxDateTime::Tm tm = wxDateTime(firstDateJDN + index % maxDates).GetTm(wxDateTime::UTC);

                return wxDateTime(tm.mday, tm.mon, tm.year);
            }
        case Value_Currency:
            {
                // CURRENCY is a 64-bit integer scaled by 10,000
                CURRENCY currency;

                currency.int64 = (index + 1) * 10000 + 2500;
                return wxVariant(new wxVariantDataCurrency(currency));
            }
        case Value_Bool:
            return index % 2 == 0;
        default:
            wxFAIL_MSG("Unexpected value type");
    }
    return wxVariant();
}

// Returns the number format for the values of the given type, or an empty
// string for those which Excel formats itself (dates and currency values)
// or which need no format.
wxString MyFrame::GetNumberFormat(ValueType valueType)
{
    switch ( valueType )
    {
        case Value_String:
            // the text format keeps e.g. "0001" a string,
            // which Excel would convert to the number 1 otherwise
            return "@";
        case Value_Double:
            return "#,##0.00";
        case Value_Long:
            return "#,##0";
        default:
            break;
    }
    return wxString();
}

// Sets the number formats of the columns before the values are written.
void MyFrame::FormatColumns(wxExcelRange& dataRange)
{
    // format all the columns at once when they are of the same type
    if ( m_valueType != Value_Mixed )
    {
        const wxString format = GetNumberFormat(m_valueType);

        if ( !format.empty() )
            dataRange.SetNumberformat(format);
        return;
    }

    for ( long column = 0; column < m_numCols; column++ )
    {
        const wxString format = GetNumberFormat(GetColumnValueType(column));

        if ( !format.empty() )
            dataRange.GetColumns(column + 1).SetNumberformat(format);
    }
}

// Returns the value with its type, e.g. "1.5 (double)".
wxString MyFrame::DescribeValue(const wxVariant& value)
{
    wxString str = value.MakeString();

    if ( value.IsType("bool") )
        str = value.GetBool() ? "true" : "false";

    return wxString::Format("%s (%s)", str, value.GetType());
}

void MyFrame::OnWriteRangeData(wxCommandEvent& WXUNUSED(event))
{
    wxExcelApplication app;
    wxExcelWorkbook workbook;
    wxExcelWorksheet worksheet;

    if ( !SetupExcel(app, workbook, worksheet) )
        return;
    CheckAndAdjustLimits();
    PrepareValues();

    wxExcelRange headerRange = WriteHeader(worksheet);
    wxExcelRange dataRange;
    bool written = false;

    wxStopWatch sw;
    long timeGenerating = 0, timeTotal = 0;

    wxMessageBox(_("After the data are written to MS Excel, switch back to the bulkdata sample application."));

    { // new scope for wxBusyInfo
        wxBusyInfo wait(wxString::Format(_("Attempting to write %s %s (%s columns x %s rows)..."),
            wxNumberFormatter::ToString(GetNumValues()),
            GetValueTypeName(),
            wxNumberFormatter::ToString(m_numCols),
            wxNumberFormatter::ToString(m_numRows)));

        // generate data, the cells are numbered from 0
        wxExcelRangeData data(m_numRows, m_numCols);

        sw.Start();
        for ( long row = 0; row < m_numRows; row++ )
        {
            for ( long col = 0; col < m_numCols; col++ )
                data(row, col) = MakeValue(row, col);
        }
        timeGenerating = sw.Time();

        // create a range with m_numCols columns and m_numRows rows,
        // its size must be the same as that of the data
        dataRange = headerRange.GetOffset(1, 0).GetResize(&m_numRows);
        FormatColumns(dataRange);

        // write data to Excel, using Value instead of the default Value2,
        // so that Excel formats the cells with dates and currency values
        written = dataRange.SetValues_(data, false);
        timeTotal = sw.Time();
        if ( written )
            worksheet.GetUsedRange().GetColumns().AutoFit();
    }

    app.SetVisible(true); // display MS Excel window

    if ( !written )
    {
        wxLogError(_("Error writing the data to the sheet."));
        return;
    }

    wxMessageBox( wxString::Format(_("Range.SetValues_() using wxExcelRangeData\n---\nTime: total %s ms, creating data %s ms\n(%s %s: %s columns, %s rows, address %s)"),
        wxNumberFormatter::ToString(timeTotal), wxNumberFormatter::ToString(timeGenerating),
        wxNumberFormatter::ToString(GetNumValues()),
        GetValueTypeName(),
        wxNumberFormatter::ToString(m_numCols),
        wxNumberFormatter::ToString(m_numRows),
        dataRange.GetAddress()
        ) );

    if ( wxMessageBox(_("Attempt to obtain the copied data back from MS Excel?"),
          _("Confirm"), wxYES_NO) == wxYES )
    {
        ReadRangeData(app, dataRange);
    }
}

void MyFrame::ReadRangeData(wxExcelApplication& app, wxExcelRange& dataRange)
{
    wxStopWatch sw;
    long timeTotal;
    wxExcelRangeData data;
    bool read;

    wxMessageBox(_("After the data are read from MS Excel, switch back to the bulkdata sample application."));

    app.SetVisible(false); // hide MS Excel window

    { // new scope for wxBusyInfo
        wxBusyInfo wait(wxString::Format(_("Attempting to read %s values (%s columns x %s rows)..."),
            wxNumberFormatter::ToString(GetNumValues()),
            wxNumberFormatter::ToString(m_numCols),
            wxNumberFormatter::ToString(m_numRows)));

        // read Value instead of Value2, so that dates and currency
        // values are not returned as numbers
        sw.Start();
        read = dataRange.GetValues_(data, false);
        timeTotal = sw.Time();
    }
    app.SetVisible(true); // display MS Excel window

    if ( !read )
    {
        wxLogError(_("Error during reading the data from the sheet."));
        return;
    }

    const size_t lastRow = data.GetRowCount() - 1;
    const size_t lastColumn = data.GetColumnCount() - 1;

    wxMessageBox( wxString::Format(_("Range.GetValues_() using wxExcelRangeData\n---\nTime = %s ms\n(%s values: %s columns, %s rows,\n"
        "first value = %s, last value = %s,\nRange.Address %s)"),
        wxNumberFormatter::ToString(timeTotal),
        wxNumberFormatter::ToString(static_cast<wxLongLong_t>(data.GetRowCount() * data.GetColumnCount())),
        wxNumberFormatter::ToString(static_cast<long>(data.GetColumnCount())),
        wxNumberFormatter::ToString(static_cast<long>(data.GetRowCount())),
        DescribeValue(data(0, 0)), DescribeValue(data(lastRow, lastColumn)),
        dataRange.GetAddress()
        ) );
}

void MyFrame::OnWriteSafeArray(wxCommandEvent& WXUNUSED(event))
{
    wxExcelApplication app;
    wxExcelWorkbook workbook;
    wxExcelWorksheet worksheet;

    if ( !SetupExcel(app, workbook, worksheet) )
        return;
    CheckAndAdjustLimits();
    PrepareValues();

    wxExcelRange headerRange = WriteHeader(worksheet);
    wxExcelRange dataRange;

    wxStopWatch sw;
    long timeGenerating = 0, timeTotal = 0;

    wxMessageBox(_("After the data are written to MS Excel, switch back to the bulkdata sample application."));

    { // new scope for wxBusyInfo
        wxBusyInfo wait(wxString::Format(_("Attempting to write %s %s (%s columns x %s rows)..."),
            wxNumberFormatter::ToString(GetNumValues()),
            GetValueTypeName(),
            wxNumberFormatter::ToString(m_numCols),
            wxNumberFormatter::ToString(m_numRows)));

        // generate data

        SAFEARRAYBOUND bounds[2]; // 2 dimensions
        long indices[2];

        // copy values as VARIANTs, so that they can be of any type;
        // for values of a single type, e.g. VT_R8 for doubles can be used
        wxSafeArray<VT_VARIANT> safeArray;

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
                if ( !safeArray.SetElement(indices, MakeValue(row, col)) )
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
        FormatColumns(dataRange);

        dataRange = wxVariant(new wxVariantDataSafeArray(safeArray.Detach()));
        if ( dataRange ) // we succeeded to write the data
        {
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
    if ( dataRange.GetCount() !=  GetNumValues() )
    {
        wxLogError(_("Failed to write all the data to the sheet."));
        return;
    }
    else
    {
        wxMessageBox( wxString::Format(_("Range.SetValue() using SAFEARRAY\n---\nTime: total %s ms, creating array %s ms\n(%s %s: %s columns, %s rows, address %s)"),
            wxNumberFormatter::ToString(timeTotal), wxNumberFormatter::ToString(timeGenerating),
            wxNumberFormatter::ToString(GetNumValues()),
            GetValueTypeName(),
            wxNumberFormatter::ToString(m_numCols),
            wxNumberFormatter::ToString(m_numRows),
            dataRange.GetAddress()
            ) );

        if ( GetNumValues() > 1 // if there is just one value, it won't be returned as an array but as a simple wxVariant
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
            wxNumberFormatter::ToString(GetNumValues()),
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
        wxLongLong_t count = 1;

        for ( size_t i = 1; i <= dims; i++ )
            count *= safeArray.GetCount(i);

        if ( count != GetNumValues() )
        {
            wxLogError(_("Failed to read all the values from MS Excel."));
            return;
        }

        wxVariant valueFirst, valueLast;
        long indices[2];

        safeArray.GetLBound(1, indices[0]);
        safeArray.GetLBound(dims, indices[1]);
        safeArray.GetElement(indices, valueFirst);

        safeArray.GetUBound(1, indices[0]);
        safeArray.GetUBound(dims, indices[1]);
        safeArray.GetElement(indices, valueLast);

        wxMessageBox( wxString::Format(_("Range.GetValue() using SAFEARRAY\n---\nTime = %s ms\n(%s values: %s columns, %s rows,\n"
            "first value = %s, last value = %s,\nRange.Address %s)"),
            wxNumberFormatter::ToString(timeTotal),
            wxNumberFormatter::ToString(GetNumValues()),
            wxNumberFormatter::ToString(m_numCols),
            wxNumberFormatter::ToString(m_numRows),
            DescribeValue(valueFirst), DescribeValue(valueLast),
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
    PrepareValues();

    wxExcelRange headerRange = WriteHeader(worksheet);

    wxExcelRange dataRange;

    // generate data
    wxVariant data;
    wxStopWatch sw;
    long timeTotal;
    wxString msg;

    { // new scope for wxBusyInfo
        wxBusyInfo wait(wxString::Format("Attempting to write %s %s (%s columns x %s rows)...",
                wxNumberFormatter::ToString(GetNumValues()),
                GetValueTypeName(),
                wxNumberFormatter::ToString(m_numCols),
                wxNumberFormatter::ToString(m_numRows)));

        // wxVariant list doesn't support 2-dimensional array
        // so let's copy it row by row. Depending on how
        // the data are laid out, it may be more efficient
        // to copy them by columns instead (not shown here)

        dataRange = headerRange.GetOffset(1, 0).GetResize(&m_numRows);
        FormatColumns(dataRange);
        dataRange = headerRange.GetOffset(1, 0);

        sw.Start();
        for ( long row = 0; row < m_numRows; row++ )
        {
            data.ClearList();
            for ( long col = 0; col < m_numCols; col++ )
            {
                data.Append(MakeValue(row, col));
            }
            dataRange = data;
            dataRange = dataRange.GetOffset(1, 0);
        }
        timeTotal = sw.Time();

        worksheet.GetUsedRange().GetColumns().AutoFit();
    }

    app.SetVisible(true); // display MS Excel window

    msg.Printf("Range.SetValue() using wxVariant list\n---\nTime = %s ms\n(%s %s: %s columns, %s rows)",
        wxNumberFormatter::ToString(timeTotal),
        wxNumberFormatter::ToString(GetNumValues()),
        GetValueTypeName(),
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
    wxMessageBox(_("It appears you are running the debug build of the sample.\n"
        "Data operations are very likely to be (much) slower than in the release build."),
        _("Warning"), wxICON_EXCLAMATION | wxOK);
#endif

    m_locale.Init(); // needed only because of getting thousand separator with wxNumberFormatter

    MyFrame* frame = new MyFrame();
    frame->Show();

    wxLog::AddTraceMask(wxTRACE_AutoExcel);

    return true;
}

wxIMPLEMENT_APP(MyApp);
