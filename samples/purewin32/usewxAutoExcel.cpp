#include <wx/wx.h>
#include <wx/msw/ole/oleutils.h>

#include <wx/wxAutoExcel.h>

#include "usewxAutoExcel.h"

/**
     Passes messages from wxLog*() calls from within wxAutoExcel to a class
     derived from MyLogger, see MyLogger declaration in usewxAutoExcel.h.
**/
class MywxLogger : public wxLog
{
public:
    MywxLogger(MyLogger* callback) : m_callback(callback) {}
protected:
    void DoLogTextAtLevel(wxLogLevel level, const wxString& msg) wxOVERRIDE
    {
        if ( m_callback )
            m_callback->Log((int)level, msg.wc_str());
    }
private:
    MyLogger* m_callback;
};

/**
     Initializes wxWidgets and redirects all wxLog*() calls to logger
**/
class MywxInitializer
{
public:
    MywxInitializer(MyLogger* logger)
    {
        if ( wxApp::GetInstance() )
        {
            // already initialized?
            m_initialized = false;
            wxFAIL_MSG("wxWidgets have already been initialized!");
            return;
        }

        wxApp::SetInstance(new wxApp());

        m_initialized = wxEntryStart(0, NULL);
        if ( m_initialized && logger)
            wxLog::SetActiveTarget(new MywxLogger(logger));
    }

    ~MywxInitializer()
    {
        if ( m_initialized )
            wxEntryCleanup();
    }

    operator bool() const { return m_initialized; }
private:
    bool m_initialized;
};

// Converts a double to wxVariant containing CURRENCY
wxVariant DoubleToCurrencyVariant(double d, bool* success = NULL)
{
    CURRENCY cy = {0};

    HRESULT hr = VarCyFromR8(d, &cy);
    if ( success )
        *success = SUCCEEDED(hr);

    return wxVariant(new wxVariantDataCurrency(cy));
}

/**
    Function UsewxAutoExcel():
    (1) Initializes wxWidgets and if logger is non-NULL, redirects wxLog*() calls to it.
    (2) The code between lines
        wxLogMessage("*** Initializing  wxAutoExcel...");
        and
        wxLogMessage("*** Shutting down wxAutoExcel...");
        is exactly same as in MyFrame::OnCreateWorksheet() in the Minimal sample.
    (3) Shuts down wxWidgets.

    if enableLogTimeStap is true, the log message will contain a timestamp
    if enablewxAutoExcelTrace is true, wxAutoExcel debug messages will be produced in the Debug builds
**/

using namespace wxAutoExcel;

bool UsewxAutoExcel(MyLogger* logger, bool enableLogTimeStamp, bool enablewxAutoExcelTrace)
{
    MywxInitializer initializer(logger);

    if ( !initializer )
        return false;

    if ( !enableLogTimeStamp )
        wxLog::DisableTimestamp();

    if ( enablewxAutoExcelTrace )
        wxLog::AddTraceMask(wxTRACE_AutoExcel);

    wxLogMessage("*** Initializing  wxAutoExcel...");

    // first create an MS Excel instance
    wxExcelApplication app = wxExcelApplication::CreateInstance();
    if ( !app )
    {
        wxLogError(_("Failed to create an instance of MS Excel application."));
        return false;
    }
    app.SetVisible(true); // display MS Excel window

    // add a new workbook
    wxExcelWorkbook workbook = app.GetWorkbooks().Add();
    if ( !workbook )
    {
        wxLogError(_("Failed to create a new workbook."));
        return false;
    }

    // Set the workbook automation object locale to US English, so we can use
    // English names for its formulas, styles etc. in the automation calls,
    // regardless of the language Excel may be localized into. See FAQ for more information.
    // The end user will still see the localized ones in Excel.
    workbook.SetAutomationLCID_(wxExcelObject::lcidEnglishUS);

    // get the first worksheet in the newly added workbook,
    // remember that indices in MS Office collections start at 1, NOT 0
    wxExcelWorksheet worksheet = workbook.GetWorksheets()[1];
    if ( !worksheet )
    {
        wxLogError(_("Failed to obtain worksheet number 1."));
        return false;
    }
    // change worksheet name
    worksheet.SetName("A Very Silly Table");

    wxVariant variant;
    wxExcelRange range;

    // write sheet headers
    range = worksheet.GetRange("A1:E1");
    variant.ClearList();
    variant.Append("Code");
    variant.Append("Date");
    variant.Append("Quantity");
    variant.Append("Price");
    variant.Append("Subtotal");
    // set cell values
    range.SetValue(variant);
    // set headers to bold
    range.GetFont().SetBold(true);
    // center headers
    range.SetHorizontalAlignment(xlCenter);

    // write the first row of values

    // first shift the range one row down
    range = range.GetOffset(1);

    variant.ClearList();
    variant.Append("ABC0123");
    variant.Append(wxDateTime::Today());
    variant.Append(3L);
    variant.Append(DoubleToCurrencyVariant(10.5));
    variant.Append("=C2*D2");

    // set cell values

    // wxExcelRange has operator()(const wxVariant&) overloaded
    // so it behaves as if you called SetValue(variant)
    range = variant;

    // second row of values
    // shift the range one row down again
    range = range.GetOffset(1);
    // write the second row of values
    variant.ClearList();
    variant.Append("XYZ4567");
    variant.Append(wxDateTime::Today());
    variant.Append(5L);
    variant.Append(DoubleToCurrencyVariant(8.25));
    variant.Append("=C3*D3");
    range = variant;

    // shift the range one row down again
    range = range.GetOffset(1);

    // GetRange() uses addresses relative to range,
    // so e.g. GetRange("A4:E4").GetRange("A1") returns
    // a range with a worksheet absolute address A4
    range.GetRange("A1") = "TOTAL";

    range = range.GetRange("E1"); // again, range-relative address
    range.SetFormula("=SUM(E2:E3)"); // address in the formula is related to the whole worksheet
    // you could also use a relative formula to achieve the same result:
    // range.SetFormulaR1C1("=SUM(R[-2]C:R[-1]C)");

    // demonstrates another way of creating a range
    range = worksheet.GetRange("E2", "E4");
    // set the format of cells with formulas to currency
    wxExcelStyles styles = workbook.GetStyles();
    range.SetStyle(styles[wxString("Currency")]);

    // obtain a rectangular area containing all worksheet cells considered not empty
    range = worksheet.GetUsedRange();
    // add medium-weight borders on the outside and thin-weight on the inside
    wxExcelBorders borders = range.GetBorders();
    borders[xlEdgeTop].SetWeight(xlMedium);
    borders[xlEdgeLeft].SetWeight(xlMedium);
    borders[xlEdgeBottom].SetWeight(xlMedium);
    borders[xlEdgeRight].SetWeight(xlMedium);
    borders[xlInsideHorizontal].SetWeight(xlThin);
    borders[xlInsideVertical].SetWeight(xlThin);

    // format the totals row
    range = worksheet.GetRange("A4:E4");
    wxExcelFont font = range.GetFont();
    font.SetBold(true);
    font.SetColor(*wxBLUE);
    font.SetSize(font.GetSize() * 1.5); // 150% of default size

    // set the cell background to light grey
    range.GetInterior().SetColor(*wxLIGHT_GREY);

    // merge the first four cells
    range.GetRange("A1:D1").Merge();

    // get the range for cell with the total sum
    // using another method of specifying a range - row and column
    // WXAEEP is a helper macro for passing pointers to longs and Excel enums
    range = range.GetCells(NULL, WXAEEP(5L));
    // add a thick double-lined blue border around the total sum
    range.BorderAround(WXAEEP(xlDouble), WXAEEP(xlThick), NULL, wxBLUE);

    // finally, fit the columns to the content
    worksheet.GetUsedRange().GetEntireColumn().AutoFit();

    // show the text of the total sum as displayed in Excel
    wxLogMessage("Contents of cell E4: \"%s\"", worksheet.GetRange("E4").GetText());

    wxLogMessage("*** Shutting down wxAutoExcel...");

    return true;
}