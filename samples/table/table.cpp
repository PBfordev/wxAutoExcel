/////////////////////////////////////////////////////////////////////////////
// Purpose:     Demonstrates structured Excel table operations
// Author:      OpenAI Codex, under the direction of PB
// Generated:   AI-generated using GPT-5.6 Sol with High reasoning effort
// Copyright:   (c) 2026 PB <pbfordev@gmail.com>
// License:     MIT license
/////////////////////////////////////////////////////////////////////////////


/**********************************************************

wxAutoExcel Table sample shows how to:
- Create and style a structured Excel table (ListObject).
- Resize a table and add and remove table rows and columns.
- Add a calculated column and a totals row.
- Restrict input with data validation.
- Highlight values with conditional formatting.
- Sort and filter a table.

The sample requires Microsoft Excel 2007 or newer.

**********************************************************/


#include <wx/wx.h>
#include <wx/iconbndl.h>
#include <wx/msw/ole/oleutils.h>

#include <wx/wxAutoExcel.h>


#if !WXAUTOEXCEL_USE_CONDFORMAT
    #error In order to compile this sample, wxAutoExcel has to be built with WXAUTOEXCEL_USE_CONDFORMAT set to 1 in wxAutoExcel_setup.h
#endif


using namespace wxAutoExcel;


namespace
{

void SetRowValues(wxExcelRange range, const wxString& product,
                  const wxString& region, long units, double unitPrice)
{
    wxVariant values;
    values.ClearList();
    values.Append(product);
    values.Append(region);
    values.Append(units);
    values.Append(unitPrice);
    range.SetValue(values);
}

} // anonymous namespace


class TableSample
{
public:
    bool Run();

private:
    bool InitializeExcel();
    void WriteInitialData();
    bool CreateTable();
    bool ChangeTableSize();
    bool AddCalculatedColumn();
    void AddTotalsRow();
    bool AddValidation();
    bool AddConditionalFormatting();
    bool SortAndFilter();
    bool FormatWorksheet();
    void QuitExcelOnError();

    wxExcelApplication m_app;
    wxExcelWorkbook m_workbook;
    wxExcelWorksheet m_worksheet;
    wxExcelListObject m_table;
};


bool TableSample::Run()
{
    if ( !InitializeExcel() )
    {
        QuitExcelOnError();
        return false;
    }

    WriteInitialData();

    if ( !CreateTable() ||
         !ChangeTableSize() ||
         !AddCalculatedColumn() )
    {
        QuitExcelOnError();
        return false;
    }

    AddTotalsRow();

    if ( !AddValidation() ||
         !AddConditionalFormatting() ||
         !SortAndFilter() ||
         !FormatWorksheet() )
    {
        QuitExcelOnError();
        return false;
    }

    m_app.SetVisible(true);
    return true;
}


bool TableSample::InitializeExcel()
{
    m_app = wxExcelApplication::CreateInstance();
    if ( !m_app )
    {
        wxLogError(_("Failed to create an instance of MS Excel application."));
        return false;
    }

    if ( !m_app.IsVersionAtLeast_(wxExcelApplication::evExcel2007) )
    {
        wxLogError(_("This sample requires Microsoft Excel 2007 or newer."));
        return false;
    }

    m_workbook = m_app.GetWorkbooks().Add();
    if ( !m_workbook )
    {
        wxLogError(_("Failed to create a new workbook."));
        return false;
    }

    // Use English names for formulas, styles, and number formats in automation
    // calls regardless of the language into which Excel is localized.
    m_workbook.SetAutomationLCID_(wxExcelObject::lcidEnglishUS);

    m_worksheet = m_workbook.GetWorksheets()[1];
    if ( !m_worksheet )
    {
        wxLogError(_("Failed to obtain worksheet number 1."));
        return false;
    }

    m_worksheet.SetName("Sales Data");
    return true;
}


void TableSample::WriteInitialData()
{
    wxExcelRange title = m_worksheet.GetRange("A1:E1");
    title.Merge();
    title.SetValue("Sales Table");
    title.SetHorizontalAlignment(xlCenter);
    title.GetFont().SetBold(true);
    title.GetFont().SetSize(16);
    title.GetFont().SetColor(*wxWHITE);
    title.GetInterior().SetColor(wxColour(31, 78, 121));

    m_worksheet.GetRange("A2:E2").Merge();
    m_worksheet.GetRange("A2").SetValue(
        "Sorted by revenue and filtered to show rows with at least 10 units. "
        "Use the table filter buttons to explore all records.");

    wxVariant headers;
    headers.ClearList();
    headers.Append("Product");
    headers.Append("Region");
    headers.Append("Units");
    headers.Append("Unit Price");
    m_worksheet.GetRange("A4:D4").SetValue(headers);

    SetRowValues(m_worksheet.GetRange("A5:D5"), "Desk Chair", "North", 12, 89.95);
    SetRowValues(m_worksheet.GetRange("A6:D6"), "Monitor", "East", 8, 249.50);
    SetRowValues(m_worksheet.GetRange("A7:D7"), "Keyboard", "West", 25, 49.90);
    SetRowValues(m_worksheet.GetRange("A8:D8"), "Dock", "South", 15, 129.00);
    SetRowValues(m_worksheet.GetRange("A9:D9"), "Webcam", "North", 6, 79.00);
    SetRowValues(m_worksheet.GetRange("A10:D10"), "Headset", "East", 18, 59.50);

}


bool TableSample::CreateTable()
{
    wxExcelRange source = m_worksheet.GetRange("A4:D10");
    m_table = m_worksheet.GetListObjects().Add(
        &source, WXAEEP(xlYes), "TableStyleMedium2");

    if ( !m_table )
    {
        wxLogError(_("Failed to create the structured Excel table."));
        return false;
    }

    m_table.SetName("SalesTable");
    m_table.SetShowTableStyleRowStripes(true);
    return true;
}


bool TableSample::ChangeTableSize()
{
    // Resize adds another data row to the table without moving other cells.
    m_table.Resize(m_worksheet.GetRange("A4:D11"));
    wxExcelListRows rows = m_table.GetListRows();
    if ( !rows || rows.GetCount() != 7 )
    {
        wxLogError(_("Failed to resize the table."));
        return false;
    }

    SetRowValues(rows[rows.GetCount()].GetRange(),
                 "Desk Lamp", "West", 7, 39.90);

    // Add and then remove a temporary record to demonstrate ListRows methods.
    wxExcelListRow temporaryRow = rows.Add();
    if ( !temporaryRow )
    {
        wxLogError(_("Failed to add a table row."));
        return false;
    }
    SetRowValues(temporaryRow.GetRange(),
                 "Temporary record", "South", 1, 1.00);
    temporaryRow.Delete();
    return true;
}


bool TableSample::AddCalculatedColumn()
{
    wxExcelListColumns columns = m_table.GetListColumns();
    wxExcelListColumn revenueColumn = columns.Add();
    if ( !revenueColumn )
    {
        wxLogError(_("Failed to add the Revenue table column."));
        return false;
    }

    revenueColumn.SetName("Revenue");
    revenueColumn.GetDataBodyRange().SetFormula(
        "=[@Units]*[@[Unit Price]]");

    // Add and remove a temporary column to demonstrate ListColumns methods.
    wxExcelListColumn temporaryColumn = columns.Add();
    if ( !temporaryColumn )
    {
        wxLogError(_("Failed to add a table column."));
        return false;
    }
    temporaryColumn.SetName("Temporary");
    temporaryColumn.GetDataBodyRange().SetValue("Remove me");
    temporaryColumn.Delete();
    return true;
}


void TableSample::AddTotalsRow()
{
    m_table.SetShowTotals(true);

    wxExcelListColumns columns = m_table.GetListColumns();
    columns.GetItem(wxS("Revenue")).SetTotalsCalculation(
        xlTotalsCalculationSum);

    wxExcelRange totals = m_table.GetTotalsRowRange();
    totals.GetCells(NULL, WXAEEP(1L)).SetValue("TOTAL");
}


bool TableSample::AddValidation()
{
    // Write the validation choices only after the table has reached its final
    // width, so adding table columns cannot shift values in the same rows.
    m_worksheet.GetRange("H2").SetValue("North");
    m_worksheet.GetRange("H3").SetValue("South");
    m_worksheet.GetRange("H4").SetValue("East");
    m_worksheet.GetRange("H5").SetValue("West");

    wxExcelListColumns columns = m_table.GetListColumns();

    wxExcelRange regionCells =
        columns.GetItem(wxS("Region")).GetDataBodyRange();
    wxExcelValidation regionValidation = regionCells.GetValidation();
    if ( !regionValidation )
    {
        wxLogError(_("Failed to obtain data validation for the Region column."));
        return false;
    }

    regionValidation.Add(xlValidateList, WXAEEP(xlValidAlertStop), NULL,
                         "=$H$2:$H$5");
    regionValidation.SetIgnoreBlank(false);
    regionValidation.SetInCellDropdown(true);
    regionValidation.SetInputTitle("Sales region");
    regionValidation.SetInputMessage("Select a region from the list.");
    regionValidation.SetShowInput(true);
    regionValidation.SetErrorTitle("Unknown region");
    regionValidation.SetErrorMessage("Select North, South, East, or West.");
    regionValidation.SetShowError(true);

    wxExcelRange unitCells =
        columns.GetItem(wxS("Units")).GetDataBodyRange();
    wxExcelValidation unitValidation = unitCells.GetValidation();
    if ( !unitValidation )
    {
        wxLogError(_("Failed to obtain data validation for the Units column."));
        return false;
    }

    unitValidation.Add(xlValidateWholeNumber, WXAEEP(xlValidAlertStop),
                       WXAEEP(xlBetween), "0", "1000000");
    unitValidation.SetIgnoreBlank(false);
    unitValidation.SetInputTitle("Units");
    unitValidation.SetInputMessage(
        "Enter a whole number between 0 and 1,000,000.");
    unitValidation.SetShowInput(true);
    unitValidation.SetErrorTitle("Invalid quantity");
    unitValidation.SetErrorMessage(
        "Units must be a whole number between 0 and 1,000,000.");
    unitValidation.SetShowError(true);

    m_worksheet.GetColumns(8).SetHidden(true);
    return true;
}


bool TableSample::AddConditionalFormatting()
{
    wxExcelListColumns columns = m_table.GetListColumns();

    wxExcelRange unitCells =
        columns.GetItem(wxS("Units")).GetDataBodyRange();
    wxExcelFormatConditions unitConditions = unitCells.GetFormatConditions();
    if ( !unitConditions )
    {
        wxLogError(_("Failed to obtain conditional formatting for Units."));
        return false;
    }

    wxExcelFormatCondition lowStock = unitConditions.Add(
        xlCellValue, WXAEEP(xlLess), wxVariant(10L));
    if ( !lowStock )
    {
        wxLogError(_("Failed to add the low-units formatting rule."));
        return false;
    }
    lowStock.GetInterior().SetColor(wxColour(255, 199, 206));
    lowStock.GetFont().SetColor(wxColour(156, 0, 6));

    wxExcelRange revenueCells =
        columns.GetItem(wxS("Revenue")).GetDataBodyRange();
    wxExcelDatabar revenueBars =
        revenueCells.GetFormatConditions().AddDatabar();
    if ( !revenueBars )
    {
        wxLogError(_("Failed to add Revenue data bars."));
        return false;
    }
    revenueBars.GetBarColor().SetColor(wxColour(91, 155, 213));
    return true;
}


bool TableSample::SortAndFilter()
{
    wxExcelSort sort = m_table.GetSort();
    if ( !sort )
    {
        wxLogError(_("Failed to obtain the table Sort object."));
        return false;
    }

    wxExcelSortFields fields = sort.GetSortFields();
    if ( !fields )
    {
        wxLogError(_("Failed to obtain the table SortFields collection."));
        return false;
    }

    fields.Clear();
    wxExcelSortField field = fields.Add(
        m_table.GetListColumns().GetItem(wxS("Revenue")).GetDataBodyRange(),
        SortOnValues, WXAEEP(xlDescending));
    if ( !field )
    {
        wxLogError(_("Failed to add the Revenue sort field."));
        return false;
    }

    sort.SetHeader(xlYes);
    sort.SetMatchCase(false);
    sort.SetOrientation(xlSortColumns);
    sort.Apply();

    // Filter field numbers are relative to the table: Units is field 3.
    m_table.GetRange().AutoFilter(WXAELP(3L), ">=10");
    return true;
}


bool TableSample::FormatWorksheet()
{
    // The built-in Currency style uses the user's regional currency symbol,
    // separators, symbol position, and spacing. The English style name works
    // because the workbook's automation LCID is set to US English.
    wxExcelStyle currencyStyle =
        m_workbook.GetStyles().GetItem(wxS("Currency"));
    if ( !currencyStyle )
    {
        wxLogError(_("Failed to obtain the built-in Currency style."));
        return false;
    }

    wxExcelListColumns columns = m_table.GetListColumns();
    columns.GetItem(wxS("Unit Price")).GetDataBodyRange().
        SetStyle(currencyStyle);
    columns.GetItem(wxS("Revenue")).GetDataBodyRange().
        SetStyle(currencyStyle);
    m_table.GetTotalsRowRange().GetCells(NULL, WXAEEP(5L)).
        SetStyle(currencyStyle);

    m_worksheet.GetRange("A:E").GetEntireColumn().AutoFit();
    m_worksheet.GetRange("A1:E1").SetRowHeight(26);
    return true;
}


void TableSample::QuitExcelOnError()
{
    if ( m_app )
    {
        m_app.SetDisplayAlerts(false);
        m_app.Quit();
    }
}


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


MyFrame::MyFrame()
    : wxFrame(NULL, wxID_ANY, _("wxAutoExcel Table sample"))
{
    SetIcons(wxIconBundle("appIcon", NULL));

    wxMenu* menu = new wxMenu;
    menu->Append(wxID_NEW, _("&Show me!"));
    menu->Append(wxID_EXIT, _("E&xit"));

    wxMenuBar* menuBar = new wxMenuBar;
    menuBar->Append(menu, _("&Sample"));
    SetMenuBar(menuBar);

    Bind(wxEVT_COMMAND_MENU_SELECTED,
         &MyFrame::OnShowSample, this, wxID_NEW);
    Bind(wxEVT_COMMAND_MENU_SELECTED,
         &MyFrame::OnQuit, this, wxID_EXIT);
}


void MyFrame::OnShowSample(wxCommandEvent& WXUNUSED(event))
{
    TableSample sample;
    sample.Run();
}


void MyFrame::OnQuit(wxCommandEvent& WXUNUSED(event))
{
    Close(true);
}


bool MyApp::OnInit()
{
    if ( !wxApp::OnInit() )
        return false;

    MyFrame* frame = new MyFrame;
    frame->Show();

    wxLog::AddTraceMask(wxTRACE_AutoExcel);
    return true;
}


wxIMPLEMENT_APP(MyApp);
