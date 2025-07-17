/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// License:     MIT license
/////////////////////////////////////////////////////////////////////////////


/**********************************************************

wxAutoExcel Minimal sample shows how to:
- Create an MS Excel instance.
- Add a new empty workbook.
- Obtain the first worksheet from the added workbook.
- Set a worksheet name (displayed in its tab).
- Set US English LCID for a wxAutoExcel object
  so you can use English names/formats for its properties
  (e.g. Range.Address, Range.NumberFormat, Range.Formula...)
  regardless of the language MS Excel may be localized into.
- Create a range using various methods.
- Set cell values for small ranges
  (see Bulkdata sample to see how to do it efficiently for large data sets).
- Set cell alignment, font, style, borders and background color.
- Autofit columns.
- Get a formatted value as a text from a range.

**********************************************************/


#include <wx/wx.h>
#include <wx/log.h>
#include <wx/msw/ole/oleutils.h>
#include <wx/iconbndl.h>

#include <wx/wxAutoExcel.h>


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
: wxFrame(NULL, wxID_ANY, _("wxAutoExcel minimal sample"))
{
    SetIcons(wxIconBundle("appIcon", NULL));

    wxMenu *menu = new wxMenu;
    menu->Append(wxID_NEW, _("&Show me!"));
    menu->Append(wxID_EXIT, _("E&xit"));

    wxMenuBar *menuBar = new wxMenuBar();
    menuBar->Append(menu, _("&Sample"));
    SetMenuBar(menuBar);

    Bind(wxEVT_COMMAND_MENU_SELECTED, &MyFrame::OnCreateWorksheet, this, wxID_NEW);
    Bind(wxEVT_COMMAND_MENU_SELECTED, &MyFrame::OnQuit, this, wxID_EXIT);
}


wxVariant DoubleToCurrencyVariant(double d, bool* success = NULL)
{
    CURRENCY cy = {0};

    HRESULT hr = VarCyFromR8(d, &cy);
    if ( success )
        *success = SUCCEEDED(hr);

    return wxVariant(new wxVariantDataCurrency(cy));
}

void MyFrame::OnCreateWorksheet(wxCommandEvent& WXUNUSED(event))
{
    // first create an MS Excel instance
    wxExcelApplication app = wxExcelApplication::CreateInstance();
    if ( !app )
    {
        wxLogError(_("Failed to create an instance of MS Excel application."));
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
        return;
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
    wxMessageBox(worksheet.GetRange("E4").GetText(), "Contents of cell E4");
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

    // display wxAutoExcel-related traces in debug output
    wxLog::AddTraceMask(wxTRACE_AutoExcel);

    return true;
}

wxIMPLEMENT_APP(MyApp);