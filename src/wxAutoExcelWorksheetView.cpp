/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Modified by:
// RCS-ID:      $Id: $
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////


#include "wx/wxAutoExcel_prec.h"

#include "wx/wxAutoExcelWorksheetView.h"

#include "wx/wxAutoExcelPrivate.h"

namespace wxAutoExcel {


// ***** class wxExcelWorksheetView PROPERTIES *****


bool wxExcelWorksheetView::GetDisplayFormulas()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("DisplayFormulas");
}

void wxExcelWorksheetView::SetDisplayFormulas(bool displayFormulas)
{
    InvokePutProperty(wxS("DisplayFormulas"), displayFormulas);
}

bool wxExcelWorksheetView::GetDisplayGridlines()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("DisplayGridlines");
}

void wxExcelWorksheetView::SetDisplayGridlines(bool displayGridlines)
{
    InvokePutProperty(wxS("DisplayGridlines"), displayGridlines);
}

bool wxExcelWorksheetView::GetDisplayHeadings()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("DisplayHeadings");
}

void wxExcelWorksheetView::SetDisplayHeadings(bool displayHeadings)
{
    InvokePutProperty(wxS("DisplayHeadings"), displayHeadings);
}

bool wxExcelWorksheetView::GetDisplayOutline()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("DisplayOutline");
}

void wxExcelWorksheetView::SetDisplayOutline(bool displayOutline)
{
    InvokePutProperty(wxS("DisplayOutline"), displayOutline);
}

bool wxExcelWorksheetView::GetDisplayZeros()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("DisplayZeros");
}

void wxExcelWorksheetView::SetDisplayZeros(bool displayZeros)
{
    InvokePutProperty(wxS("DisplayZeros"), displayZeros);
}


wxString wxExcelWorksheetView::GetSheet()
{    
    WXAUTOEXCEL_PROPERTY_STRING_GET0("Sheet");
}



} // namespace wxAutoExcel
