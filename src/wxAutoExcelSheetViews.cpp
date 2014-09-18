/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Modified by:
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////


#include "wx/wxAutoExcel_prec.h"

#include "wx/wxAutoExcelSheetViews.h"

#include "wx/wxAutoExcelPrivate.h"

#include "wx/wxAutoExcelChartView.h"
#include "wx/wxAutoExcelWorksheetView.h"

namespace wxAutoExcel {

// ***** wxExcelSheetView *****

bool wxExcelSheetView::IsView(View view) const
{
    if ( !IsOk_() )
        return false;

    wxString name(GetAutomationObjectName_(true));

    switch ( view )
    {
        case ChartView:
            return name.IsSameAs(wxS("ChartView"));
        case DialogSheetView:
            return name.IsSameAs(wxS("DialogSheetView"));
        case WorksheetView:
            return name.IsSameAs(wxS("WorksheetView"));
        default:
            wxFAIL_MSG(wxS("Invalid switch value"));    
    }   
    return false; // just to suppress compiler warning about not all control paths returning a value
}

wxExcelChartView wxExcelSheetView::ToChartView()
{
    wxExcelChartView chartView;

    if ( IsView(ChartView) )
        CloneDispatch(this, &chartView);
    return chartView;
}


wxExcelWorksheetView wxExcelSheetView::ToWorksheetView()
{
    wxExcelWorksheetView worksheetView;

    if ( IsView(WorksheetView) )
        CloneDispatch(this, &worksheetView);
    return worksheetView;
}

// ***** class wxExcelSheetViews PROPERTIES *****

long wxExcelSheetViews::GetCount()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("Count");
}

wxExcelSheetView wxExcelSheetViews::GetItem(long index)
{
    wxExcelSheetView sheetView;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET1("Item", index, sheetView);
}


} // namespace wxAutoExcel
