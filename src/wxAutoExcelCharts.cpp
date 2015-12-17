/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Modified by:
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////


#include "wx/wxAutoExcel_prec.h"

#include "wx/wxAutoExcelCharts.h"

#if WXAUTOEXCEL_USE_CHARTS

#include "wx/wxAutoExcelChart.h"
#include "wx/wxAutoExcelPageBreaks.h"
#include "wx/wxAutoExcelSheet.h"

#include "wx/wxAutoExcel_private.h"

namespace wxAutoExcel {

// ***** class wxExcelCharts METHODS *****

wxExcelChart wxExcelCharts::Add()
{
    return DoAdd(NULL, false);
}

wxExcelChart wxExcelCharts::AddAfterOrBefore(wxExcelSheet sheetAfterOrBefore, bool after)
{
    return DoAdd(&sheetAfterOrBefore, after);
}

wxExcelChart wxExcelCharts::DoAdd(wxExcelSheet* sheetAfterOrBefore, bool after)
{
    wxExcelChart chart;
    wxVariant vAfterOrBefore;

    if (sheetAfterOrBefore != NULL) {
        if ( !ObjectToVariant(sheetAfterOrBefore, vAfterOrBefore) )
            return chart;
        vAfterOrBefore.SetName(after ? wxS("After") : wxS("Before"));
    }        
    
    WXAUTOEXCEL_CALL_METHOD1("Add", vAfterOrBefore, "void*", chart);   
    VariantToObject(vResult, &chart);    
    return chart;
}


// ***** class wxExcelCharts PROPERTIES *****


long wxExcelCharts::GetCount()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("Count");
}

wxExcelChart wxExcelCharts::GetItem(long index)
{
    wxExcelChart chart;
    WXAUTOEXCEL_PROPERTY_OBJECT_GET1("Item", index, chart);
}

wxExcelChart wxExcelCharts::operator[](long index)
{
    return GetItem(index);
}

} // namespace wxAutoExcel

#endif // #if WXAUTOEXCEL_USE_CHARTS
