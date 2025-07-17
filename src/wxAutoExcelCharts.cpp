/////////////////////////////////////////////////////////////////////////////
// Author:      PB
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

wxExcelChart wxExcelCharts::Add2(wxExcelSheet* before, wxExcelSheet* after,
                                 long* count, wxXlTribool newLayout)
{
    wxExcelChart chart;

    wxCHECK_MSG(before && after, chart, "before and after cannot be both specified");

    wxExcelSheet* sheetAfterOrBefore = NULL;
    wxVariant vAfterOrBefore;

    if ( before )
        sheetAfterOrBefore = before;
    else if ( after )
        sheetAfterOrBefore = after;

    if ( sheetAfterOrBefore ) 
    {
        if ( !ObjectToVariant(sheetAfterOrBefore, vAfterOrBefore) )
            return chart;
        vAfterOrBefore.SetName(before ? wxS("Before") : wxS("After"));
    }

    WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT_NAME(Count, count);
    WXAUTOEXCEL_OPTIONALCPPTBOOL_TO_OPTIONALVARIANT_NAME(NewLayout, newLayout);

    WXAUTOEXCEL_CALL_METHOD3("Add2", vAfterOrBefore, vCount, vNewLayout, "void*", chart);
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
