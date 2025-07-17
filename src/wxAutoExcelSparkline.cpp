/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// License:     MIT license
/////////////////////////////////////////////////////////////////////////////


#include "wx/wxAutoExcel_prec.h"

#include "wx/wxAutoExcelSparkline.h"

#if WXAUTOEXCEL_USE_CHARTS

#include "wx/wxAutoExcelRange.h"
#include "wx/wxAutoExcelSparklineGroups.h"

#include "wx/wxAutoExcel_private.h"

namespace wxAutoExcel {

// ***** class wxExcelSparkline METHODS *****

void wxExcelSparkline::ModifyLocation(wxExcelRange location)
{
    wxVariant vRange;
    if ( ObjectToVariant(&location, vRange, wxS("Location")) )
    {
        WXAUTOEXCEL_CALL_METHOD1_RET("ModifyLocation", vRange, "null");
    }
}

void wxExcelSparkline::ModifySourceData(const wxString& formula)
{
    WXAUTOEXCEL_CALL_METHOD1_RET("ModifySourceData", formula, "null");
}

// ***** class wxExcelSparkline PROPERTIES *****

wxExcelRange wxExcelSparkline::GetLocation()
{
    wxExcelRange range;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("Location", range);
}

void wxExcelSparkline::SetLocation(const wxExcelRange& location)
{
    wxVariant vLocation;
    if ( ObjectToVariant(&location, vLocation, wxS("Location")) )
    {
        InvokePutProperty(wxS("Location"), vLocation);
    }
}

wxExcelSparklineGroup wxExcelSparkline::GetParent()
{
    wxExcelSparklineGroup sparklineGroup;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("Parent", sparklineGroup);
}

wxString wxExcelSparkline::GetSourceData()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("SourceData");
}

void wxExcelSparkline::SetSourceData(const wxString& sourceData)
{
    InvokePutProperty(wxS("SourceData"), sourceData);
}


} // namespace wxAutoExcel

#endif // #if WXAUTOEXCEL_USE_CHARTS
