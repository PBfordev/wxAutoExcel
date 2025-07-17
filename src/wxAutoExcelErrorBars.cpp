/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////


#include "wx/wxAutoExcel_prec.h"

#include "wx/wxAutoExcelErrorBars.h"

#if WXAUTOEXCEL_USE_CHARTS

#include "wx/wxAutoExcelBorders.h"
#include "wx/wxAutoExcelChartFormat.h"

#include "wx/wxAutoExcel_private.h"

namespace wxAutoExcel {

// ***** class wxExcelErrorBars METHODS *****

bool wxExcelErrorBars::ClearFormats()
{
    WXAUTOEXCEL_CALL_METHOD0_BOOL("ClearFormats");
}

bool wxExcelErrorBars::Delete()
{
    WXAUTOEXCEL_CALL_METHOD0_BOOL("Delete");
}

bool wxExcelErrorBars::Select()
{
    WXAUTOEXCEL_CALL_METHOD0_BOOL("Select");
}

// ***** class wxExcelErrorBars PROPERTIES *****

wxExcelBorder wxExcelErrorBars::GetBorder()
{
    wxExcelBorder border;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("Border", border);
}

XlEndStyleCap wxExcelErrorBars::GetEndStyle()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("EndStyle", XlEndStyleCap, xlCap);
}

void wxExcelErrorBars::SetEndStyle(XlEndStyleCap endStyle)
{
    InvokePutProperty(wxS("EndStyle"), (long)endStyle);
}

wxExcelChartFormat wxExcelErrorBars::GetFormat()
{
    wxExcelChartFormat chartFormat;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("Format", chartFormat);
}

wxString wxExcelErrorBars::GetName()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("Name");
}

} // namespace wxAutoExcel

#endif // #if WXAUTOEXCEL_USE_CHARTS
