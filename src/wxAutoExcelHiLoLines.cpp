/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// License:     MIT license
/////////////////////////////////////////////////////////////////////////////


#include "wx/wxAutoExcel_prec.h"

#include "wx/wxAutoExcelHiLoLines.h"

#if WXAUTOEXCEL_USE_CHARTS

#include "wx/wxAutoExcelBorders.h"
#include "wx/wxAutoExcelChartFormat.h"

#include "wx/wxAutoExcel_private.h"

namespace wxAutoExcel {

// ***** class wxExcelHiLoLines METHODS *****

bool wxExcelHiLoLines::Delete()
{
    WXAUTOEXCEL_CALL_METHOD0_BOOL("Delete");
}

bool wxExcelHiLoLines::Select()
{
    WXAUTOEXCEL_CALL_METHOD0_BOOL("Select");
}

// ***** class wxExcelHiLoLines PROPERTIES *****


wxExcelBorder wxExcelHiLoLines::GetBorder()
{
    wxExcelBorder border;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("Border", border);
}

wxExcelChartFormat wxExcelHiLoLines::GetFormat()
{
    wxExcelChartFormat chartFormat;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("Format", chartFormat);
}

wxString wxExcelHiLoLines::GetName()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("Name");
}

} // namespace wxAutoExcel

#endif // #if WXAUTOEXCEL_USE_CHARTS
