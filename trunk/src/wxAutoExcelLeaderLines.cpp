/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Modified by:
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////


#include "wx/wxAutoExcel_prec.h"

#include "wx/wxAutoExcelLeaderLines.h"

#if WXAUTOEXCEL_USE_CHARTS

#include "wx/wxAutoExcelBorders.h"
#include "wx/wxAutoExcelChartFormat.h"

#include "wx/wxAutoExcel_private.h"

namespace wxAutoExcel {

// ***** class wxExcelLeaderLines METHODS *****

bool wxExcelLeaderLines::Delete()
{
    WXAUTOEXCEL_CALL_METHOD0_BOOL("Delete");
}

bool wxExcelLeaderLines::Select()
{
    WXAUTOEXCEL_CALL_METHOD0_BOOL("Select");
}

// ***** class wxExcelLeaderLines PROPERTIES *****

wxExcelBorder wxExcelLeaderLines::GetBorder()
{
    wxExcelBorder border;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("Border", border);
}

wxExcelChartFormat wxExcelLeaderLines::GetFormat()
{
    wxExcelChartFormat chartFormat;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("Format", chartFormat);
}

} // namespace wxAutoExcel

#endif // #if WXAUTOEXCEL_USE_CHARTS
