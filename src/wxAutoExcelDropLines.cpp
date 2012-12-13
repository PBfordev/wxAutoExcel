/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Modified by:
// RCS-ID:      $Id: $
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////


#include "wx/wxAutoExcel_prec.h"

#include "wx/wxAutoExcelDropLines.h"

#if WXAUTOEXCEL_USE_CHARTS

#include "wx/wxAutoExcelBorders.h"
#include "wx/wxAutoExcelChartFormat.h"

#include "wx/wxAutoExcelPrivate.h"

namespace wxAutoExcel {

// ***** class wxExcelDropLines METHODS *****

bool wxExcelDropLines::Delete()
{
    WXAUTOEXCEL_CALL_METHOD0_BOOL("Delete");
}

bool wxExcelDropLines::Select()
{
    WXAUTOEXCEL_CALL_METHOD0_BOOL("Select");
}

// ***** class wxExcelDropLines PROPERTIES *****


wxExcelBorder wxExcelDropLines::GetBorder()
{
    wxExcelBorder border;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("Border", border);
}

wxExcelChartFormat wxExcelDropLines::GetFormat()
{
    wxExcelChartFormat chartFormat;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("Format", chartFormat);
}

wxString wxExcelDropLines::GetName()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("Name");
}

} // namespace wxAutoExcel

#endif // #if WXAUTOEXCEL_USE_CHARTS
