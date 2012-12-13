/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Modified by:
// RCS-ID:      $Id: $
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////


#include "wx/wxAutoExcel_prec.h"

#include "wx/wxAutoExcelChartColorFormat.h"

#if WXAUTOEXCEL_USE_CHARTS

#include "wx/wxAutoExcelPrivate.h"

namespace wxAutoExcel {

// ***** class wxExcelChartColorFormat PROPERTIES *****

wxColour wxExcelChartColorFormat::GetRGB()
{
    WXAUTOEXCEL_PROPERTY_COLOR_GET0("RGB");
}

long wxExcelChartColorFormat::GetSchemeColor()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("SchemeColor");
}

MsoColorType wxExcelChartColorFormat::GetType()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("Type", MsoColorType, msoColorTypeRGB);
}

} // namespace wxAutoExcel

#endif // #if WXAUTOEXCEL_USE_CHARTS
