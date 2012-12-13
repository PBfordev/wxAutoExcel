/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Modified by:
// RCS-ID:      $Id: $
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////


#include "wx/wxAutoExcel_prec.h"

#include "wx/wxAutoExcelChartFormat.h"

#if WXAUTOEXCEL_USE_CHARTS

#include "wx/wxAutoExcelFillFormat.h"
#include "wx/wxAutoExcelGlowFormat.h"
#include "wx/wxAutoExcelLineFormat.h"
#include "wx/wxAutoExcelPictureFormat.h"
#include "wx/wxAutoExcelReflectionFormat.h"
#include "wx/wxAutoExcelShadowFormat.h"
#include "wx/wxAutoExcelSoftEdgeFormat.h"
#include "wx/wxAutoExcelThreeDFormat.h"
#include "wx/wxAutoExcelTextFrame.h"

#include "wx/wxAutoExcelPrivate.h"

namespace wxAutoExcel {

// ***** class wxExcelChartFormat PROPERTIES *****

wxExcelFillFormat wxExcelChartFormat::GetFill()
{
    wxExcelFillFormat fillFormat;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("Fill", fillFormat);
}

wxExcelGlowFormat wxExcelChartFormat::GetGlow()
{
    wxExcelGlowFormat glowFormat;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("Glow", glowFormat);
}

wxExcelLineFormat wxExcelChartFormat::GetLine()
{
    wxExcelLineFormat lineFormat;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("Line", lineFormat);
}

wxExcelPictureFormat wxExcelChartFormat::GetPictureFormat()
{
    wxExcelPictureFormat pictureFormat;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("PictureFormat", pictureFormat);
}

wxExcelReflectionFormat wxExcelChartFormat::GetReflection()
{
    wxExcelReflectionFormat reflectionFormat;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("Reflection", reflectionFormat);
}

wxExcelShadowFormat wxExcelChartFormat::GetShadow()
{
    wxExcelShadowFormat shadowFormat;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("Shadow", shadowFormat);
}

wxExcelSoftEdgeFormat wxExcelChartFormat::GetSoftEdge()
{
    wxExcelSoftEdgeFormat softEdgeFormat;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("SoftEdge", softEdgeFormat);
}

wxExcelTextFrame wxExcelChartFormat::GetTextFrame()
{
    wxExcelTextFrame textFrame;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("TextFrame", textFrame);
}

wxExcelThreeDFormat wxExcelChartFormat::GetThreeD()
{
    wxExcelThreeDFormat threeDFormat;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("ThreeD", threeDFormat);
}

} // namespace wxAutoExcel

#endif // #if WXAUTOEXCEL_USE_CHARTS
