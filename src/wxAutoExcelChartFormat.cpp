/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// License:     MIT license
/////////////////////////////////////////////////////////////////////////////


#include "wx/wxAutoExcel_prec.h"

#include "wx/wxAutoExcelChartFormat.h"

#if WXAUTOEXCEL_USE_CHARTS

#include "wx/wxAutoExcelAdjustments.h"
#include "wx/wxAutoExcelFillFormat.h"
#include "wx/wxAutoExcelGlowFormat.h"
#include "wx/wxAutoExcelLineFormat.h"
#include "wx/wxAutoExcelPictureFormat.h"
#include "wx/wxAutoExcelShadowFormat.h"
#include "wx/wxAutoExcelSoftEdgeFormat.h"
#include "wx/wxAutoExcelThreeDFormat.h"
#include "wx/wxAutoExcelTextFrame2.h"

#include "wx/wxAutoExcel_private.h"

namespace wxAutoExcel {

// ***** class wxExcelChartFormat PROPERTIES *****

wxExcelAdjustments wxExcelChartFormat::GetAdjustments()
{
    wxExcelAdjustments adjustments;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("Adjustments", adjustments);
}

MsoAutoShapeType wxExcelChartFormat::GetAutoShapeType()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("AutoShapeType", MsoAutoShapeType, msoShapeRectangle);
}

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

wxExcelTextFrame2 wxExcelChartFormat::GetTextFrame2()
{
    wxExcelTextFrame2 textFrame;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("TextFrame2", textFrame);
}

wxExcelThreeDFormat wxExcelChartFormat::GetThreeD()
{
    wxExcelThreeDFormat threeDFormat;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("ThreeD", threeDFormat);
}

} // namespace wxAutoExcel

#endif // #if WXAUTOEXCEL_USE_CHARTS
