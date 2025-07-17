/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////


#include "wx/wxAutoExcel_prec.h"

#include "wx/wxAutoExcelSparkPoints.h"

#if WXAUTOEXCEL_USE_CHARTS

#include "wx/wxAutoExcelSparkColor.h"
#include "wx/wxAutoExcelSparklineGroups.h"

#include "wx/wxAutoExcel_private.h"

namespace wxAutoExcel {

// ***** class wxExcelSparkPoints PROPERTIES *****

wxExcelSparkColor wxExcelSparkPoints::GetFirstpoint()
{
    wxExcelSparkColor sparkColor;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("Firstpoint", sparkColor);
}

wxExcelSparkColor wxExcelSparkPoints::GetHighpoint()
{
    wxExcelSparkColor sparkColor;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("Highpoint", sparkColor);
}

wxExcelSparkColor wxExcelSparkPoints::GetLastpoint()
{
    wxExcelSparkColor sparkColor;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("Lastpoint", sparkColor);
}

wxExcelSparkColor wxExcelSparkPoints::GetLowpoint()
{
    wxExcelSparkColor sparkColor;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("Lowpoint", sparkColor);
}

wxExcelSparkColor wxExcelSparkPoints::GetMarkers()
{
    wxExcelSparkColor sparkColor;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("Markers", sparkColor);
}

wxExcelSparkColor wxExcelSparkPoints::GetNegative()
{
    wxExcelSparkColor sparkColor;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("Negative", sparkColor);
}

wxExcelSparklineGroup wxExcelSparkPoints::GetParent()
{
    wxExcelSparklineGroup sparklineGroup;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("Parent", sparklineGroup);
}




} // namespace wxAutoExcel

#endif // #if WXAUTOEXCEL_USE_CHARTS
