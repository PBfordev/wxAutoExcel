/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Modified by:
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////


#include "wx/wxAutoExcel_prec.h"

#include "wx/wxAutoExcelSparkAxes.h"

#if WXAUTOEXCEL_USE_CHARTS

#include "wx/wxAutoExcelSparkColor.h"
#include "wx/wxAutoExcelSparklineGroups.h"

#include "wx/wxAutoExcelPrivate.h"

namespace wxAutoExcel {

// ***** class wxExcelSparkHorizontalAxis PROPERTIES *****

wxExcelSparkColor wxExcelSparkHorizontalAxis::GetAxis()
{
    wxExcelSparkColor sparkColor;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("Axis", sparkColor);
}

bool wxExcelSparkHorizontalAxis::GetIsDateAxis()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("IsDateAxis");
}

wxExcelSparklineGroup wxExcelSparkHorizontalAxis::GetParent()
{
    wxExcelSparklineGroup sparklineGroup;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("Parent", sparklineGroup);
}

bool wxExcelSparkHorizontalAxis::GetRightToLeftPlotOrder()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("RightToLeftPlotOrder");
}


// ***** class wxExcelSparkVerticalAxis PROPERTIES *****

double wxExcelSparkVerticalAxis::GetCustomMaxScaleValue()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("CustomMaxScaleValue");
}

void wxExcelSparkVerticalAxis::SetCustomMaxScaleValue(double customMaxScaleValue)
{
    InvokePutProperty(wxS("CustomMaxScaleValue"), customMaxScaleValue);
}

double wxExcelSparkVerticalAxis::GetCustomMinScaleValue()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("CustomMinScaleValue");
}

void wxExcelSparkVerticalAxis::SetCustomMinScaleValue(double customMinScaleValue)
{
    InvokePutProperty(wxS("CustomMinScaleValue"), customMinScaleValue);
}

XlSparkScale wxExcelSparkVerticalAxis::GetMaxScaleType()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("MaxScaleType", XlSparkScale, xlSparkScaleGroup);
}

void wxExcelSparkVerticalAxis::SetMaxScaleType(XlSparkScale maxScaleType)
{
    InvokePutProperty(wxS("MaxScaleType"), (long)maxScaleType);
}

XlSparkScale wxExcelSparkVerticalAxis::GetMinScaleType()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("MinScaleType", XlSparkScale, xlSparkScaleGroup);
}

void wxExcelSparkVerticalAxis::SetMinScaleType(XlSparkScale minScaleType)
{
    InvokePutProperty(wxS("MinScaleType"), (long)minScaleType);
}

wxExcelSparklineGroup wxExcelSparkVerticalAxis::GetParent()
{
    wxExcelSparklineGroup sparklineGroup;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("Parent", sparklineGroup);
}


// ***** class wxExcelSparkAxes PROPERTIES *****

wxExcelSparkHorizontalAxis wxExcelSparkAxes::GetHorizontal()
{
    wxExcelSparkHorizontalAxis sparkHorizontalAxis;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("Horizontal", sparkHorizontalAxis);
}

wxExcelSparklineGroup wxExcelSparkAxes::GetParent()
{
    wxExcelSparklineGroup sparklineGroup;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("Parent", sparklineGroup);
}

wxExcelSparkVerticalAxis wxExcelSparkAxes::GetVertical()
{
    wxExcelSparkVerticalAxis sparkVerticalAxis;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("Vertical", sparkVerticalAxis);
}


} // namespace wxAutoExcel

#endif // #if WXAUTOEXCEL_USE_CHARTS
