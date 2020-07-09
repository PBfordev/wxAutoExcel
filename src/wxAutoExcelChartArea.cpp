/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Modified by:
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////


#include "wx/wxAutoExcel_prec.h"

#include "wx/wxAutoExcelChartArea.h"

#if WXAUTOEXCEL_USE_CHARTS

#include "wx/wxAutoExcelChartFormat.h"

#include "wx/wxAutoExcel_private.h"

namespace wxAutoExcel {

    // ***** class wxExcelChartArea METHODS *****

bool wxExcelChartArea::Clear()
{
    WXAUTOEXCEL_CALL_METHOD0_BOOL("Clear");
}

bool wxExcelChartArea::ClearContents()
{
    WXAUTOEXCEL_CALL_METHOD0_BOOL("ClearContents");
}

bool wxExcelChartArea::ClearFormats()
{
    WXAUTOEXCEL_CALL_METHOD0_BOOL("ClearFormats");
}

bool wxExcelChartArea::Copy()
{
    WXAUTOEXCEL_CALL_METHOD0_BOOL("Copy");
}

bool wxExcelChartArea::Select()
{
    WXAUTOEXCEL_CALL_METHOD0_BOOL("Select");
}


bool wxExcelChartArea::GetAutoScaleFont()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("AutoScaleFont");
}

void wxExcelChartArea::SetAutoScaleFont(bool autoScaleFont)
{
    InvokePutProperty(wxS("AutoScaleFont"), autoScaleFont);
}

wxExcelChartFormat wxExcelChartArea::GetFormat()
{
    wxExcelChartFormat chartFormat;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("Format", chartFormat);
}

double wxExcelChartArea::GetHeight()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("Height");
}

void wxExcelChartArea::SetHeight(double height)
{
    InvokePutProperty(wxS("Height"), height);
}

double wxExcelChartArea::GetLeft()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("Left");
}

void wxExcelChartArea::SetLeft(double left)
{
    InvokePutProperty(wxS("Left"), left);
}

wxString wxExcelChartArea::GetName()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("Name");
}

bool wxExcelChartArea::GetRoundedCorners()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("RoundedCorners");
}

void wxExcelChartArea::SetRoundedCorners(bool roundedCorners)
{
    InvokePutProperty(wxS("RoundedCorners"), roundedCorners);
}

bool wxExcelChartArea::GetShadow()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("Shadow");
}

void wxExcelChartArea::SetShadow(bool shadow)
{
    InvokePutProperty(wxS("Shadow"), shadow);
}

double wxExcelChartArea::GetTop()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("Top");
}

void wxExcelChartArea::SetTop(double top)
{
    InvokePutProperty(wxS("Top"), top);
}

double wxExcelChartArea::GetWidth()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("Width");
}

void wxExcelChartArea::SetWidth(double width)
{
    InvokePutProperty(wxS("Width"), width);
}

} // namespace wxAutoExcel

#endif // #if WXAUTOEXCEL_USE_CHARTS
