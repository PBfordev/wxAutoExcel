/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Modified by:
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////


#include "wx/wxAutoExcel_prec.h"

#include "wx/wxAutoExcelLegend.h"

#if WXAUTOEXCEL_USE_CHARTS

#include "wx/wxAutoExcelChartFormat.h"
#include "wx/wxAutoExcelLegendEntries.h"

#include "wx/wxAutoExcelPrivate.h"

namespace wxAutoExcel {

// ***** class wxExcelLegend METHODS *****

bool wxExcelLegend::Clear()
{
    WXAUTOEXCEL_CALL_METHOD0_BOOL("Clear");
}

bool wxExcelLegend::Delete()
{
    WXAUTOEXCEL_CALL_METHOD0_BOOL("Delete");
}

wxExcelLegendEntries wxExcelLegend::LegendEntries()
{
    wxExcelLegendEntries entries;
    WXAUTOEXCEL_CALL_METHOD0_OBJECT("LegendEntries", entries);
}

bool wxExcelLegend::Select()
{
    WXAUTOEXCEL_CALL_METHOD0_BOOL("Select");
}

// ***** class wxExcelLegend PROPERTIES *****

bool wxExcelLegend::GetAutoScaleFont()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("AutoScaleFont");
}

void wxExcelLegend::SetAutoScaleFont(bool autoScaleFont)
{
    InvokePutProperty(wxS("AutoScaleFont"), autoScaleFont);
}

wxExcelChartFormat wxExcelLegend::GetFormat()
{
    wxExcelChartFormat chartFormat;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("Format", chartFormat);
}

double wxExcelLegend::GetHeight()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("Height");
}

void wxExcelLegend::SetHeight(double height)
{
    InvokePutProperty(wxS("Height"), height);
}

bool wxExcelLegend::GetIncludeInLayout()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("IncludeInLayout");
}

void wxExcelLegend::SetIncludeInLayout(bool includeInLayout)
{
    InvokePutProperty(wxS("IncludeInLayout"), includeInLayout);
}

double wxExcelLegend::GetLeft()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("Left");
}

void wxExcelLegend::SetLeft(double left)
{
    InvokePutProperty(wxS("Left"), left);
}

wxString wxExcelLegend::GetName()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("Name");
}


XlLegendPosition wxExcelLegend::GetPosition()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("Position", XlLegendPosition, xlLegendPositionCorner);
}

void wxExcelLegend::SetPosition(XlLegendPosition position)
{
    InvokePutProperty(wxS("Position"), (long)position);
}

bool wxExcelLegend::GetShadow()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("Shadow");
}

void wxExcelLegend::SetShadow(bool shadow)
{
    InvokePutProperty(wxS("Shadow"), shadow);
}

double wxExcelLegend::GetTop()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("Top");
}

void wxExcelLegend::SetTop(double top)
{
    InvokePutProperty(wxS("Top"), top);
}

double wxExcelLegend::GetWidth()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("Width");
}

void wxExcelLegend::SetWidth(double width)
{
    InvokePutProperty(wxS("Width"), width);
}

} // namespace wxAutoExcel


#endif // #if WXAUTOEXCEL_USE_CHARTS
