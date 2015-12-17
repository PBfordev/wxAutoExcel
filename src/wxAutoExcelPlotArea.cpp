/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Modified by:
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////


#include "wx/wxAutoExcel_prec.h"

#include "wx/wxAutoExcelPlotArea.h"

#if WXAUTOEXCEL_USE_CHARTS

#include "wx/wxAutoExcelChartFormat.h"

#include "wx/wxAutoExcel_private.h"

namespace wxAutoExcel {

// ***** class wxExcelPlotArea METHODS *****

bool wxExcelPlotArea::ClearFormats()
{
    WXAUTOEXCEL_CALL_METHOD0_BOOL("ClearFormats");
}


bool wxExcelPlotArea::Select()
{
    WXAUTOEXCEL_CALL_METHOD0_BOOL("Select");
}


// ***** class wxExcelPlotArea PROPERTIES *****


wxExcelChartFormat wxExcelPlotArea::GetFormat()
{
    wxExcelChartFormat chartFormat;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("Format", chartFormat);
}

double wxExcelPlotArea::GetHeight()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("Height");
}

void wxExcelPlotArea::SetHeight(double height)
{
    InvokePutProperty(wxS("Height"), height);
}

double wxExcelPlotArea::GetInsideHeight()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("InsideHeight");
}

double wxExcelPlotArea::GetInsideLeft()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("InsideLeft");
}

double wxExcelPlotArea::GetInsideTop()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("InsideTop");
}

double wxExcelPlotArea::GetInsideWidth()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("InsideWidth");
}

double wxExcelPlotArea::GetLeft()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("Left");
}

void wxExcelPlotArea::SetLeft(double left)
{
    InvokePutProperty(wxS("Left"), left);
}

wxString wxExcelPlotArea::GetName()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("Name");
}

XlChartElementPosition wxExcelPlotArea::GetPosition()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("Position", XlChartElementPosition, xlChartElementPositionAutomatic);
}

void wxExcelPlotArea::SetPosition(XlChartElementPosition position)
{
    InvokePutProperty(wxS("Position"), (long)position);
}

double wxExcelPlotArea::GetTop()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("Top");
}

void wxExcelPlotArea::SetTop(double top)
{
    InvokePutProperty(wxS("Top"), top);
}

double wxExcelPlotArea::GetWidth()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("Width");
}

void wxExcelPlotArea::SetWidth(double width)
{
    InvokePutProperty(wxS("Width"), width);
}

} // namespace wxAutoExcel


#endif // #if WXAUTOEXCEL_USE_CHARTS
