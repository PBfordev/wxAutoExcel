/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// License:     MIT license
/////////////////////////////////////////////////////////////////////////////


#include "wx/wxAutoExcel_prec.h"

#include "wx/wxAutoExcelSparklineGroups.h"

#if WXAUTOEXCEL_USE_CHARTS

#include "wx/wxAutoExcelSparkAxes.h"
#include "wx/wxAutoExcelSparkline.h"
#include "wx/wxAutoExcelSparkPoints.h"

#include "wx/wxAutoExcelRange.h"
#include "wx/wxAutoExcelFormatColor.h"

#include "wx/wxAutoExcel_private.h"

namespace wxAutoExcel {

// ***** class wxExcelSparklineGroup METHODS *****

void wxExcelSparklineGroup::Delete()
{
    WXAUTOEXCEL_CALL_METHOD0_RET("Delete", "null");
}

void wxExcelSparklineGroup::Modify(wxExcelRange range, const wxString& sourceData)
{
    wxVariant vRange;
    if ( ObjectToVariant(&range, vRange, wxS("Range")) )
    {
        wxVariant vSourceData(sourceData, wxS("SourceData"));
        WXAUTOEXCEL_CALL_METHOD2_RET("Modify", vRange, vSourceData, "null");
    }
}

void wxExcelSparklineGroup::ModifyDateRange(const wxString& dateRange)
{
    WXAUTOEXCEL_CALL_METHOD1_RET("ModifyDateRange", dateRange, "null");
}

void wxExcelSparklineGroup::ModifyLocation(wxExcelRange location)
{
    wxVariant vRange;
    if ( ObjectToVariant(&location, vRange, wxS("Location")) )
    {
        WXAUTOEXCEL_CALL_METHOD1_RET("ModifyLocation", vRange, "null");
    }
}

void wxExcelSparklineGroup::ModifySourceData(const wxString& sourceData)
{
    WXAUTOEXCEL_CALL_METHOD1_RET("ModifySourceData", sourceData, "null");
}

// ***** class wxExcelSparklineGroup PROPERTIES *****

wxExcelSparkAxes wxExcelSparklineGroup::GetAxes()
{
    wxExcelSparkAxes sparkAxes;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("Axes", sparkAxes);
}

long wxExcelSparklineGroup::GetCount()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("Count");
}

wxString wxExcelSparklineGroup::GetDateRange()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("DateRange");
}

void wxExcelSparklineGroup::SetDateRange(const wxString& dateRange)
{
    InvokePutProperty(wxS("DateRange"), dateRange);
}

XlDisplayBlanksAs wxExcelSparklineGroup::GetDisplayBlanksAs()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("DisplayBlanksAs", XlDisplayBlanksAs, xlNotPlotted);
}

void wxExcelSparklineGroup::SetDisplayBlanksAs(XlDisplayBlanksAs displayBlanksAs)
{
    InvokePutProperty(wxS("DisplayBlanksAs"), (long)displayBlanksAs);
}

bool wxExcelSparklineGroup::GetDisplayHidden()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("DisplayHidden");
}

void wxExcelSparklineGroup::SetDisplayHidden(bool displayHidden)
{
    InvokePutProperty(wxS("DisplayHidden"), displayHidden);
}

wxExcelSparkline wxExcelSparklineGroup::GetItem(long index)
{
    wxASSERT( index > 0 );

    wxExcelSparkline sparkline;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET1("Item", index, sparkline);
}

wxExcelSparkline wxExcelSparklineGroup::operator[](long index)
{
    return GetItem(index);
}

double wxExcelSparklineGroup::GetLineWeight()
{
    wxVariant result(0.);
    InvokeGetProperty(wxS("LineWeight"), result);
    return result;
}

void wxExcelSparklineGroup::SetLineWeight(double lineWeight)
{
    InvokePutProperty(wxS("LineWeight"), lineWeight);
}

wxExcelRange wxExcelSparklineGroup::GetLocation()
{
    wxExcelRange range;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("Location", range);
}

void wxExcelSparklineGroup::SetLocation(const wxExcelRange& location)
{
    wxVariant vLocation;
    if ( ObjectToVariant(&location, vLocation, wxS("Location")) )
    {
        InvokePutProperty(wxS("Location"), vLocation);
    }
}

XlSparklineRowCol wxExcelSparklineGroup::GetPlotBy()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("PlotBy", XlSparklineRowCol, xlSparklineNonSquare);
}

void wxExcelSparklineGroup::SetPlotBy(XlSparklineRowCol plotBy)
{
    InvokePutProperty(wxS("PlotBy"), (long)plotBy);
}

wxExcelSparkPoints wxExcelSparklineGroup::GetPoints()
{
    wxExcelSparkPoints sparkPoints;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("Points", sparkPoints);
}

wxExcelFormatColor wxExcelSparklineGroup::GetSeriesColor()
{
    wxExcelFormatColor formatColor;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("SeriesColor", formatColor);
}

wxString wxExcelSparklineGroup::GetSourceData()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("SourceData");
}

void wxExcelSparklineGroup::SetSourceData(const wxString& sourceData)
{
    InvokePutProperty(wxS("SourceData"), sourceData);
}

XlSparkType wxExcelSparklineGroup::GetType()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("Type", XlSparkType, xlSparkLine);
}

void wxExcelSparklineGroup::SetType(XlSparkType type)
{
    InvokePutProperty(wxS("Type"), (long)type);
}

// ***** class wxExcelSparklineGroups METHODS *****

wxExcelSparklineGroup wxExcelSparklineGroups::Add(XlSparkType sparkType, const wxString& sourceData)
{
    wxExcelSparklineGroup group;

    WXAUTOEXCEL_CALL_METHOD2_OBJECT("Add", (long)sparkType, sourceData, group);
}

void wxExcelSparklineGroups::Clear()
{
    WXAUTOEXCEL_CALL_METHOD0_RET("Clear", "null");
}

void wxExcelSparklineGroups::ClearGroups()
{
    WXAUTOEXCEL_CALL_METHOD0_RET("ClearGroups", "null");
}

void wxExcelSparklineGroups::Group(wxExcelRange location)
{
    wxVariant vLocation;
    if ( ObjectToVariant(&location, vLocation, wxS("Location")) )
    {
        WXAUTOEXCEL_CALL_METHOD1_RET("Group", vLocation, "null");
    }
}

void wxExcelSparklineGroups::Ungroup()
{
    WXAUTOEXCEL_CALL_METHOD0_RET("Ungroup", "null");
}

// ***** class wxExcelSparklineGroups PROPERTIES *****

long wxExcelSparklineGroups::GetCount()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("Count");
}

wxExcelSparklineGroup wxExcelSparklineGroups::GetItem(long index)
{
    wxASSERT( index > 0 );

    wxExcelSparklineGroup sparklineGroup;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET1("Item", index, sparklineGroup);
}

wxExcelSparklineGroup wxExcelSparklineGroups::operator[](long index)
{
    return GetItem(index);
}

} // namespace wxAutoExcel

#endif // #if WXAUTOEXCEL_USE_CHARTS
