/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Modified by:
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////


#include "wx/wxAutoExcel_prec.h"

#include "wx/wxAutoExcelChartGroups.h"

#if WXAUTOEXCEL_USE_CHARTS

#include "wx/wxAutoExcelChartCategory.h"
#include "wx/wxAutoExcelDownBars.h"
#include "wx/wxAutoExcelDropLines.h"
#include "wx/wxAutoExcelHiLoLines.h"
#include "wx/wxAutoExcelSeries.h"
#include "wx/wxAutoExcelSeriesCollection.h"
#include "wx/wxAutoExcelSeriesLines.h"
#include "wx/wxAutoExcelTickLabels.h"
#include "wx/wxAutoExcelUpBars.h"

#include "wx/wxAutoExcel_private.h"

namespace wxAutoExcel {

// ***** class wxExcelChartGroup METHODS *****

wxExcelCategoryCollection wxExcelChartGroup::CategoryCollection()
{
    wxExcelCategoryCollection collection;
    WXAUTOEXCEL_CALL_METHOD0_OBJECT("CategoryCollection", collection);
}

wxExcelChartCategory wxExcelChartGroup::CategoryCollection(long index)
{
    wxExcelChartCategory ChartCategory;
    WXAUTOEXCEL_CALL_METHOD1_OBJECT("CategoryCollection", index, ChartCategory);
}

wxExcelChartCategory wxExcelChartGroup::CategoryCollection(const wxString& name)
{
    wxExcelChartCategory ChartCategory;
    WXAUTOEXCEL_CALL_METHOD1_OBJECT("CategoryCollection", name, ChartCategory);
}

wxExcelCategoryCollection wxExcelChartGroup::FullCategoryCollection()
{
    wxExcelCategoryCollection collection;
    WXAUTOEXCEL_CALL_METHOD0_OBJECT("FullCategoryCollection", collection);
}

wxExcelChartCategory wxExcelChartGroup::FullCategoryCollection(long index)
{
    wxExcelChartCategory ChartCategory;
    WXAUTOEXCEL_CALL_METHOD1_OBJECT("FullCategoryCollection", index, ChartCategory);
}

wxExcelChartCategory wxExcelChartGroup::FullCategoryCollection(const wxString& name)
{
    wxExcelChartCategory ChartCategory;
    WXAUTOEXCEL_CALL_METHOD1_OBJECT("FullCategoryCollection", name, ChartCategory);
}

wxExcelSeriesCollection wxExcelChartGroup::SeriesCollection()
{
    wxExcelSeriesCollection collection;
    WXAUTOEXCEL_CALL_METHOD0_OBJECT("SeriesCollection", collection);
}

wxExcelSeries wxExcelChartGroup::SeriesCollection(long index)
{
    wxExcelSeries series;
    WXAUTOEXCEL_CALL_METHOD1_OBJECT("SeriesCollection", index, series);
}

wxExcelSeries wxExcelChartGroup::SeriesCollection(const wxString& name)
{
    wxExcelSeries series;
    WXAUTOEXCEL_CALL_METHOD1_OBJECT("SeriesCollection", name, series);
}

// ***** class wxExcelChartGroup PROPERTIES *****


XlAxisGroup wxExcelChartGroup::GetAxisGroup()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("AxisGroup", XlAxisGroup, xlPrimary);
}

long wxExcelChartGroup::GetBinsCountValue()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("BinsCountValue");
}

void wxExcelChartGroup::SetBinsCountValue(long value)
{
    InvokePutProperty(wxS("BinsCountValue"), value);
}

bool wxExcelChartGroup::GetBinsOverflowEnabled()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("BinsOverflowEnabled");
}

void wxExcelChartGroup::SetBinsOverflowEnabled(bool enabled)
{
    InvokePutProperty(wxS("BinsOverflowEnabled"), enabled);
}

double wxExcelChartGroup::GetBinsOverflowValue()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("BinsOverflowValue");
}

void wxExcelChartGroup::SetBinsOverflowValue(double value)
{
    InvokePutProperty(wxS("BinsOverflowValue"), value);
}

XlBinsType wxExcelChartGroup::GetBinsType()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("BinsType", XlBinsType, xlBinsTypeAutomatic);
}

void wxExcelChartGroup::SetBinsType(XlBinsType type)
{
    InvokePutProperty(wxS("BinsType"), (long)type);
}

bool wxExcelChartGroup::GetBinsUnderflowEnabled()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("BinsUnderflowEnabled");
}

void wxExcelChartGroup::SetBinsUnderflowEnabled(bool enabled)
{
    InvokePutProperty(wxS("BinsUnderflowEnabled"), enabled);
}

double wxExcelChartGroup::GetBinsUnderflowValue()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("BinsUnderflowValue");
}

void wxExcelChartGroup::SetBinsUnderflowValue(double value)
{
    InvokePutProperty(wxS("BinsUnderflowValue"), value);
}

double wxExcelChartGroup::GetBinWidthValue()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("BinWidthValue");
}

void wxExcelChartGroup::SetBinWidthValue(double value)
{
    InvokePutProperty(wxS("BinWidthValue"), value);
}

long wxExcelChartGroup::GetBubbleScale()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("BubbleScale");
}

void wxExcelChartGroup::SetBubbleScale(long bubbleScale)
{
    InvokePutProperty(wxS("BubbleScale"), bubbleScale);
}

long wxExcelChartGroup::GetDoughnutHoleSize()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("DoughnutHoleSize");
}

wxExcelDownBars wxExcelChartGroup::GetDownBars()
{
    wxExcelDownBars downBars;
    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("DownBars", downBars);
}

wxExcelDropLines wxExcelChartGroup::GetDropLines()
{
    wxExcelDropLines dropLines;
    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("DropLines", dropLines);
}

long wxExcelChartGroup::GetFirstSliceAngle()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("FirstSliceAngle");
}

void wxExcelChartGroup::SetFirstSliceAngle(long firstSliceAngle)
{
    InvokePutProperty(wxS("FirstSliceAngle"), firstSliceAngle);
}

long wxExcelChartGroup::GetGapWidth()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("GapWidth");
}

void wxExcelChartGroup::SetGapWidth(long gapWidth)
{
    InvokePutProperty(wxS("GapWidth"), gapWidth);
}

bool wxExcelChartGroup::GetHas3DShading()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("Has3DShading");
}

void wxExcelChartGroup::SetHas3DShading(bool has3DShading)
{
    InvokePutProperty(wxS("Has3DShading"), has3DShading);
}

bool wxExcelChartGroup::GetHasDropLines()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("HasDropLines");
}

void wxExcelChartGroup::SetHasDropLines(bool hasDropLines)
{
    InvokePutProperty(wxS("HasDropLines"), hasDropLines);
}

bool wxExcelChartGroup::GetHasHiLoLines()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("HasHiLoLines");
}

void wxExcelChartGroup::SetHasHiLoLines(bool hasHiLoLines)
{
    InvokePutProperty(wxS("HasHiLoLines"), hasHiLoLines);
}

bool wxExcelChartGroup::GetHasRadarAxisLabels()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("HasRadarAxisLabels");
}

void wxExcelChartGroup::SetHasRadarAxisLabels(bool hasRadarAxisLabels)
{
    InvokePutProperty(wxS("HasRadarAxisLabels"), hasRadarAxisLabels);
}

bool wxExcelChartGroup::GetHasSeriesLines()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("HasSeriesLines");
}

void wxExcelChartGroup::SetHasSeriesLines(bool hasSeriesLines)
{
    InvokePutProperty(wxS("HasSeriesLines"), hasSeriesLines);
}

bool wxExcelChartGroup::GetHasUpDownBars()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("HasUpDownBars");
}

void wxExcelChartGroup::SetHasUpDownBars(bool hasUpDownBars)
{
    InvokePutProperty(wxS("HasUpDownBars"), hasUpDownBars);
}

wxExcelHiLoLines wxExcelChartGroup::GetHiLoLines()
{
    wxExcelHiLoLines hiLoLines;
    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("HiLoLines", hiLoLines);
}

long wxExcelChartGroup::GetIndex()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("Index");
}

long wxExcelChartGroup::GetOverlap()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("Overlap");
}

void wxExcelChartGroup::SetOverlap(long overlap)
{
    InvokePutProperty(wxS("Overlap"), overlap);
}


wxExcelTickLabels wxExcelChartGroup::GetRadarAxisLabels()
{
    wxExcelTickLabels tickLabels;
    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("RadarAxisLabels", tickLabels);
}

long wxExcelChartGroup::GetSecondPlotSize()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("SecondPlotSize");
}

void wxExcelChartGroup::SetSecondPlotSize(long secondPlotSize)
{
    InvokePutProperty(wxS("SecondPlotSize"), secondPlotSize);
}

wxExcelSeriesLines wxExcelChartGroup::GetSeriesLines()
{
    wxExcelSeriesLines seriesLines;
    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("SeriesLines", seriesLines);
}

bool wxExcelChartGroup::GetShowNegativeBubbles()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("ShowNegativeBubbles");
}

void wxExcelChartGroup::SetShowNegativeBubbles(bool showNegativeBubbles)
{
    InvokePutProperty(wxS("ShowNegativeBubbles"), showNegativeBubbles);
}

XlSizeRepresents wxExcelChartGroup::GetSizeRepresents()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("SizeRepresents", XlSizeRepresents, xlSizeIsArea);
}

void wxExcelChartGroup::SetSizeRepresents(XlSizeRepresents sizeRepresents)
{
    InvokePutProperty(wxS("SizeRepresents"), (long)sizeRepresents);
}

XlChartSplitType wxExcelChartGroup::GetSplitType()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("SplitType", XlChartSplitType, xlSplitByPosition);
}

void wxExcelChartGroup::SetSplitType(XlChartSplitType splitType)
{
    InvokePutProperty(wxS("SplitType"), (long)splitType);
}

double wxExcelChartGroup::GetSplitValue()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("SplitValue");
}

void wxExcelChartGroup::SetSplitValue(double splitValue)
{
    InvokePutProperty(wxS("SplitValue"), splitValue);
}

wxExcelUpBars wxExcelChartGroup::GetUpBars()
{
    wxExcelUpBars upBars;
    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("UpBars", upBars);
}

bool wxExcelChartGroup::GetVaryByCategories()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("VaryByCategories");
}

void wxExcelChartGroup::SetVaryByCategories(bool varyByCategories)
{
    InvokePutProperty(wxS("VaryByCategories"), varyByCategories);
}

// ***** class wxExcelChartGroups METHODS *****

wxExcelChartGroup wxExcelChartGroups::Item(long index)
{
    wxExcelChartGroup group;
    WXAUTOEXCEL_CALL_METHOD1_OBJECT("Item", index, group);
}

wxExcelChartGroup wxExcelChartGroups::operator[](long index)
{
    return Item(index);
}

// ***** class wxExcelChartGroups PROPERTIES *****

long wxExcelChartGroups::GetCount()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("Count");
}

} // namespace wxAutoExcel

#endif // #if WXAUTOEXCEL_USE_CHARTS
