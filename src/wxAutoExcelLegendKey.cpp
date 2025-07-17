/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// License:     MIT license
/////////////////////////////////////////////////////////////////////////////


#include "wx/wxAutoExcel_prec.h"

#include "wx/wxAutoExcelLegendKey.h"

#if WXAUTOEXCEL_USE_CHARTS

#include "wx/wxAutoExcelChartFormat.h"

#include "wx/wxAutoExcel_private.h"

namespace wxAutoExcel {

// ***** class wxExcelLegendKey METHODS *****

bool wxExcelLegendKey::ClearFormats()
{
    WXAUTOEXCEL_CALL_METHOD0_BOOL("ClearFormats");
}

bool wxExcelLegendKey::Delete()
{
    WXAUTOEXCEL_CALL_METHOD0_BOOL("Delete");
}

// ***** class wxExcelLegendKey PROPERTIES *****

wxExcelChartFormat wxExcelLegendKey::GetFormat()
{
    wxExcelChartFormat chartFormat;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("Format", chartFormat);
}

double wxExcelLegendKey::GetHeight()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("Height");
}

bool wxExcelLegendKey::GetInvertIfNegative()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("InvertIfNegative");
}

void wxExcelLegendKey::SetInvertIfNegative(bool invertIfNegative)
{
    InvokePutProperty(wxS("InvertIfNegative"), invertIfNegative);
}

double wxExcelLegendKey::GetLeft()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("Left");
}

wxColour wxExcelLegendKey::GetMarkerBackgroundColor()
{
    WXAUTOEXCEL_PROPERTY_COLOR_GET0("MarkerBackgroundColor");
}

void wxExcelLegendKey::SetMarkerBackgroundColor(const wxColour& markerBackgroundColor)
{
    InvokePutProperty(wxS("MarkerBackgroundColor"), (long)markerBackgroundColor.GetRGB());
}

long wxExcelLegendKey::GetMarkerBackgroundColorIndex()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("MarkerBackgroundColorIndex");
}

void wxExcelLegendKey::SetMarkerBackgroundColorIndex(long markerBackgroundColorIndex)
{
    InvokePutProperty(wxS("MarkerBackgroundColorIndex"), markerBackgroundColorIndex);
}

wxColour wxExcelLegendKey::GetMarkerForegroundColor()
{
    WXAUTOEXCEL_PROPERTY_COLOR_GET0("MarkerForegroundColor");
}

void wxExcelLegendKey::SetMarkerForegroundColor(const wxColour& markerForegroundColor)
{
    InvokePutProperty(wxS("MarkerForegroundColor"), (long)markerForegroundColor.GetRGB());
}

long wxExcelLegendKey::GetMarkerForegroundColorIndex()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("MarkerForegroundColorIndex");
}

void wxExcelLegendKey::SetMarkerForegroundColorIndex(long markerForegroundColorIndex)
{
    InvokePutProperty(wxS("MarkerForegroundColorIndex"), markerForegroundColorIndex);
}

long wxExcelLegendKey::GetMarkerSize()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("MarkerSize");
}

void wxExcelLegendKey::SetMarkerSize(long markerSize)
{
    InvokePutProperty(wxS("MarkerSize"), markerSize);
}

XlMarkerStyle wxExcelLegendKey::GetMarkerStyle()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("MarkerStyle", XlMarkerStyle, xlMarkerStyleSquare);
}

void wxExcelLegendKey::SetMarkerStyle(XlMarkerStyle markerStyle)
{
    InvokePutProperty(wxS("MarkerStyle"), (long)markerStyle);
}


XlChartPictureType wxExcelLegendKey::GetPictureType()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("PictureType", XlChartPictureType, xlStretch);
}

void wxExcelLegendKey::SetPictureType(XlChartPictureType pictureType)
{
    InvokePutProperty(wxS("PictureType"), (long)pictureType);
}

long wxExcelLegendKey::GetPictureUnit()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("PictureUnit");
}

void wxExcelLegendKey::SetPictureUnit(long pictureUnit)
{
    InvokePutProperty(wxS("PictureUnit"), pictureUnit);
}

double wxExcelLegendKey::GetPictureUnit2()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("PictureUnit2");
}

void wxExcelLegendKey::SetPictureUnit2(double pictureUnit2)
{
    InvokePutProperty(wxS("PictureUnit2"), pictureUnit2);
}

bool wxExcelLegendKey::GetShadow()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("Shadow");
}

void wxExcelLegendKey::SetShadow(bool shadow)
{
    InvokePutProperty(wxS("Shadow"), shadow);
}

bool wxExcelLegendKey::GetSmooth()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("Smooth");
}

void wxExcelLegendKey::SetSmooth(bool smooth)
{
    InvokePutProperty(wxS("Smooth"), smooth);
}

double wxExcelLegendKey::GetTop()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("Top");
}

double wxExcelLegendKey::GetWidth()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("Width");
}

} // namespace wxAutoExcel


#endif // #if WXAUTOEXCEL_USE_CHARTS
