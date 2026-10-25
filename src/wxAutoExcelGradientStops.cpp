/////////////////////////////////////////////////////////////////////////////
// Author:      Claude Opus 5.5 (Anthropic), under the direction of PB
// Copyright:   (c) 2026 PB <pbfordev@gmail.com>
// License:     MIT license
/////////////////////////////////////////////////////////////////////////////


#include "wx/wxAutoExcel_prec.h"

#include "wx/wxAutoExcelGradientStops.h"

#if WXAUTOEXCEL_USE_SHAPES || WXAUTOEXCEL_USE_CHARTS

#include <wx/colour.h>

#include "wx/wxAutoExcelColorFormat.h"

#include "wx/wxAutoExcel_private.h"

namespace wxAutoExcel {

// ***** class wxExcelGradientStop PROPERTIES *****

wxExcelColorFormat wxExcelGradientStop::GetColor()
{
    wxExcelColorFormat colorFormat;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("Color", colorFormat);
}

double wxExcelGradientStop::GetPosition()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("Position");
}

void wxExcelGradientStop::SetPosition(double position)
{
    InvokePutProperty(wxS("Position"), position);
}

double wxExcelGradientStop::GetTransparency()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("Transparency");
}

void wxExcelGradientStop::SetTransparency(double transparency)
{
    InvokePutProperty(wxS("Transparency"), transparency);
}


// ***** class wxExcelGradientStops METHODS *****

void wxExcelGradientStops::Delete(long index)
{
    wxVariant vIndex(index, wxS("Index"));

    WXAUTOEXCEL_CALL_METHOD1_RET("Delete", vIndex, "null");
}

void wxExcelGradientStops::Insert(const wxColour& color, double position, double transparency, long index)
{
    wxVariant vRGB((long)color.GetRGB(), wxS("RGB"));
    wxVariant vPosition(position, wxS("Position"));
    wxVariant vTransparency(transparency, wxS("Transparency"));
    wxVariant vIndex(index, wxS("Index"));

    WXAUTOEXCEL_CALL_METHOD4_RET("Insert", vRGB, vPosition, vTransparency, vIndex, "null");
}

void wxExcelGradientStops::Insert2(const wxColour& color, double position, double transparency, long index,
                                   double brightness)
{
    wxVariant vRGB((long)color.GetRGB(), wxS("RGB"));
    wxVariant vPosition(position, wxS("Position"));
    wxVariant vTransparency(transparency, wxS("Transparency"));
    wxVariant vIndex(index, wxS("Index"));
    wxVariant vBrightness(brightness, wxS("Brightness"));

    WXAUTOEXCEL_CALL_METHOD5_RET("Insert2", vRGB, vPosition, vTransparency, vIndex, vBrightness, "null");
}


// ***** class wxExcelGradientStops PROPERTIES *****

long wxExcelGradientStops::GetCount()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("Count");
}

wxExcelGradientStop wxExcelGradientStops::GetItem(long index)
{
    wxASSERT( index > 0 );

    wxExcelGradientStop gradientStop;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET1("Item", index, gradientStop);
}

wxExcelGradientStop wxExcelGradientStops::operator[](long index)
{
    return GetItem(index);
}

} // namespace wxAutoExcel

#endif // #if WXAUTOEXCEL_USE_SHAPES || WXAUTOEXCEL_USE_CHARTS
