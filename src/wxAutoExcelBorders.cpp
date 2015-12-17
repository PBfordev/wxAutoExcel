/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Modified by:
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////


#include "wx/wxAutoExcel_prec.h"

#include "wx/wxAutoExcelBorders.h"

#include "wx/wxAutoExcel_private.h"

namespace wxAutoExcel {


wxColour wxExcelBorder::GetColor()
{
    WXAUTOEXCEL_PROPERTY_COLOR_GET0("Color");
}

void wxExcelBorder::SetColor(const wxColour& color)
{
    InvokePutProperty(wxS("Color"), (long)color.GetRGB());
}

long wxExcelBorder::GetColorIndex()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("ColorIndex");
}

void wxExcelBorder::SetColorIndex(long colorIndex)
{
    InvokePutProperty(wxS("ColorIndex"), colorIndex);
}

long wxExcelBorder::GetLineStyle()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("LineStyle");
}

void wxExcelBorder::SetLineStyle(long lineStyle)
{
    InvokePutProperty(wxS("LineStyle"), lineStyle);
}


XlThemeColor wxExcelBorder::GetThemeColor()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("ThemeColor", XlThemeColor, xlThemeColorLight1);
}

void wxExcelBorder::SetThemeColor(XlThemeColor themeColor)
{
    InvokePutProperty(wxS("ThemeColor"), (long)themeColor);
}

double wxExcelBorder::GetTintAndShade()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("TintAndShade");
}

void wxExcelBorder::SetTintAndShade(double tintAndShade)
{
    InvokePutProperty(wxS("TintAndShade"), tintAndShade);
}

XlBorderWeight wxExcelBorder::GetWeight()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("Weight", XlBorderWeight, (XlBorderWeight)0);
}

void wxExcelBorder::SetWeight(XlBorderWeight weight)
{
    InvokePutProperty(wxS("Weight"), (long)weight);
}


long wxExcelBorders::GetCount()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("Count");
}


wxExcelBorder wxExcelBorders::GetItem(XlBordersIndex index)
{
    wxExcelBorder border;
    WXAUTOEXCEL_PROPERTY_OBJECT_GET1("Item", (long)index, border);
}

wxExcelBorder wxExcelBorders::operator[](XlBordersIndex index)
{
    return GetItem(index);
}


} // namespace wxAutoExcel
