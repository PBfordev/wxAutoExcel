/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// License:     MIT license
/////////////////////////////////////////////////////////////////////////////


#include "wx/wxAutoExcel_prec.h"

#include "wx/wxAutoExcelGradient.h"

#include "wx/wxAutoExcel_private.h"

namespace wxAutoExcel {

// ***** class wxExcelColorStop METHODS *****

bool wxExcelColorStop::Delete()
{
    WXAUTOEXCEL_CALL_METHOD0_BOOL("Delete");
}

// ***** class wxExcelColorStop PROPERTIES *****

wxColour wxExcelColorStop::GetColor()
{
    WXAUTOEXCEL_PROPERTY_COLOR_GET0("Color");
}

void wxExcelColorStop::SetColor(const wxColour& color)
{
    InvokePutProperty(wxS("Color"), (long)color.GetRGB());
}

double wxExcelColorStop::GetPosition()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("Position");
}

void wxExcelColorStop::SetPosition(double position)
{
    InvokePutProperty(wxS("Position"), position);
}

XlThemeColor wxExcelColorStop::GetThemeColor()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("ThemeColor", XlThemeColor, xlThemeColorLight1);
}

void wxExcelColorStop::SetThemeColor(XlThemeColor themeColor)
{
    InvokePutProperty(wxS("ThemeColor"), (long)themeColor);
}

double wxExcelColorStop::GetTintAndShade()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("TintAndShade");
}

void wxExcelColorStop::SetTintAndShade(double tintAndShade)
{
    InvokePutProperty(wxS("TintAndShade"), tintAndShade);
}


// ***** class wxExcelColorStops METHODS *****

wxExcelColorStop wxExcelColorStops::Add(double position)
{
    wxExcelColorStop colorStop;

    WXAUTOEXCEL_CALL_METHOD1("Add", position, "void*", colorStop);
    VariantToObject(vResult, &colorStop);
    return colorStop;
}

void wxExcelColorStops::Clear()
{
    WXAUTOEXCEL_CALL_METHOD0_RET("Clear", "null");
}


// ***** class wxExcelColorStops PROPERTIES *****

// This is actually declared as a method?!
wxExcelColorStop wxExcelColorStops::GetItem(long index)
{
    wxExcelColorStop colorStop;

    WXAUTOEXCEL_CALL_METHOD1("Item", index, "void*", colorStop);
    VariantToObject(vResult, &colorStop);
    return colorStop;
}


long wxExcelColorStops::GetCount()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("Count");
}



// ***** class wxExcelLinearGradient PROPERTIES *****

wxExcelColorStops wxExcelLinearGradient::GetColorStops()
{
    wxExcelColorStops colorStops;
    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("ColorStops", colorStops);
}

long wxExcelLinearGradient::GetDegree()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("Degree");
}

void wxExcelLinearGradient::SetDegree(long degree)
{
    InvokePutProperty(wxS("Degree"), degree);
}



// ***** class wxExcelRectangularGradient PROPERTIES *****

wxExcelColorStops wxExcelRectangularGradient::GetColorStops()
{
    wxExcelColorStops colorStops;
    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("ColorStops", colorStops);
}

double wxExcelRectangularGradient::GetRectangleBottom()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("RectangleBottom");
}

void wxExcelRectangularGradient::SetRectangleBottom(double rectangleBottom)
{
    InvokePutProperty(wxS("RectangleBottom"), rectangleBottom);
}

double wxExcelRectangularGradient::GetRectangleLeft()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("RectangleLeft");
}

void wxExcelRectangularGradient::SetRectangleLeft(double rectangleLeft)
{
    InvokePutProperty(wxS("RectangleLeft"), rectangleLeft);
}

double wxExcelRectangularGradient::GetRectangleRight()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("RectangleRight");
}

void wxExcelRectangularGradient::SetRectangleRight(double rectangleRight)
{
    InvokePutProperty(wxS("RectangleRight"), rectangleRight);
}

double wxExcelRectangularGradient::GetRectangleTop()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("RectangleTop");
}

void wxExcelRectangularGradient::SetRectangleTop(double rectangleTop)
{
    InvokePutProperty(wxS("RectangleTop"), rectangleTop);
}

} // namespace wxAutoExcel
