/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// License:     MIT license
/////////////////////////////////////////////////////////////////////////////


#include "wx/wxAutoExcel_prec.h"

#include "wx/wxAutoExcelInterior.h"

#include "wx/wxAutoExcelGradient.h"

#include "wx/wxAutoExcel_private.h"

namespace wxAutoExcel {

wxColour wxExcelInterior::GetColor()
{
    WXAUTOEXCEL_PROPERTY_COLOR_GET0("Color");
}

void wxExcelInterior::SetColor(const wxColour& color)
{
    InvokePutProperty(wxS("Color"), (long)color.GetRGB());
}

long wxExcelInterior::GetColorIndex()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("ColorIndex");
}

void wxExcelInterior::SetColorIndex(long colorIndex)
{
    InvokePutProperty(wxS("ColorIndex"), colorIndex);
}

wxExcelLinearGradient wxExcelInterior::GetLinearGradient()
{
    wxExcelLinearGradient gradient;
    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("LinearGradient", gradient);
}

wxExcelRectangularGradient wxExcelInterior::GetRectangularGradient()
{
    wxExcelRectangularGradient gradient;
    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("RectangularGradient", gradient);
}

bool wxExcelInterior::GetInvertIfNegative()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("InvertIfNegative");
}

void wxExcelInterior::SetInvertIfNegative(bool invertIfNegative)
{
    InvokePutProperty(wxS("InvertIfNegative"), invertIfNegative);
}

XlPattern wxExcelInterior::GetPattern()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("Pattern", XlPattern, xlPatternNone);
}

void wxExcelInterior::SetPattern(const XlPattern pattern)
{
    InvokePutProperty(wxS("Pattern"), (long)pattern);
}

wxColour wxExcelInterior::GetPatternColor()
{
    WXAUTOEXCEL_PROPERTY_COLOR_GET0("PatternColor");
}

void wxExcelInterior::SetPatternColor(const wxColour& patternColor)
{
    InvokePutProperty(wxS("PatternColor"), (long)patternColor.GetRGB());
}

long wxExcelInterior::GetPatternColorIndex()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("PatternColorIndex");
}

void wxExcelInterior::SetPatternColorIndex(long patternColorIndex)
{
    InvokePutProperty(wxS("PatternColorIndex"), patternColorIndex);
}

XlThemeColor wxExcelInterior::GetPatternThemeColor()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("PatternThemeColor", XlThemeColor, xlThemeColorLight1);
}

void wxExcelInterior::SetPatternThemeColor(XlThemeColor patternThemeColor)
{
    InvokePutProperty(wxS("PatternThemeColor"), (long)patternThemeColor);
}

double wxExcelInterior::GetPatternTintAndShade()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("PatternTintAndShade");
}

void wxExcelInterior::SetPatternTintAndShade(double patternTintAndShade)
{
    InvokePutProperty(wxS("PatternTintAndShade"), patternTintAndShade);
}

XlThemeColor wxExcelInterior::GetThemeColor()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("ThemeColor", XlThemeColor, xlThemeColorLight1);
}

void wxExcelInterior::SetThemeColor(XlThemeColor themeColor)
{
    InvokePutProperty(wxS("ThemeColor"), (long)themeColor);
}

double wxExcelInterior::GetTintAndShade()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("TintAndShade");
}

void wxExcelInterior::SetTintAndShade(double tintAndShade)
{
    InvokePutProperty(wxS("TintAndShade"), tintAndShade);
}


} // namespace wxAutoExcel
