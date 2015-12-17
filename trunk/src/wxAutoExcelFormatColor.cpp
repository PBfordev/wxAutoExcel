/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Modified by:
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////


#include "wx/wxAutoExcel_prec.h"

#include "wx/wxAutoExcelFormatColor.h"

#if WXAUTOEXCEL_USE_CONDFORMAT

#include "wx/wxAutoExcel_private.h"

namespace wxAutoExcel {

// ***** class wxExcelFormatColor PROPERTIES *****

wxColour wxExcelFormatColor::GetColor()
{
    WXAUTOEXCEL_PROPERTY_COLOR_GET0("Color");
}

void wxExcelFormatColor::SetColor(const wxColour& color)
{
    InvokePutProperty(wxS("Color"), (long)color.GetRGB());
}

XlColorIndex wxExcelFormatColor::GetColorIndex()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("ColorIndex", XlColorIndex, xlColorIndexNone);
}

void wxExcelFormatColor::SetColorIndex(XlColorIndex colorIndex)
{
    InvokePutProperty(wxS("ColorIndex"), (long)colorIndex);
}

XlThemeColor wxExcelFormatColor::GetThemeColor()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("ThemeColor", XlThemeColor, xlThemeColorDark1);
}

void wxExcelFormatColor::SetThemeColor(XlThemeColor themeColor)
{
    InvokePutProperty(wxS("ThemeColor"), (long)themeColor);
}

double wxExcelFormatColor::GetTintAndShade()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("TintAndShade");
}

void wxExcelFormatColor::SetTintAndShade(double tintAndShade)
{
    InvokePutProperty(wxS("TintAndShade"), tintAndShade);
}

} // namespace wxAutoExcel

#endif // WXAUTOEXCEL_USE_CONDFORMAT