/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Modified by:
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////


#include "wx/wxAutoExcel_prec.h"

#include "wx/wxAutoExcelTab.h"

#include "wx/wxAutoExcelApplication.h"

#include "wx/wxAutoExcelPrivate.h"

namespace wxAutoExcel {

// ***** class wxExcelTab PROPERTIES *****

wxColour wxExcelTab::GetColor()
{
    WXAUTOEXCEL_PROPERTY_COLOR_GET0("Color");
}

void wxExcelTab::SetColor(const wxColour& color)
{
    InvokePutProperty(wxS("Color"), (long)color.GetRGB());
}

long wxExcelTab::GetColorIndex()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("ColorIndex");
}

void wxExcelTab::SetColorIndex(long colorIndex)
{
    InvokePutProperty(wxS("ColorIndex"), colorIndex);
}


XlThemeFont wxExcelTab::GetThemeColor()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("ThemeColor", XlThemeFont, (XlThemeFont)0);
}

void wxExcelTab::SetThemeColor(XlThemeFont themeColor)
{
    InvokePutProperty(wxS("ThemeColor"), (long)themeColor);
}

double wxExcelTab::GetTintAndShade()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("TintAndShade");
}

void wxExcelTab::SetTintAndShade(double tintAndShade)
{
    InvokePutProperty(wxS("TintAndShade"), tintAndShade);
}

} // namespace wxAutoExcel
