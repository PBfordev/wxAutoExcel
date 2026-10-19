/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// License:     MIT license
/////////////////////////////////////////////////////////////////////////////


#include "wx/wxAutoExcel_prec.h"

#include "wx/wxAutoExcelTab.h"

#include "wx/wxAutoExcelApplication.h"

#include "wx/wxAutoExcel_private.h"

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
    wxVariant vResult;
    if ( !InvokeGetProperty(wxS("TintAndShade"), vResult) ) return 0;

    // Excel returns False when the tab has no color
    if ( vResult.IsType(wxS("bool")) )
        return 0;

    WXAUTOEXCEL_CHECK_VARIANT_NUMBER(vResult, "double", "TintAndShade", 0);
    return vResult.GetDouble();
}

void wxExcelTab::SetTintAndShade(double tintAndShade)
{
    InvokePutProperty(wxS("TintAndShade"), tintAndShade);
}

} // namespace wxAutoExcel
