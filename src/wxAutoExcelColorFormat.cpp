/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// License:     MIT license
/////////////////////////////////////////////////////////////////////////////


#include "wx/wxAutoExcel_prec.h"

#include "wx/wxAutoExcelColorFormat.h"

#if WXAUTOEXCEL_USE_SHAPES || WXAUTOEXCEL_USE_CHARTS

#include "wx/wxAutoExcelApplication.h"

#include "wx/wxAutoExcel_private.h"

namespace wxAutoExcel {

// ***** class wxExcelColorFormat PROPERTIES *****


double wxExcelColorFormat::GetBrightness()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("Brightness");
}

MsoThemeColorIndex wxExcelColorFormat::GetObjectThemeColor()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("ObjectThemeColor", MsoThemeColorIndex, msoNotThemeColor);
}

void wxExcelColorFormat::SetObjectThemeColor(MsoThemeColorIndex objectThemeColor)
{
    InvokePutProperty(wxS("ObjectThemeColor"), (long)objectThemeColor);
}


wxColour wxExcelColorFormat::GetRGB()
{
    WXAUTOEXCEL_PROPERTY_COLOR_GET0("RGB");
}

void wxExcelColorFormat::SetRGB(const wxColour& RGB)
{
    InvokePutProperty(wxS("RGB"), (long)RGB.GetRGB());
}

long wxExcelColorFormat::GetSchemeColor()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("SchemeColor");
}

void wxExcelColorFormat::SetSchemeColor(long schemeColor)
{
    InvokePutProperty(wxS("SchemeColor"), schemeColor);
}

double wxExcelColorFormat::GetTintAndShade()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("TintAndShade");
}

void wxExcelColorFormat::SetTintAndShade(double tintAndShade)
{
    InvokePutProperty(wxS("TintAndShade"), tintAndShade);
}

MsoColorType  wxExcelColorFormat::GetType()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("Type", MsoColorType, msoColorTypeRGB);
}

} // namespace wxAutoExcel

#endif // #if WXAUTOEXCEL_USE_SHAPES || WXAUTOEXCEL_USE_CHARTS
