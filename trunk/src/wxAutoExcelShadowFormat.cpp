/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Modified by:
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////


#include "wx/wxAutoExcel_prec.h"

#include "wx/wxAutoExcelShadowFormat.h"

#if WXAUTOEXCEL_USE_SHAPES || WXAUTOEXCEL_USE_CHARTS

#include "wx/wxAutoExcelColorFormat.h"

#include "wx/wxAutoExcel_private.h"

namespace wxAutoExcel {


// ***** class wxExcelShadowFormat PROPERTIES *****

double wxExcelShadowFormat::GetBlur()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("Blur");
}

void wxExcelShadowFormat::SetBlur(double blur)
{
    InvokePutProperty(wxS("Blur"), blur);
}


wxExcelColorFormat wxExcelShadowFormat::GetForeColor()
{
    wxExcelColorFormat colorFormat;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("ForeColor", colorFormat);
}

void wxExcelShadowFormat::SetForeColor(const wxExcelColorFormat& foreColor)
{
    wxVariant vColor;
    if ( ObjectToVariant(&foreColor, vColor) )    
        InvokePutProperty(wxS("ForeColor"), vColor);
}

MsoTriState wxExcelShadowFormat::GetObscured()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("Obscured", MsoTriState, msoFalse);
}

void wxExcelShadowFormat::SetObscured(MsoTriState obscured)
{
    InvokePutProperty(wxS("Obscured"), (long)obscured);
}

double wxExcelShadowFormat::GetOffsetX()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("OffsetX");
}

void wxExcelShadowFormat::SetOffsetX(double offsetX)
{
    InvokePutProperty(wxS("OffsetX"), offsetX);
}

double wxExcelShadowFormat::GetOffsetY()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("OffsetY");
}

void wxExcelShadowFormat::SetOffsetY(double offsetY)
{
    InvokePutProperty(wxS("OffsetY"), offsetY);
}


MsoTriState wxExcelShadowFormat::GetRotateWithShape()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("RotateWithShape", MsoTriState, msoFalse);
}

void wxExcelShadowFormat::SetRotateWithShape(MsoTriState rotateWithShape)
{
    InvokePutProperty(wxS("RotateWithShape"), (long)rotateWithShape);
}

double wxExcelShadowFormat::GetSize()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("Size");
}

void wxExcelShadowFormat::SetSize(double size)
{
    InvokePutProperty(wxS("Size"), size);
}

MsoShadowStyle wxExcelShadowFormat::GetStyle()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("Style", MsoShadowStyle, msoShadowStyleInnerShadow);
}

void wxExcelShadowFormat::SetStyle(MsoShadowStyle style)
{
    InvokePutProperty(wxS("Style"), (long)style);
}

double wxExcelShadowFormat::GetTransparency()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("Transparency");
}

void wxExcelShadowFormat::SetTransparency(double transparency)
{
    InvokePutProperty(wxS("Transparency"), transparency);
}

MsoShadowType  wxExcelShadowFormat::GetType()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("Type", MsoShadowType, msoShadow1);
}

void wxExcelShadowFormat::SetType(MsoShadowType  type)
{
    InvokePutProperty(wxS("Type"), (long)type);
}

MsoTriState wxExcelShadowFormat::GetVisible()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("Visible", MsoTriState, msoFalse);
}

void wxExcelShadowFormat::SetVisible(MsoTriState visible)
{
    InvokePutProperty(wxS("Visible"), (long)visible);
}

} // namespace wxAutoExcel

#endif // #if WXAUTOEXCEL_USE_SHAPES || WXAUTOEXCEL_USE_CHARTS
