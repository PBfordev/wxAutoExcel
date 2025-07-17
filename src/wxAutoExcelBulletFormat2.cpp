/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// License:     MIT license
/////////////////////////////////////////////////////////////////////////////


#include "wx/wxAutoExcel_prec.h"

#include "wx/wxAutoExcelBulletFormat2.h"

#if WXAUTOEXCEL_USE_SHAPES || WXAUTOEXCEL_USE_CHARTS

#include "wx/wxAutoExcelFont2.h"

#include "wx/wxAutoExcel_private.h"

namespace wxAutoExcel {

// ***** class wxExcelBulletFormat2 METHODS *****

void wxExcelBulletFormat2::Picture(const wxString& fileName)
{
    WXAUTOEXCEL_CALL_METHOD1_RET("Picture", fileName, "null");
}

// ***** class wxExcelBulletFormat2 PROPERTIES *****


long wxExcelBulletFormat2::GetCharacter()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("Character");
}

void wxExcelBulletFormat2::SetCharacter(long character)
{
    InvokePutProperty(wxS("Character"), character);
}


wxExcelFont2 wxExcelBulletFormat2::GetFont()
{
    wxExcelFont2 font2;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("Font", font2);
}

long wxExcelBulletFormat2::GetNumber()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("Number");
}

double wxExcelBulletFormat2::GetRelativeSize()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("RelativeSize");
}

void wxExcelBulletFormat2::SetRelativeSize(double relativeSize)
{
    InvokePutProperty(wxS("RelativeSize"), relativeSize);
}

long wxExcelBulletFormat2::GetStartValue()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("StartValue");
}

void wxExcelBulletFormat2::SetStartValue(long startValue)
{
    InvokePutProperty(wxS("StartValue"), startValue);
}

MsoNumberedBulletStyle wxExcelBulletFormat2::GetStyle()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("Style", MsoNumberedBulletStyle, msoBulletAlphaLCPeriod);
}

void wxExcelBulletFormat2::SetStyle(MsoNumberedBulletStyle style)
{
    InvokePutProperty(wxS("Style"), (long)style);
}

MsoBulletType wxExcelBulletFormat2::GetType()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("Type", MsoBulletType, msoBulletNone);
}

void wxExcelBulletFormat2::SetType(MsoBulletType type)
{
    InvokePutProperty(wxS("Type"), (long)type);
}

MsoTriState wxExcelBulletFormat2::GetUseTextColor()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("UseTextColor", MsoTriState, msoFalse);
}

void wxExcelBulletFormat2::SetUseTextColor(MsoTriState useTextColor)
{
    InvokePutProperty(wxS("UseTextColor"), (long)useTextColor);
}

MsoTriState wxExcelBulletFormat2::GetUseTextFont()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("UseTextFont", MsoTriState, msoFalse);
}

void wxExcelBulletFormat2::SetUseTextFont(MsoTriState useTextFont)
{
    InvokePutProperty(wxS("UseTextFont"), (long)useTextFont);
}

MsoTriState wxExcelBulletFormat2::GetVisible()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("Visible", MsoTriState, msoFalse);
}

void wxExcelBulletFormat2::SetVisible(MsoTriState visible)
{
    InvokePutProperty(wxS("Visible"), (long)visible);
}

} // namespace wxAutoExcel

#endif // #if WXAUTOEXCEL_USE_SHAPES || WXAUTOEXCEL_USE_CHARTS
