/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// License:     MIT license
/////////////////////////////////////////////////////////////////////////////


#include "wx/wxAutoExcel_prec.h"

#include "wx/wxAutoExcelFont.h"

#include "wx/wxAutoExcel_private.h"

namespace wxAutoExcel {


// ***** class wxPBAutoExcelFont PROPERTIES *****


XlBackground wxExcelFont::GetBackground()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("Background", XlBackground, xlBackgroundAutomatic);
}

void wxExcelFont::SetBackground(XlBackground background)
{
    InvokePutProperty(wxS("Background"), (long)background);
}

bool wxExcelFont::GetBold()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("Bold");
}

void wxExcelFont::SetBold(bool bold)
{
    InvokePutProperty(wxS("Bold"), bold);
}

wxColour wxExcelFont::GetColor()
{
    WXAUTOEXCEL_PROPERTY_COLOR_GET0("Color");
}

void wxExcelFont::SetColor(const wxColour& color)
{
    InvokePutProperty(wxS("Color"), (long)color.GetRGB());
}

long wxExcelFont::GetColorIndex()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("ColorIndex");
}

void wxExcelFont::SetColorIndex(long colorIndex)
{
    InvokePutProperty(wxS("ColorIndex"), colorIndex);
}

wxString wxExcelFont::GetFontStyle()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("FontStyle");
}

void wxExcelFont::SetFontStyle(const wxString& fontStyle)
{
    InvokePutProperty(wxS("FontStyle"), fontStyle);
}

bool wxExcelFont::GetItalic()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("Italic");
}

void wxExcelFont::SetItalic(bool italic)
{
    InvokePutProperty(wxS("Italic"), italic);
}

wxString wxExcelFont::GetName()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("Name");
}

void wxExcelFont::SetName(const wxString& name)
{
    InvokePutProperty(wxS("Name"), name);
}


double wxExcelFont::GetSize()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("Size");
}

void wxExcelFont::SetSize(double size)
{
    InvokePutProperty(wxS("Size"), size);
}

bool wxExcelFont::GetStrikethrough()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("Strikethrough");
}

void wxExcelFont::SetStrikethrough(bool strikethrough)
{
    InvokePutProperty(wxS("Strikethrough"), strikethrough);
}

bool wxExcelFont::GetSubscript()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("Subscript");
}

void wxExcelFont::SetSubscript(bool subscript)
{
    InvokePutProperty(wxS("Subscript"), subscript);
}

bool wxExcelFont::GetSuperscript()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("Superscript");
}

void wxExcelFont::SetSuperscript(bool superscript)
{
    InvokePutProperty(wxS("Superscript"), superscript);
}

XlThemeColor wxExcelFont::GetThemeColor()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("ThemeColor", XlThemeColor, xlThemeColorLight1);
}

void wxExcelFont::SetThemeColor(XlThemeColor themeColor)
{
    InvokePutProperty(wxS("ThemeColor"), (long)themeColor);
}

XlThemeFont wxExcelFont::GetThemeFont()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("ThemeFont", XlThemeFont, xlThemeFontNone);
}

void wxExcelFont::SetThemeFont(XlThemeFont themeFont)
{
    InvokePutProperty(wxS("ThemeFont"), (long)themeFont);
}

double wxExcelFont::GetTintAndShade()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("TintAndShade");
}

void wxExcelFont::SetTintAndShade(double tintAndShade)
{
    InvokePutProperty(wxS("TintAndShade"), tintAndShade);
}

} // namespace wxAutoExcel