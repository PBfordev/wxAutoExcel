/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Modified by:
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////


#include "wx/wxAutoExcel_prec.h"

#include "wx/wxAutoExcelTextEffectFormat.h"

#if WXAUTOEXCEL_USE_SHAPES

#include "wx/wxAutoExcel_private.h"

namespace wxAutoExcel {

// ***** class wxExcelTextEffectFormat METHODS *****

void wxExcelTextEffectFormat::ToggleVerticalText()
{
    WXAUTOEXCEL_CALL_METHOD0_RET("ToggleVerticalText", "null");
}

// ***** class wxExcelTextEffectFormat PROPERTIES *****

MsoTextEffectAlignment  wxExcelTextEffectFormat::GetAlignment()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("Alignment", MsoTextEffectAlignment , msoTextEffectAlignmentLeft);
}

MsoTriState wxExcelTextEffectFormat::GetFontBold()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("FontBold", MsoTriState, msoFalse);
}

void wxExcelTextEffectFormat::SetFontBold(MsoTriState fontBold)
{
    InvokePutProperty(wxS("FontBold"), (long)fontBold);
}

MsoTriState wxExcelTextEffectFormat::GetFontItalic()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("FontItalic", MsoTriState, msoFalse);
}

void wxExcelTextEffectFormat::SetFontItalic(MsoTriState fontItalic)
{
    InvokePutProperty(wxS("FontItalic"), (long)fontItalic);
}

wxString wxExcelTextEffectFormat::GetFontName()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("FontName");
}

void wxExcelTextEffectFormat::SetFontName(const wxString& fontName)
{
    InvokePutProperty(wxS("FontName"), fontName);
}

double wxExcelTextEffectFormat::GetFontSize()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("FontSize");
}

void wxExcelTextEffectFormat::SetFontSize(double fontSize)
{
    InvokePutProperty(wxS("FontSize"), fontSize);
}

MsoTriState wxExcelTextEffectFormat::GetKernedPairs()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("KernedPairs", MsoTriState, msoFalse);
}

void wxExcelTextEffectFormat::SetKernedPairs(MsoTriState kernedPairs)
{
    InvokePutProperty(wxS("KernedPairs"), (long)kernedPairs);
}

MsoTriState wxExcelTextEffectFormat::GetNormalizedHeight()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("NormalizedHeight", MsoTriState, msoFalse);
}

void wxExcelTextEffectFormat::SetNormalizedHeight(MsoTriState normalizedHeight)
{
    InvokePutProperty(wxS("NormalizedHeight"), (long)normalizedHeight);
}


MsoPresetTextEffectShape wxExcelTextEffectFormat::GetPresetShape()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("PresetShape", MsoPresetTextEffectShape, msoTextEffectShapePlainText);
}

void wxExcelTextEffectFormat::SetPresetShape(MsoPresetTextEffectShape presetShape)
{
    InvokePutProperty(wxS("PresetShape"), (long)presetShape);
}

MsoPresetTextEffect wxExcelTextEffectFormat::GetPresetTextEffect()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("PresetTextEffect", MsoPresetTextEffect, msoTextEffect1);
}

void wxExcelTextEffectFormat::SetPresetTextEffect(MsoPresetTextEffect presetTextEffect)
{
    InvokePutProperty(wxS("PresetTextEffect"), (long)presetTextEffect);
}

MsoTriState wxExcelTextEffectFormat::GetRotatedChars()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("RotatedChars", MsoTriState, msoFalse);
}

void wxExcelTextEffectFormat::SetRotatedChars(MsoTriState rotatedChars)
{
    InvokePutProperty(wxS("RotatedChars"), (long)rotatedChars);
}

wxString wxExcelTextEffectFormat::GetText()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("Text");
}

void wxExcelTextEffectFormat::SetText(const wxString& text)
{
    InvokePutProperty(wxS("Text"), text);
}

double wxExcelTextEffectFormat::GetTracking()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("Tracking");
}

void wxExcelTextEffectFormat::SetTracking(double tracking)
{
    InvokePutProperty(wxS("Tracking"), tracking);
}

} // namespace wxAutoExcel

#endif // #if WXAUTOEXCEL_USE_SHAPES
