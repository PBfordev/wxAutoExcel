/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////


#include "wx/wxAutoExcel_prec.h"

#include "wx/wxAutoExcelFont2.h"

#if WXAUTOEXCEL_USE_SHAPES || WXAUTOEXCEL_USE_CHARTS

#include "wx/wxAutoExcelColorFormat.h"
#include "wx/wxAutoExcelFillFormat.h"
#include "wx/wxAutoExcelGlowFormat.h"
#include "wx/wxAutoExcelLineFormat.h"
#include "wx/wxAutoExcelReflectionFormat.h"
#include "wx/wxAutoExcelShadowFormat.h"
#include "wx/wxAutoExcelTextEffectFormat.h"

#include "wx/wxAutoExcel_private.h"

namespace wxAutoExcel {

// ***** class wxExcelFont2 PROPERTIES *****

MsoTriState wxExcelFont2::GetAllcaps()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("Allcaps", MsoTriState, msoFalse);
}

void wxExcelFont2::SetAllcaps(MsoTriState allcaps)
{
    InvokePutProperty(wxS("Allcaps"), (long)allcaps);
}


MsoTriState wxExcelFont2::GetAutorotateNumbers()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("AutorotateNumbers", MsoTriState, msoFalse);
}

void wxExcelFont2::SetAutorotateNumbers(MsoTriState autorotateNumbers)
{
    InvokePutProperty(wxS("AutorotateNumbers"), (long)autorotateNumbers);
}

double wxExcelFont2::GetBaselineOffset()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("BaselineOffset");
}

void wxExcelFont2::SetBaselineOffset(double baselineOffset)
{
    InvokePutProperty(wxS("BaselineOffset"), baselineOffset);
}

MsoTriState wxExcelFont2::GetBold()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("Bold", MsoTriState, msoFalse);
}

void wxExcelFont2::SetBold(MsoTriState bold)
{
    InvokePutProperty(wxS("Bold"), (long)bold);
}

MsoTextCaps wxExcelFont2::GetCaps()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("Caps", MsoTextCaps, msoNoCaps );
}

void wxExcelFont2::SetCaps(MsoTextCaps caps)
{
    InvokePutProperty(wxS("Caps"), (long)caps);
}


MsoTriState wxExcelFont2::GetDoubleStrikeThrough()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("DoubleStrikeThrough", MsoTriState, msoFalse);
}

void wxExcelFont2::SetDoubleStrikeThrough(MsoTriState doubleStrikeThrough)
{
    InvokePutProperty(wxS("DoubleStrikeThrough"), (long)doubleStrikeThrough);
}

MsoTriState wxExcelFont2::GetEmbeddable()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("Embeddable", MsoTriState, msoFalse);
}

MsoTriState wxExcelFont2::GetEmbedded()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("Embedded", MsoTriState, msoFalse);
}

MsoTriState wxExcelFont2::GetEqualize()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("Equalize", MsoTriState, msoFalse);
}

void wxExcelFont2::SetEqualize(MsoTriState equalize)
{
    InvokePutProperty(wxS("Equalize"), (long)equalize);
}

wxExcelFillFormat wxExcelFont2::GetFill()
{
    wxExcelFillFormat fillFormat;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("Fill", fillFormat);
}

wxExcelGlowFormat wxExcelFont2::GetGlow()
{
    wxExcelGlowFormat glowFormat;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("Glow", glowFormat);
}

wxExcelColorFormat wxExcelFont2::GetHighlight()
{
    wxExcelColorFormat colorFormat;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("Highlight", colorFormat);
}

MsoTriState wxExcelFont2::GetItalic()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("Italic", MsoTriState, msoFalse);
}

void wxExcelFont2::SetItalic(MsoTriState italic)
{
    InvokePutProperty(wxS("Italic"), (long)italic);
}

double wxExcelFont2::GetKerning()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("Kerning");
}

void wxExcelFont2::SetKerning(double kerning)
{
    InvokePutProperty(wxS("Kerning"), kerning);
}

wxExcelLineFormat wxExcelFont2::GetLine()
{
    wxExcelLineFormat lineFormat;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("Line", lineFormat);
}

wxString wxExcelFont2::GetName()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("Name");
}

void wxExcelFont2::SetName(const wxString& name)
{
    InvokePutProperty(wxS("Name"), name);
}

wxString wxExcelFont2::GetNameAscii()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("NameAscii");
}

void wxExcelFont2::SetNameAscii(const wxString& nameAscii)
{
    InvokePutProperty(wxS("NameAscii"), nameAscii);
}

wxString wxExcelFont2::GetNameComplexScript()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("NameComplexScript");
}

void wxExcelFont2::SetNameComplexScript(const wxString& nameComplexScript)
{
    InvokePutProperty(wxS("NameComplexScript"), nameComplexScript);
}

wxString wxExcelFont2::GetNameFarEast()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("NameFarEast");
}

void wxExcelFont2::SetNameFarEast(const wxString& nameFarEast)
{
    InvokePutProperty(wxS("NameFarEast"), nameFarEast);
}

wxString wxExcelFont2::GetNameOther()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("NameOther");
}

void wxExcelFont2::SetNameOther(const wxString& nameOther)
{
    InvokePutProperty(wxS("NameOther"), nameOther);
}


wxExcelReflectionFormat wxExcelFont2::GetReflection()
{
    wxExcelReflectionFormat reflectionFormat;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("Reflection", reflectionFormat);
}

wxExcelShadowFormat wxExcelFont2::GetShadow()
{
    wxExcelShadowFormat shadowFormat;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("Shadow", shadowFormat);
}

double wxExcelFont2::GetSize()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("Size");
}

void wxExcelFont2::SetSize(double size)
{
    InvokePutProperty(wxS("Size"), size);
}

MsoTriState wxExcelFont2::GetSmallcaps()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("Smallcaps", MsoTriState, msoFalse);
}

void wxExcelFont2::SetSmallcaps(MsoTriState smallcaps)
{
    InvokePutProperty(wxS("Smallcaps"), (long)smallcaps);
}

MsoSoftEdgeType wxExcelFont2::GetSoftEdgeFormat()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("SoftEdgeFormat", MsoSoftEdgeType, msoSoftEdgeTypeNone);
}

void wxExcelFont2::SetSoftEdgeFormat(MsoSoftEdgeType softEdgeFormat)
{
    InvokePutProperty(wxS("SoftEdgeFormat"), (long)softEdgeFormat);
}

double wxExcelFont2::GetSpacing()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("Spacing");
}

void wxExcelFont2::SetSpacing(double spacing)
{
    InvokePutProperty(wxS("Spacing"), spacing);
}

MsoTextStrike wxExcelFont2::GetStrike()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("Strike", MsoTextStrike, msoNoStrike);
}

void wxExcelFont2::SetStrike(MsoTextStrike strike)
{
    InvokePutProperty(wxS("Strike"), (long)strike);
}

MsoTriState wxExcelFont2::GetStrikeThrough()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("StrikeThrough", MsoTriState, msoFalse);
}

void wxExcelFont2::SetStrikeThrough(MsoTriState strikeThrough)
{
    InvokePutProperty(wxS("StrikeThrough"), (long)strikeThrough);
}

MsoTriState wxExcelFont2::GetSubscript()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("Subscript", MsoTriState, msoFalse);
}

void wxExcelFont2::SetSubscript(MsoTriState subscript)
{
    InvokePutProperty(wxS("Subscript"), (long)subscript);
}

MsoTriState wxExcelFont2::GetSuperscript()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("Superscript", MsoTriState, msoFalse);
}

void wxExcelFont2::SetSuperscript(MsoTriState superscript)
{
    InvokePutProperty(wxS("Superscript"), (long)superscript);
}

wxExcelColorFormat wxExcelFont2::GetUnderlineColor()
{
    wxExcelColorFormat colorFormat;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("UnderlineColor", colorFormat);
}

MsoTextUnderlineType wxExcelFont2::GetUnderlineStyle()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("UnderlineStyle", MsoTextUnderlineType, msoNoUnderline);
}

void wxExcelFont2::SetUnderlineStyle(MsoTextUnderlineType underlineStyle)
{
    InvokePutProperty(wxS("UnderlineStyle"), (long)underlineStyle);
}

MsoPresetTextEffect wxExcelFont2::GetWordArtformat()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("WordArtformat", MsoPresetTextEffect, msoTextEffect1);
}

void wxExcelFont2::SetWordArtformat(MsoPresetTextEffect wordArtformat)
{
    InvokePutProperty(wxS("WordArtformat"), (long)wordArtformat);
}

} // namespace wxAutoExcel


#endif // #if WXAUTOEXCEL_USE_SHAPES || WXAUTOEXCEL_USE_CHARTS
