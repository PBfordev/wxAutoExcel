/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Modified by:
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////


#include "wx/wxAutoExcel_prec.h"

#include "wx/wxAutoExcelFillFormat.h"

#if WXAUTOEXCEL_USE_SHAPES || WXAUTOEXCEL_USE_CHARTS

#include "wx/wxAutoExcelColorFormat.h"

#include "wx/wxAutoExcelPrivate.h"

namespace wxAutoExcel {

// ***** class wxExcelFillFormat METHODS *****

void wxExcelFillFormat::OneColorGradient(MsoGradientStyle style, long variant, double degree)
{
    WXAUTOEXCEL_CALL_METHOD3_RET("OneColorGradient", (long)style, variant, degree, "null");
}

void wxExcelFillFormat::Patterned(MsoPatternType pattern)
{
    WXAUTOEXCEL_CALL_METHOD1_RET("Patterned", (long)pattern, "null");
}

void wxExcelFillFormat::PresetGradient(MsoGradientStyle style, long variant, MsoPresetGradientType presetGradientType)
{
    WXAUTOEXCEL_CALL_METHOD3_RET("PresetGradient", (long)style, variant, (long)presetGradientType, "null");
}

void wxExcelFillFormat::PresetTextured(MsoPresetTexture presetTexture)
{
    WXAUTOEXCEL_CALL_METHOD1_RET("PresetTextured", (long)presetTexture, "null");
}

void wxExcelFillFormat::Solid()
{
    WXAUTOEXCEL_CALL_METHOD0_RET("Solid", "null");
}

void wxExcelFillFormat::TwoColorGradient(MsoGradientStyle style, long variant)
{
    WXAUTOEXCEL_CALL_METHOD2_RET("TwoColorGradient", (long)style, variant, "null");
}

void wxExcelFillFormat::UserPicture(const wxString& pictureFile)
{
    WXAUTOEXCEL_CALL_METHOD1_RET("UserPicture", pictureFile, "null");
}

void wxExcelFillFormat::UserTextured(const wxString& pictureFile)
{
    WXAUTOEXCEL_CALL_METHOD1_RET("UserTextured", pictureFile, "null");
}

// ***** class wxExcelFillFormat PROPERTIES *****

wxExcelColorFormat wxExcelFillFormat::GetBackColor()
{
    wxExcelColorFormat colorFormat;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("BackColor", colorFormat);
}

wxExcelColorFormat wxExcelFillFormat::GetForeColor()
{
    wxExcelColorFormat colorFormat;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("ForeColor", colorFormat);
}

MsoGradientColorType wxExcelFillFormat::GetGradientColorType()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("GradientColorType", MsoGradientColorType, msoGradientOneColor);
}

double wxExcelFillFormat::GetGradientDegree()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("GradientDegree");
}

//wxExcelGradientStops wxExcelFillFormat::GetGradientStops()
//{
//    wxExcelGradientStops gradientStops;
//
//    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("GradientStops", gradientStops);
//
//}

MsoGradientStyle wxExcelFillFormat::GetGradientStyle()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("GradientStyle", MsoGradientStyle, msoGradientHorizontal);
}

long wxExcelFillFormat::GetGradientVariant()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("GradientVariant");
}

MsoPatternType  wxExcelFillFormat::GetPattern()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("Pattern", MsoPatternType, msoPattern5Percent);
}

void wxExcelFillFormat::SetPattern(MsoPatternType pattern)
{
    InvokePutProperty(wxS("Pattern"), (long)pattern);
}

MsoPresetGradientType wxExcelFillFormat::GetPresetGradientType()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("PresetGradientType", MsoPresetGradientType, msoGradientEarlySunset);
}

MsoPresetTexture wxExcelFillFormat::GetPresetTexture()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("PresetTexture", MsoPresetTexture, msoTexturePapyrus);
}

MsoTriState wxExcelFillFormat::GetRotateWithObject()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("RotateWithObject", MsoTriState, msoFalse);
}

void wxExcelFillFormat::SetRotateWithObject(MsoTriState rotateWithObject)
{
    InvokePutProperty(wxS("RotateWithObject"), (long)rotateWithObject);
}

MsoTextureAlignment wxExcelFillFormat::GetTextureAlignment()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("TextureAlignment", MsoTextureAlignment, msoTextureTopLeft);
}

double wxExcelFillFormat::GetTextureHorizontalScale()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("TextureHorizontalScale");
}

void wxExcelFillFormat::SetTextureHorizontalScale(double textureHorizontalScale)
{
    InvokePutProperty(wxS("TextureHorizontalScale"), textureHorizontalScale);
}

wxString wxExcelFillFormat::GetTextureName()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("TextureName");
}

double wxExcelFillFormat::GetTextureOffsetX()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("TextureOffsetX");
}

void wxExcelFillFormat::SetTextureOffsetX(double textureOffsetX)
{
    InvokePutProperty(wxS("TextureOffsetX"), textureOffsetX);
}

double wxExcelFillFormat::GetTextureOffsetY()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("TextureOffsetY");
}

void wxExcelFillFormat::SetTextureOffsetY(double textureOffsetY)
{
    InvokePutProperty(wxS("TextureOffsetY"), textureOffsetY);
}

MsoTriState wxExcelFillFormat::GetTextureTile()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("TextureTile", MsoTriState, msoFalse);
}

void wxExcelFillFormat::SetTextureTile(MsoTriState textureTile)
{
    InvokePutProperty(wxS("TextureTile"), (long)textureTile);
}

MsoTextureType wxExcelFillFormat::GetTextureType()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("TextureType", MsoTextureType, msoTexturePreset);
}

double wxExcelFillFormat::GetTextureVerticalScale()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("TextureVerticalScale");
}

void wxExcelFillFormat::SetTextureVerticalScale(double textureVerticalScale)
{
    InvokePutProperty(wxS("TextureVerticalScale"), textureVerticalScale);
}

double wxExcelFillFormat::GetTransparency()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("Transparency");
}

void wxExcelFillFormat::SetTransparency(double transparency)
{
    InvokePutProperty(wxS("Transparency"), transparency);
}

MsoFillType wxExcelFillFormat::GetType()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("Type", MsoFillType, msoFillSolid);
}

MsoTriState  wxExcelFillFormat::GetVisible()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("Visible", MsoTriState , msoFalse);
}

void wxExcelFillFormat::SetVisible(MsoTriState  visible)
{
    InvokePutProperty(wxS("Visible"), (long)visible);
}

} // namespace wxAutoExcel

#endif // #if WXAUTOEXCEL_USE_SHAPES || WXAUTOEXCEL_USE_CHARTS
