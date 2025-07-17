/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////


#include "wx/wxAutoExcel_prec.h"

#include "wx/wxAutoExcelChartFillFormat.h"

#if WXAUTOEXCEL_USE_CHARTS

#include "wx/wxAutoExcelChartColorFormat.h"

#include "wx/wxAutoExcel_private.h"

namespace wxAutoExcel {

// ***** class wxExcelChartFillFormat METHODS *****

void wxExcelChartFillFormat::OneColorGradient(MsoGradientStyle style, long variant, double degree)
{
    WXAUTOEXCEL_CALL_METHOD3_RET("OneColorGradient", (long)style, variant, degree, "null");
}

void wxExcelChartFillFormat::Patterned(MsoPatternType pattern)
{
    WXAUTOEXCEL_CALL_METHOD1_RET("Patterned", (long)pattern, "null");
}

void wxExcelChartFillFormat::PresetGradient(MsoGradientStyle style, long variant, MsoPresetGradientType presetGradientType)
{
    WXAUTOEXCEL_CALL_METHOD3_RET("PresetGradient", (long)style, variant, (long)presetGradientType, "null");
}

void wxExcelChartFillFormat::PresetTextured(MsoPresetTexture presetTexture)
{
    WXAUTOEXCEL_CALL_METHOD1_RET("PresetTextured", (long)presetTexture, "null");
}

void wxExcelChartFillFormat::Solid()
{
    WXAUTOEXCEL_CALL_METHOD0_RET("Solid", "null");
}

void wxExcelChartFillFormat::TwoColorGradient(MsoGradientStyle style, long variant)
{
    WXAUTOEXCEL_CALL_METHOD2_RET("TwoColorGradient", (long)style, variant, "null");
}

void wxExcelChartFillFormat::UserPicture(const wxString& pictureFile, XlChartPictureType* pictureFormat,
                                         double* pictureStackUnit, XlChartPicturePlacement* picturePlacement)
{
    WXAUTOEXCEL_OPTIONALCPPSTR_TO_OPTIONALVARIANT_NAME(PictureFile, pictureFile);
    WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT_NAME(PictureFormat, ((long*)pictureFormat));
    WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT_NAME(PictureStackUnit, pictureStackUnit);
    WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT_NAME(PicturePlacement, ((long*)picturePlacement));

    WXAUTOEXCEL_CALL_METHOD4_RET("UserPicture", vPictureFile, vPictureFormat, vPictureStackUnit, vPicturePlacement, "null");
}

void wxExcelChartFillFormat::UserTextured(const wxString& textureFile)
{
    WXAUTOEXCEL_CALL_METHOD1_RET("UserTextured", textureFile, "null");
}

// ***** class wxExcelChartFillFormat PROPERTIES *****

wxExcelChartColorFormat wxExcelChartFillFormat::GetBackColor()
{
    wxExcelChartColorFormat chartColorFormat;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("BackColor", chartColorFormat);
}

wxExcelChartColorFormat wxExcelChartFillFormat::GetForeColor()
{
    wxExcelChartColorFormat chartColorFormat;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("ForeColor", chartColorFormat);
}

MsoGradientColorType wxExcelChartFillFormat::GetGradientColorType()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("GradientColorType", MsoGradientColorType, msoGradientOneColor);
}

double wxExcelChartFillFormat::GetGradientDegree()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("GradientDegree");
}

MsoGradientStyle wxExcelChartFillFormat::GetGradientStyle()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("GradientStyle", MsoGradientStyle, msoGradientHorizontal);
}

long wxExcelChartFillFormat::GetGradientVariant()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("GradientVariant");
}

MsoPatternType wxExcelChartFillFormat::GetPattern()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("Pattern", MsoPatternType, msoPattern5Percent);
}

void wxExcelChartFillFormat::SetPattern(MsoPatternType pattern)
{
    InvokePutProperty(wxS("Pattern"), (long)pattern);
}

MsoPresetGradientType wxExcelChartFillFormat::GetPresetGradientType()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("PresetGradientType", MsoPresetGradientType, msoGradientEarlySunset);
}

MsoPresetTexture wxExcelChartFillFormat::GetPresetTexture()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("PresetTexture", MsoPresetTexture, msoTexturePapyrus);
}

wxString wxExcelChartFillFormat::GetTextureName()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("TextureName");
}

MsoTextureType wxExcelChartFillFormat::GetTextureType()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("TextureType", MsoTextureType, msoTexturePreset);
}

MsoFillType wxExcelChartFillFormat::GetType()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("Type", MsoFillType, msoFillSolid);
}

MsoTriState wxExcelChartFillFormat::GetVisible()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("Visible", MsoTriState, msoFalse);
}

void wxExcelChartFillFormat::SetVisible(MsoTriState visible)
{
    InvokePutProperty(wxS("Visible"), (long)visible);
}

} // namespace wxAutoExcel

#endif // #if WXAUTOEXCEL_USE_CHARTS
