/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////


#include "wx/wxAutoExcel_prec.h"

#include "wx/wxAutoExcelPictureFormat.h"

#if WXAUTOEXCEL_USE_SHAPES || WXAUTOEXCEL_USE_CHARTS

#include "wx/wxAutoExcel_private.h"

namespace wxAutoExcel {

// ***** class wxExcelPictureFormat METHODS *****

void wxExcelPictureFormat::IncrementBrightness(double increment)
{
    WXAUTOEXCEL_CALL_METHOD1_RET("IncrementBrightness", increment, "null");
}

void wxExcelPictureFormat::IncrementContrast(double increment)
{
    WXAUTOEXCEL_CALL_METHOD1_RET("IncrementContrast", increment, "null");
}

// ***** class wxExcelPictureFormat PROPERTIES *****

double wxExcelPictureFormat::GetBrightness()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("Brightness");
}

void wxExcelPictureFormat::SetBrightness(double brightness)
{
    InvokePutProperty(wxS("Brightness"), brightness);
}

MsoPictureColorType wxExcelPictureFormat::GetColorType()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("ColorType", MsoPictureColorType, msoPictureAutomatic);
}

void wxExcelPictureFormat::SetColorType(MsoPictureColorType colorType)
{
    InvokePutProperty(wxS("ColorType"), (long)colorType);
}

double wxExcelPictureFormat::GetContrast()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("Contrast");
}

void wxExcelPictureFormat::SetContrast(double contrast)
{
    InvokePutProperty(wxS("Contrast"), contrast);
}

double wxExcelPictureFormat::GetCropBottom()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("CropBottom");
}

void wxExcelPictureFormat::SetCropBottom(double cropBottom)
{
    InvokePutProperty(wxS("CropBottom"), cropBottom);
}

double wxExcelPictureFormat::GetCropLeft()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("CropLeft");
}

void wxExcelPictureFormat::SetCropLeft(double cropLeft)
{
    InvokePutProperty(wxS("CropLeft"), cropLeft);
}

double wxExcelPictureFormat::GetCropRight()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("CropRight");
}

void wxExcelPictureFormat::SetCropRight(double cropRight)
{
    InvokePutProperty(wxS("CropRight"), cropRight);
}

double wxExcelPictureFormat::GetCropTop()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("CropTop");
}

void wxExcelPictureFormat::SetCropTop(double cropTop)
{
    InvokePutProperty(wxS("CropTop"), cropTop);
}

wxColour wxExcelPictureFormat::GetTransparencyColor()
{
    WXAUTOEXCEL_PROPERTY_COLOR_GET0("TransparencyColor");
}

void wxExcelPictureFormat::SetTransparencyColor(const wxColour& transparencyColor)
{
    InvokePutProperty(wxS("TransparencyColor"), (long)transparencyColor.GetRGB());
}

MsoTriState wxExcelPictureFormat::GetTransparentBackground()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("TransparentBackground", MsoTriState, msoFalse);
}

void wxExcelPictureFormat::SetTransparentBackground(MsoTriState transparentBackground)
{
    InvokePutProperty(wxS("TransparentBackground"), (long)transparentBackground);
}

} // namespace wxAutoExcel

#endif // #if WXAUTOEXCEL_USE_SHAPES || WXAUTOEXCEL_USE_CHARTS
