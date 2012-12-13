/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Modified by:
// RCS-ID:      $Id: $
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////


#include "wx/wxAutoExcel_prec.h"

#include "wx/wxAutoExcelGraphic.h"

#include "wx/wxAutoExcelPrivate.h"

namespace wxAutoExcel {

// ***** class wxExcelGraphic PROPERTIES *****

double wxExcelGraphic::GetBrightness()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("Brightness");
}

void wxExcelGraphic::SetBrightness(double brightness)
{
    InvokePutProperty(wxS("Brightness"), brightness);
}

MsoPictureColorType wxExcelGraphic::GetColorType()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("ColorType", MsoPictureColorType, msoPictureAutomatic);
}

void wxExcelGraphic::SetColorType(MsoPictureColorType colorType)
{
    InvokePutProperty(wxS("ColorType"), (long)colorType);
}

double wxExcelGraphic::GetContrast()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("Contrast");
}

void wxExcelGraphic::SetContrast(double contrast)
{
    InvokePutProperty(wxS("Contrast"), contrast);
}

double wxExcelGraphic::GetCropBottom()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("CropBottom");
}

void wxExcelGraphic::SetCropBottom(double cropBottom)
{
    InvokePutProperty(wxS("CropBottom"), cropBottom);
}

double wxExcelGraphic::GetCropLeft()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("CropLeft");
}

void wxExcelGraphic::SetCropLeft(double cropLeft)
{
    InvokePutProperty(wxS("CropLeft"), cropLeft);
}

double wxExcelGraphic::GetCropRight()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("CropRight");
}

void wxExcelGraphic::SetCropRight(double cropRight)
{
    InvokePutProperty(wxS("CropRight"), cropRight);
}

double wxExcelGraphic::GetCropTop()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("CropTop");
}

void wxExcelGraphic::SetCropTop(double cropTop)
{
    InvokePutProperty(wxS("CropTop"), cropTop);
}

wxString wxExcelGraphic::GetFilename()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("Filename");
}

void wxExcelGraphic::SetFilename(const wxString& filename)
{
    InvokePutProperty(wxS("Filename"), filename);
}

double wxExcelGraphic::GetHeight()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("Height");
}

void wxExcelGraphic::SetHeight(double height)
{
    InvokePutProperty(wxS("Height"), height);
}

MsoTriState wxExcelGraphic::GetLockAspectRatio()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("LockAspectRatio", MsoTriState, msoFalse);
}

void wxExcelGraphic::SetLockAspectRatio(MsoTriState lockAspectRatio)
{
    InvokePutProperty(wxS("LockAspectRatio"), (long)lockAspectRatio);
}


double wxExcelGraphic::GetWidth()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("Width");
}

void wxExcelGraphic::SetWidth(double width)
{
    InvokePutProperty(wxS("Width"), width);
}

} // namespace wxAutoExcel