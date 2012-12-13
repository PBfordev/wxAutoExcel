/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Modified by:
// RCS-ID:      $Id: $
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////


#include "wx/wxAutoExcel_prec.h"

#include "wx/wxAutoExcelWalls.h"

#if WXAUTOEXCEL_USE_CHARTS

#include "wx/wxAutoExcelChartFormat.h"

#include "wx/wxAutoExcelPrivate.h"

namespace wxAutoExcel {

// ***** class wxExcelWalls METHODS *****

bool wxExcelWalls::ClearFormats()
{
    WXAUTOEXCEL_CALL_METHOD0_BOOL("ClearFormats");
}


bool wxExcelWalls::Paste()
{
    WXAUTOEXCEL_CALL_METHOD0_BOOL("Paste");
}

bool wxExcelWalls::Select()
{
    WXAUTOEXCEL_CALL_METHOD0_BOOL("Select");
}

// ***** class wxExcelWalls PROPERTIES *****


wxExcelChartFormat wxExcelWalls::GetFormat()
{
    wxExcelChartFormat chartFormat;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("Format", chartFormat);
}

wxString wxExcelWalls::GetName()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("Name");
}

long wxExcelWalls::GetPictureType()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("PictureType");
}

void wxExcelWalls::SetPictureType(long pictureType)
{
    InvokePutProperty(wxS("PictureType"), pictureType);
}

long wxExcelWalls::GetPictureUnit()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("PictureUnit");
}

void wxExcelWalls::SetPictureUnit(long pictureUnit)
{
    InvokePutProperty(wxS("PictureUnit"), pictureUnit);
}

long wxExcelWalls::GetThickness()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("Thickness");
}

void wxExcelWalls::SetThickness(long thickness)
{
    InvokePutProperty(wxS("Thickness"), thickness);
}

} // namespace wxAutoExcel

#endif // #if WXAUTOEXCEL_USE_CHARTS
