/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Modified by:
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////


#include "wx/wxAutoExcel_prec.h"

#include "wx/wxAutoExcelFloor.h"

#if WXAUTOEXCEL_USE_CHARTS

#include "wx/wxAutoExcelChartFormat.h"

#include "wx/wxAutoExcelPrivate.h"

namespace wxAutoExcel {

// ***** class wxExcelFloor METHODS *****

bool wxExcelFloor::ClearFormats()
{
    WXAUTOEXCEL_CALL_METHOD0_BOOL("ClearFormats");
}

bool wxExcelFloor::Paste()
{
    WXAUTOEXCEL_CALL_METHOD0_BOOL("Paste");
}

bool wxExcelFloor::Select()
{
    WXAUTOEXCEL_CALL_METHOD0_BOOL("Select");
}

// ***** class wxExcelFloor PROPERTIES *****


wxExcelChartFormat wxExcelFloor::GetFormat()
{
    wxExcelChartFormat chartFormat;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("Format", chartFormat);
}

wxString wxExcelFloor::GetName()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("Name");
}


long wxExcelFloor::GetPictureType()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("PictureType");
}

void wxExcelFloor::SetPictureType(long pictureType)
{
    InvokePutProperty(wxS("PictureType"), pictureType);
}

long wxExcelFloor::GetThickness()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("Thickness");
}

void wxExcelFloor::SetThickness(long thickness)
{
    InvokePutProperty(wxS("Thickness"), thickness);
}

} // namespace wxAutoExcel

#endif // #if WXAUTOEXCEL_USE_CHARTS
