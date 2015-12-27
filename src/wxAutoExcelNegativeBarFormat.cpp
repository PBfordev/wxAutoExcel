/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Modified by:
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////


#include "wx/wxAutoExcel_prec.h"

#include "wx/wxAutoExcelNegativeBarFormat.h"

#if WXAUTOEXCEL_USE_CONDFORMAT

#include "wx/wxAutoExcelFormatColor.h"

#include "wx/wxAutoExcel_private.h"
#include "wx/wxAutoExcel_enums.h"

namespace wxAutoExcel {


// ***** class wxExcelNegativeBarFormat PROPERTIES *****

wxExcelFormatColor wxExcelNegativeBarFormat::GetBorderColor()
{
    wxExcelFormatColor formatColor;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("BorderColor", formatColor);
}

XlDataBarNegativeColorType wxExcelNegativeBarFormat::GetBorderColorType()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("BorderColorType", XlDataBarNegativeColorType, xlDataBarColor);
}

void wxExcelNegativeBarFormat::SetBorderColorType(XlDataBarNegativeColorType borderColorType)
{
    InvokePutProperty(wxS("BorderColorType"), (long)borderColorType);
}

wxExcelFormatColor wxExcelNegativeBarFormat::GetColor()
{
    wxExcelFormatColor formatColor;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("Color", formatColor);
}

XlDataBarNegativeColorType wxExcelNegativeBarFormat::GetColorType()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("ColorType", XlDataBarNegativeColorType, xlDataBarColor);
}

void wxExcelNegativeBarFormat::SetColorType(XlDataBarNegativeColorType colorType)
{
    InvokePutProperty(wxS("ColorType"), (long)colorType);
}

} // namespace wxAutoExcel

#endif // WXAUTOEXCEL_USE_CONDFORMAT