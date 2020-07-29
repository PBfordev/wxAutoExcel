/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Modified by:
// Copyright:   (c) 2020 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////

#include "wx/wxAutoExcel_prec.h"

#include "wx/wxAutoExcelDataBarBorder.h"

#if WXAUTOEXCEL_USE_CONDFORMAT

#include "wx/wxAutoExcel_private.h"

namespace wxAutoExcel {

// ***** class wxExcelDataBarBorder PROPERTIES *****

wxColour wxExcelDataBarBorder::GetColor()
{
    WXAUTOEXCEL_PROPERTY_COLOR_GET0("Color");
}

XlDataBarBorderType wxExcelDataBarBorder::GetType()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("Type", XlDataBarBorderType, xlDataBarBorderNone);
}

void wxExcelDataBarBorder::SetType(XlDataBarBorderType type)
{
    InvokePutProperty(wxS("Type"), (long)type);
}

} // namespace wxAutoExcel 

#endif // WXAUTOEXCEL_USE_CONDFORMAT