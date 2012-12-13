/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Modified by:
// RCS-ID:      $Id: $
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////


#include "wx/wxAutoExcel_prec.h"

#include "wx/wxAutoExcelReflectionFormat.h"

#if WXAUTOEXCEL_USE_SHAPES || WXAUTOEXCEL_USE_CHARTS

#include "wx/wxAutoExcelPrivate.h"

namespace wxAutoExcel {

// ***** class wxExcelReflectionFormat PROPERTIES *****


MsoReflectionType wxExcelReflectionFormat::GetType()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("Type", MsoReflectionType, msoReflectionTypeNone);
}

void wxExcelReflectionFormat::SetType(MsoReflectionType type)
{
    InvokePutProperty(wxS("Type"), (long)type);
}

} // namespace wxAutoExcel

#endif // #if WXAUTOEXCEL_USE_SHAPES || WXAUTOEXCEL_USE_CHARTS
