/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Modified by:
// RCS-ID:      $Id: $
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////


#include "wx/wxAutoExcel_prec.h"

#include "wx/wxAutoExcelGlowFormat.h"

#if WXAUTOEXCEL_USE_SHAPES || WXAUTOEXCEL_USE_CHARTS

#include "wx/wxAutoExcelColorFormat.h"

#include "wx/wxAutoExcelPrivate.h"

namespace wxAutoExcel {

// ***** class wxExcelGlowFormat PROPERTIES *****

wxExcelColorFormat wxExcelGlowFormat::GetColor()
{
    wxExcelColorFormat color;
    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("Color", color);    
}

double wxExcelGlowFormat::GetRadius()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("Radius");
}

void wxExcelGlowFormat::SetRadius(double radius)
{
    InvokePutProperty(wxS("Radius"), radius);
}

} // namespace wxAutoExcel

#endif // #if WXAUTOEXCEL_USE_SHAPES || WXAUTOEXCEL_USE_CHARTS
