/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Modified by:
// RCS-ID:      $Id: $
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////


#include "wx/wxAutoExcel_prec.h"

#include "wx/wxAutoExcelTextColumn2.h"

#if WXAUTOEXCEL_USE_SHAPES || WXAUTOEXCEL_USE_CHARTS

#include "wx/wxAutoExcelPrivate.h"

namespace wxAutoExcel {


// ***** class wxExcelTextColumn2 PROPERTIES *****

long wxExcelTextColumn2::GetNumber()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("Number");
}

void wxExcelTextColumn2::SetNumber(long number)
{
    InvokePutProperty(wxS("Number"), number);
}

double wxExcelTextColumn2::GetSpacing()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("Spacing");
}

void wxExcelTextColumn2::SetSpacing(double spacing)
{
    InvokePutProperty(wxS("Spacing"), spacing);
}

MsoTextDirection wxExcelTextColumn2::GetTextDirection()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("TextDirection", MsoTextDirection, msoTextDirectionLeftToRight);
}

void wxExcelTextColumn2::SetTextDirection(MsoTextDirection textDirection)
{
    InvokePutProperty(wxS("TextDirection"), (long)textDirection);
}

} // namespace wxAutoExcel

#endif #if WXAUTOEXCEL_USE_SHAPES || WXAUTOEXCEL_USE_CHARTS
