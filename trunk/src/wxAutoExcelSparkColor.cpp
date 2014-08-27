/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Modified by:
// RCS-ID:      $Id: $
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////


#include "wx/wxAutoExcel_prec.h"

#include "wx/wxAutoExcelSparkColor.h"

#if WXAUTOEXCEL_USE_CHARTS

#include "wx/wxAutoExcelFormatColor.h"
#include "wx/wxAutoExcelSparklineGroups.h"

#include "wx/wxAutoExcelPrivate.h"

namespace wxAutoExcel {

// ***** class wxExcelSparkColor PROPERTIES *****

wxExcelFormatColor wxExcelSparkColor::GetColor()
{
    wxExcelFormatColor formatColor;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("Color", formatColor);
}

wxExcelSparklineGroup wxExcelSparkColor::GetParent()
{
    wxExcelSparklineGroup sparklineGroup;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("Parent", sparklineGroup);
}

bool wxExcelSparkColor::GetVisible()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("Visible");
}

void wxExcelSparkColor::SetVisible(bool visible)
{
    InvokePutProperty(wxS("Visible"), visible);
}



} // namespace wxAutoExcel

#endif // #if WXAUTOEXCEL_USE_CHARTS
