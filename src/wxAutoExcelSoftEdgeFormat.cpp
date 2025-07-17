/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////


#include "wx/wxAutoExcel_prec.h"

#include "wx/wxAutoExcelSoftEdgeFormat.h"

#if WXAUTOEXCEL_USE_SHAPES || WXAUTOEXCEL_USE_CHARTS

#include "wx/wxAutoExcel_private.h"

namespace wxAutoExcel {


MsoSoftEdgeType wxExcelSoftEdgeFormat::GetType()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("Type", MsoSoftEdgeType, msoSoftEdgeTypeNone);
}

void wxExcelSoftEdgeFormat::SetType(MsoSoftEdgeType type)
{
    InvokePutProperty(wxS("Type"), (long)type);
}

} // namespace wxAutoExcel

#endif // #if WXAUTOEXCEL_USE_SHAPES || WXAUTOEXCEL_USE_CHARTS
