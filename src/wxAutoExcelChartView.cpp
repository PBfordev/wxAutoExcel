/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Modified by:
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////


#include "wx/wxAutoExcel_prec.h"

#include "wx/wxAutoExcelChartView.h"

#if WXAUTOEXCEL_USE_CHARTS

#include "wx/wxAutoExcelSheet.h"

#include "wx/wxAutoExcel_private.h"

namespace wxAutoExcel {

wxExcelSheet wxExcelChartView::GetSheet()
{
    wxExcelSheet sheet;
    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("Sheet", sheet);
}

} // namespace wxAutoExcel

#endif // #if WXAUTOEXCEL_USE_CHARTS
