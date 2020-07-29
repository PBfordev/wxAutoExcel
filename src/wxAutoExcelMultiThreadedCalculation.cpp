/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Modified by:
// Copyright:   (c) 2020 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////

#include "wx/wxAutoExcel_prec.h"

#include "wx/wxAutoExcelMultiThreadedCalculation.h"

#include "wx/wxAutoExcel_private.h"

namespace wxAutoExcel {

// ***** class wxExcelMultiThreadedCalculation PROPERTIES *****

bool wxExcelMultiThreadedCalculation::GetEnabled()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("Enabled");
}

void wxExcelMultiThreadedCalculation::SetEnabled(bool enabled)
{
    InvokePutProperty(wxS("Enabled"), enabled);
}

long wxExcelMultiThreadedCalculation::GetThreadCount()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("ThreadCount");
}

XlThreadMode wxExcelMultiThreadedCalculation::GetThreadMode()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("ThreadMode", XlThreadMode, xlThreadModeAutomatic);
}

void wxExcelMultiThreadedCalculation::SetThreadMode(XlThreadMode threadMode)
{
    InvokePutProperty(wxS("ThreadMode"), (long)threadMode);
}

} // namespace wxAutoExcel
