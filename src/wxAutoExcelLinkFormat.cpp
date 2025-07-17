/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////


#include "wx/wxAutoExcel_prec.h"

#include "wx/wxAutoExcelLinkFormat.h"

#if WXAUTOEXCEL_USE_SHAPES

#include "wx/wxAutoExcel_private.h"

namespace wxAutoExcel {

// ***** class wxExcelLinkFormat METHODS *****

void wxExcelLinkFormat::Update()
{
    WXAUTOEXCEL_CALL_METHOD0_RET("Update", "null");
}

// ***** class wxExcelLinkFormat PROPERTIES *****

bool wxExcelLinkFormat::GetAutoUpdate()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("AutoUpdate");
}

bool wxExcelLinkFormat::GetLocked()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("Locked");
}

void wxExcelLinkFormat::SetLocked(bool locked)
{
    InvokePutProperty(wxS("Locked"), locked);
}

} // namespace wxAutoExcel

#endif // #if WXAUTOEXCEL_USE_SHAPES
