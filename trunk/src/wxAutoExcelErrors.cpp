/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Modified by:
// RCS-ID:      $Id: $
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////


#include "wx/wxAutoExcel_prec.h"

#include "wx/wxAutoExcelErrors.h"

#include "wx/wxAutoExcelPrivate.h"


namespace wxAutoExcel {

// ***** class wxExcelError PROPERTIES *****


bool wxExcelError::GetIgnore()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("Ignore");
}

void wxExcelError::SetIgnore(bool ignore)
{
    InvokePutProperty(wxS("Ignore"), ignore);
}

bool wxExcelError::GetValue()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("Value");
}

// ***** class wxExcelErrors PROPERTIES *****

wxExcelError wxExcelErrors::GetItem(long item)
{
    wxExcelError error;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET1("Item", item, error);
}


} // namespace wxAutoExcel
