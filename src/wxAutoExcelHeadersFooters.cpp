/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// License:     MIT license
/////////////////////////////////////////////////////////////////////////////


#include "wx/wxAutoExcel_prec.h"

#include "wx/wxAutoExcelHeadersFooters.h"

#include "wx/wxAutoExcel_private.h"

namespace wxAutoExcel {

// ***** class wxExcelHeaderFooter PROPERTIES *****

wxString wxExcelHeaderFooter::GetText()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("Text");
}

void wxExcelHeaderFooter::SetText(const wxString& text)
{
    InvokePutProperty(wxS("Text"), text);
}

} // namespace wxAutoExcel