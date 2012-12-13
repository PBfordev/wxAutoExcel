/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Modified by:
// RCS-ID:      $Id: $
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////


#include "wx/wxAutoExcel_prec.h"

#include "wx/wxAutoExcelCharacters.h"

#include "wx/wxAutoExcelFont.h"

#include "wx/wxAutoExcelPrivate.h"

namespace wxAutoExcel {

// ***** class wxExcelCharacters METHODS *****

void wxExcelCharacters::Delete()
{
    WXAUTOEXCEL_CALL_METHOD0_RET("Delete", "null");
}

wxString wxExcelCharacters::Insert(const wxString& str)
{
    WXAUTOEXCEL_CALL_METHOD1_STRING("Insert", str);
}

// ***** class wxExcelCharacters PROPERTIES *****

wxString wxExcelCharacters::GetCaption()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("Caption");
}

void wxExcelCharacters::SetCaption(wxString& caption)
{
    InvokePutProperty(wxS("Caption"), caption);
}

long wxExcelCharacters::GetCount()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("Count");
}

wxExcelFont wxExcelCharacters::GetFont()
{
    wxExcelFont font;
    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("Font", font);
}

wxString wxExcelCharacters::GetText()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("Text");
}

void wxExcelCharacters::SetText(const wxString& text)
{
    InvokePutProperty(wxS("Text"), text);
}


} // namespace wxAutoExcel