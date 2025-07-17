/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2020 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////

#include "wx/wxAutoExcel_prec.h"

#include "wx/wxAutoExcelAuthor.h"

#include "wx/wxAutoExcel_private.h"

namespace wxAutoExcel {

// ***** class wxExcelAuthor PROPERTIES *****

wxString wxExcelAuthor::GetName()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("Name");
}

wxString wxExcelAuthor::GetProviderID()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("ProviderID");
}

wxString wxExcelAuthor::GetUserID()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("UserID");
}

} // namespace wxAutoExcel
