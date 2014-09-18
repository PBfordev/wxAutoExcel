/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Modified by:
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////


#include "wx/wxAutoExcel_prec.h"

#include "wx/wxAutoExcelPages.h"

#include "wx/wxAutoExcelPrivate.h"

namespace wxAutoExcel {


// ***** class wxExcelPage PROPERTIES *****

wxString wxExcelPage::GetCenterFooter()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("CenterFooter");
}

wxString wxExcelPage::GetCenterHeader()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("CenterHeader");
}

wxString wxExcelPage::GetLeftFooter()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("LeftFooter");
}

wxString wxExcelPage::GetLeftHeader()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("LeftHeader");
}

wxString wxExcelPage::GetRightFooter()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("RightFooter");
}

wxString wxExcelPage::GetRightHeader()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("RightHeader");
}

// ***** class wxExcelPages PROPERTIES *****

long wxExcelPages::GetCount()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("Count");
}

wxExcelPage wxExcelPages::GetItem(long index)
{
    wxASSERT( index > 0);

    wxExcelPage page;
    WXAUTOEXCEL_PROPERTY_OBJECT_GET1("Item", index, page);
}

wxExcelPage wxExcelPages::operator[](long index)
{
    return GetItem(index);
}

} // namespace wxAutoExcel
