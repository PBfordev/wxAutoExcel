/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// License:     MIT license
/////////////////////////////////////////////////////////////////////////////


#include "wx/wxAutoExcel_prec.h"

#include "wx/wxAutoExcelPages.h"

#include "wx/wxAutoExcelHeadersFooters.h"

#include "wx/wxAutoExcel_private.h"

namespace wxAutoExcel {


// ***** class wxExcelPage PROPERTIES *****

wxExcelHeaderFooter wxExcelPage::GetCenterFooter()
{
    wxExcelHeaderFooter headerFooter;
    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("CenterFooter", headerFooter);
}

wxExcelHeaderFooter wxExcelPage::GetCenterHeader()
{
    wxExcelHeaderFooter headerFooter;
    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("CenterHeader", headerFooter);
}

wxExcelHeaderFooter wxExcelPage::GetLeftFooter()
{
    wxExcelHeaderFooter headerFooter;
    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("LeftFooter", headerFooter);
}

wxExcelHeaderFooter wxExcelPage::GetLeftHeader()
{
    wxExcelHeaderFooter headerFooter;
    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("LeftHeader", headerFooter);
}

wxExcelHeaderFooter wxExcelPage::GetRightFooter()
{
    wxExcelHeaderFooter headerFooter;
    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("RightFooter", headerFooter);
}

wxExcelHeaderFooter wxExcelPage::GetRightHeader()
{
    wxExcelHeaderFooter headerFooter;
    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("RightHeader", headerFooter);
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
