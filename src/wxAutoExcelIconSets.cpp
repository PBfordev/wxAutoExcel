/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// License:     MIT license
/////////////////////////////////////////////////////////////////////////////


#include "wx/wxAutoExcel_prec.h"

#include "wx/wxAutoExcelIconSets.h"

#if WXAUTOEXCEL_USE_CONDFORMAT

#include "wx/wxAutoExcel_private.h"

#include "wx/wxAutoExcelIcon.h"

namespace wxAutoExcel {

// ***** class wxExcelIconSet PROPERTIES *****

long wxExcelIconSet::GetCount()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("Count");
}

XlIconSet wxExcelIconSet::GetID()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("ID", XlIconSet, xl3Arrows);
}

wxExcelIcon wxExcelIconSet::GetItem(long index)
{
    wxExcelIcon icon;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET1("Item", index, icon);
}


// ***** class wxExcelIconSets PROPERTIES *****

long wxExcelIconSets::GetCount()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("Count");
}

wxExcelIconSet wxExcelIconSets::GetItem(long index)
{
    wxASSERT( index > 0 );

    wxExcelIconSet iconSet;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET1("Item", index, iconSet);
}

wxExcelIconSet wxExcelIconSets::operator[](long index)
{
    return GetItem(index);
}


} // namespace wxAutoExcel

#endif // WXAUTOEXCEL_USE_CONDFORMAT
