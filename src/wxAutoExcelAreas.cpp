/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////


#include "wx/wxAutoExcel_prec.h"


#include "wx/wxAutoExcelAreas.h"

#include "wx/wxAutoExcelRange.h"

#include "wx/wxAutoExcel_private.h"

namespace wxAutoExcel {


long wxExcelAreas::GetCount()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("Count");
}


wxExcelRange wxExcelAreas::GetItem(long index)
{
    wxASSERT( index > 0 );
    wxExcelRange range;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET1("Item", index, range);
}


wxExcelRange wxExcelAreas::operator[](long index)
{
    return GetItem(index);
}


} // namespace wxAutoExcel
