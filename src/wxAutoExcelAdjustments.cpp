/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// License:     MIT license
/////////////////////////////////////////////////////////////////////////////


#include "wx/wxAutoExcel_prec.h"

#include "wx/wxAutoExcelAdjustments.h"

#if WXAUTOEXCEL_USE_SHAPES

#include "wx/wxAutoExcel_private.h"

namespace wxAutoExcel {

long wxExcelAdjustments::GetCount()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("Count");
}

double wxExcelAdjustments::GetItem(long index)
{
    wxASSERT( index > 0 );
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET1("Item", index);
}

double wxExcelAdjustments::operator[](long index)
{
    return GetItem(index);
}


} // namespace wxAutoExcel

#endif // #if WXAUTOEXCEL_USE_SHAPES
