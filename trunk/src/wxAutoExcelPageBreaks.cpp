/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Modified by:
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////


#include "wx/wxAutoExcel_prec.h"

#include "wx/wxAutoExcelPageBreaks.h"

#include "wx/wxAutoExcelRange.h"

#include "wx/wxAutoExcelPrivate.h"

namespace wxAutoExcel {
// ***** class wxExcelPageBreak METHODS *****

void wxExcelPageBreak::Delete()
{
    WXAUTOEXCEL_CALL_METHOD0_RET("Delete", "null");
}

void wxExcelPageBreak::DragOff(XlDirection direction, long regionIndex)
{
    WXAUTOEXCEL_CALL_METHOD2_RET("DragOff", (long)direction, regionIndex, "null");
}

// ***** class wxExcelPageBreak PROPERTIES *****


XlPageBreakExtent wxExcelPageBreak::GetExtent()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("Extent", XlPageBreakExtent, xlPageBreakFull);
}

wxExcelRange wxExcelPageBreak::GetLocation()
{
    wxExcelRange range;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("Location", range);
}

void wxExcelPageBreak::SetLocation(wxExcelRange location)
{
    wxVariant vLocation;

    if ( ObjectToVariant(&location, vLocation) )
    {
        InvokePutProperty(wxS("Location"), vLocation);
    }
}


XlPageBreak wxExcelPageBreak::GetType()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("Type", XlPageBreak, xlPageBreakAutomatic);
}

void wxExcelPageBreak::SetType(XlPageBreak type)
{
    InvokePutProperty(wxS("Type"), (long)type);
}

// ***** class wxExcelPageBreaks METHODS *****

wxExcelPageBreak wxExcelPageBreaks::Add(wxExcelRange before)
{
    wxExcelPageBreak pageBreak;
    wxVariant vBefore;

    if ( ObjectToVariant(&before, vBefore) )
    {
        WXAUTOEXCEL_CALL_METHOD1("Add", vBefore, "void*", pageBreak);
        VariantToObject(vResult, &pageBreak);
    }
    return pageBreak;    
}

// ***** class wxExcelPageBreaks PROPERTIES *****

long wxExcelPageBreaks::GetCount()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("Count");
}

wxExcelPageBreak wxExcelPageBreaks::GetItem(long index)
{
    wxASSERT( index > 0 );

    wxExcelPageBreak PageBreak;
    WXAUTOEXCEL_PROPERTY_OBJECT_GET1("Item", index, PageBreak);
}

wxExcelPageBreak wxExcelPageBreaks::operator[](long index)
{
    return GetItem(index);    
}


} // namespace wxAutoExcel
