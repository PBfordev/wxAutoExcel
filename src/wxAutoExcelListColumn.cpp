/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Modified by:
// Copyright:   (c) 2020 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////

#include "wx/wxAutoExcel_prec.h"

#include "wx/wxAutoExcelListColumn.h"

#include "wx/wxAutoExcelRange.h"

#include "wx/wxAutoExcel_private.h"

namespace wxAutoExcel {

// ***** class wxExcelListColumn METHODS *****

void wxExcelListColumn::Delete()
{
    WXAUTOEXCEL_CALL_METHOD0_RET("Delete", "null");
}

// ***** class wxExcelListColumn PROPERTIES *****

wxExcelRange wxExcelListColumn::GetDataBodyRange()
{
    wxExcelRange range;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("DataBodyRange", range);
}

long wxExcelListColumn::GetIndex()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("Index");
}

wxString wxExcelListColumn::GetName()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("Name");
}

void wxExcelListColumn::SetName(const wxString& name)
{
    InvokePutProperty(wxS("Name"), name);
}

wxExcelRange wxExcelListColumn::GetRange()
{
    wxExcelRange range;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("Range", range);
}

wxExcelRange wxExcelListColumn::GetTotal()
{
    wxExcelRange range;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("Total", range);
}

XlTotalsCalculation wxExcelListColumn::GetTotalsCalculation()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("TotalsCalculation", XlTotalsCalculation, xlTotalsCalculationNone);
}

void wxExcelListColumn::SetTotalsCalculation(XlTotalsCalculation totalsCalculation)
{
    InvokePutProperty(wxS("TotalsCalculation"), (long)totalsCalculation);
}


// ***** class wxExcelListColumns METHODS *****

wxExcelListColumn wxExcelListColumns::Add(long* position)
{
    wxExcelListColumn column;

    WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT_NAME(Position, position);

    WXAUTOEXCEL_CALL_METHOD1("Add", vPosition, "void*", column);
    VariantToObject(vResult, &column);
    return column;
}


// ***** class wxExcelListColumns PROPERTIES *****

long wxExcelListColumns::GetCount()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("Count");
}

wxExcelListColumn wxExcelListColumns::GetItem(long index)
{
    wxASSERT( index > 0 );

    wxExcelListColumn column;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET1("Item", index, column);
}

wxExcelListColumn wxExcelListColumns::GetItem(const wxString& name)
{
    wxASSERT( !name.empty() );

    wxExcelListColumn column;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET1("Item", name, column);
}

wxExcelListColumn wxExcelListColumns::operator[](long index)
{
    return GetItem(index);
}

wxExcelListColumn wxExcelListColumns::operator[](const wxString& name)
{
    return GetItem(name);
}


} // namespace wxAutoExcel
