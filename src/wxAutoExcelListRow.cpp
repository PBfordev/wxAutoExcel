/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2020 PB <pbfordev@gmail.com>
// License:     MIT license
/////////////////////////////////////////////////////////////////////////////

#include "wx/wxAutoExcel_prec.h"

#include "wx/wxAutoExcelListRow.h"

#include "wx/wxAutoExcelRange.h"

#include "wx/wxAutoExcel_private.h"

namespace wxAutoExcel {

// ***** class wxExcelListRow METHODS *****

void wxExcelListRow::Delete()
{
    WXAUTOEXCEL_CALL_METHOD0_RET("Delete", "null");
}

// ***** class wxExcelListRow PROPERTIES *****

long wxExcelListRow::GetIndex()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("Index");
}

wxExcelRange wxExcelListRow::GetRange()
{
    wxExcelRange range;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("Range", range);
}


// ***** class wxExcelListRows METHODS *****

wxExcelListRow wxExcelListRows::Add(long* position, wxXlTribool alwaysInsert)
{
    wxExcelListRow row;

    WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT_NAME(Position, position);
    WXAUTOEXCEL_OPTIONALCPPTBOOL_TO_OPTIONALVARIANT_NAME(AlwaysInsert, alwaysInsert);

    WXAUTOEXCEL_CALL_METHOD2("Add", vPosition, vAlwaysInsert, "void*", row);
    VariantToObject(vResult, &row);
    return row;
}


// ***** class wxExcelListRows PROPERTIES *****

long wxExcelListRows::GetCount()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("Count");
}

wxExcelListRow wxExcelListRows::GetItem(long index)
{
    wxASSERT( index > 0 );

    wxExcelListRow row;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET1("Item", index, row);
}

wxExcelListRow wxExcelListRows::GetItem(const wxString& name)
{
     wxASSERT( !name.empty() );

    wxExcelListRow row;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET1("Item", name, row);
}

wxExcelListRow wxExcelListRows::operator[](long index)
{
    return GetItem(index);
}

wxExcelListRow wxExcelListRows::operator[](const wxString& name)
{
    return GetItem(name);
}


} // namespace wxAutoExcel
