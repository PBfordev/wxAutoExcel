/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Modified by:
// Copyright:   (c) 2020 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////

#include "wx/wxAutoExcel_prec.h"

#include "wx/wxAutoExcelCustomProperties.h"

#include "wx/wxAutoExcel_private.h"

namespace wxAutoExcel {

// ***** class wxExcelCustomProperty METHODS *****

void wxExcelCustomProperty::Delete()
{
    WXAUTOEXCEL_CALL_METHOD0_RET("Delete", "null");
}

// ***** class wxExcelCustomProperty PROPERTIES *****

wxString wxExcelCustomProperty::GetName()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("Name");
}

void wxExcelCustomProperty::SetName(const wxString& name)
{
    InvokePutProperty(wxS("Name"), name);
}

wxVariant wxExcelCustomProperty::GetValue()
{
    wxVariant value;

    InvokeGetProperty("Value", value);
    return value;
}

void wxExcelCustomProperty::SetValue(const wxVariant& value)
{
    InvokePutProperty("Value", value);
}


// ***** class wxExcelCustomProperties METHODS *****

wxExcelCustomProperty wxExcelCustomProperties::Add(const wxString& name, const wxVariant& value)
{
    wxExcelCustomProperty object;

    WXAUTOEXCEL_CALL_METHOD2("Add", name, value, "void*", object);
    VariantToObject(vResult, &object);
    return object;
}


// ***** class wxExcelCustomProperties PROPERTIES *****

long wxExcelCustomProperties::GetCount()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("Count");
}

wxExcelCustomProperty wxExcelCustomProperties::GetItem(long index)
{
    wxASSERT( index > 0 );

    wxExcelCustomProperty row;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET1("Item", index, row);
}

wxExcelCustomProperty wxExcelCustomProperties::GetItem(const wxString& name)
{
     wxASSERT( !name.empty() );

    wxExcelCustomProperty row;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET1("Item", name, row);
}

wxExcelCustomProperty wxExcelCustomProperties::operator[](long index)
{
    return GetItem(index);
}

wxExcelCustomProperty wxExcelCustomProperties::operator[](const wxString& name)
{
    return GetItem(name);
}


} // namespace wxAutoExcel
