/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Modified by:
// Copyright:   (c) 2020 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////

#include "wx/wxAutoExcel_prec.h"

#include "wx/wxAutoExcelUserAccess.h"

#include "wx/wxAutoExcel_private.h"

namespace wxAutoExcel {

// ***** class wxExcelUserAccess METHODS *****

void wxExcelUserAccess::Delete()
{
    WXAUTOEXCEL_CALL_METHOD0_RET("Delete", "null");
}

// ***** class wxExcelUserAccess PROPERTIES *****

bool wxExcelUserAccess::GetAllowEdit()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("AllowEdit");
}

void wxExcelUserAccess::SetAllowEdit(bool allowEdit)
{
    InvokePutProperty(wxS("AllowEdit"), allowEdit);
}

wxString wxExcelUserAccess::GetName()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("Name");
}

void wxExcelUserAccess::SetName(const wxString& name)
{
    InvokePutProperty(wxS("Name"), name);
}

// ***** class wxExcelUserAccessList METHODS *****

wxExcelUserAccess wxExcelUserAccessList::Add(const wxString& name, bool allowEdit)
{
    wxExcelUserAccess object;

    WXAUTOEXCEL_CALL_METHOD2("Add", name, allowEdit, "void*", object);
    VariantToObject(vResult, &object);
    return object;
}


// ***** class wxExcelUserAccessList PROPERTIES *****

long wxExcelUserAccessList::GetCount()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("Count");
}

wxExcelUserAccess wxExcelUserAccessList::GetItem(long index)
{
    wxASSERT( index > 0 );

    wxExcelUserAccess object;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET1("Item", index, object);
}

wxExcelUserAccess wxExcelUserAccessList::GetItem(const wxString& name)
{
     wxASSERT( !name.empty() );

    wxExcelUserAccess object;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET1("Item", name, object);
}

wxExcelUserAccess wxExcelUserAccessList::operator[](long index)
{
    return GetItem(index);
}

wxExcelUserAccess wxExcelUserAccessList::operator[](const wxString& name)
{
    return GetItem(name);
}

} // namespace wxAutoExcel
