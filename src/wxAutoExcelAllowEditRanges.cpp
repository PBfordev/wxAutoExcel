/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Modified by:
// Copyright:   (c) 2020 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////

#include "wx/wxAutoExcel_prec.h"

#include "wx/wxAutoExcelAllowEditRanges.h"
#include "wx/wxAutoExcelRange.h"
#include "wx/wxAutoExcelUserAccess.h"

#include "wx/wxAutoExcel_private.h"

namespace wxAutoExcel {

// ***** class wxExcelAllowEditRange METHODS *****

void wxExcelAllowEditRange::ChangePassword(const wxString& password)
{
    WXAUTOEXCEL_CALL_METHOD1_RET("ChangePassword", password, "null");
}

void wxExcelAllowEditRange::Delete()
{
    WXAUTOEXCEL_CALL_METHOD0_RET("Delete", "null");
}

void wxExcelAllowEditRange::Unprotect(const wxString& password)
{
    WXAUTOEXCEL_CALL_METHOD1_RET("Unprotect", password, "null");
}

// ***** class wxExcelAllowEditRange PROPERTIES *****

wxExcelRange wxExcelAllowEditRange::GetRange()
{
    wxExcelRange range;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("Range", range);
}

wxString wxExcelAllowEditRange::GetTitle()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("Title");
}

void wxExcelAllowEditRange::SetTitle(const wxString& title)
{
    InvokePutProperty(wxS("Title"), title);
}

wxExcelUserAccessList wxExcelAllowEditRange::GetUsers()
{
    wxExcelUserAccessList userAccessList;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("Users", userAccessList);
}

// ***** class wxExcelAllowEditRanges METHODS *****

wxExcelAllowEditRange wxExcelAllowEditRanges::Add(const wxString& title, wxExcelRange& range,
                                                  const wxString& password)
{
    wxCHECK_MSG(!title.empty(), wxExcelAllowEditRange(), "title cannot be empty");
    wxCHECK_MSG(!range.IsOk_(), wxExcelAllowEditRange(), "range must be a valid Range");

    wxExcelAllowEditRange object;
    wxVariant vRange;

    if ( !ObjectToVariant(&range, vRange, "Range") )
        return object;

    WXAUTOEXCEL_OPTIONALCPPSTR_TO_OPTIONALVARIANT_NAME(Password, password);

    WXAUTOEXCEL_CALL_METHOD3("Add", title, vRange, vPassword, "void*", object);
    VariantToObject(vResult, &object);
    return object;
}


// ***** class wxExcelAllowEditRanges PROPERTIES *****

long wxExcelAllowEditRanges::GetCount()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("Count");
}

wxExcelAllowEditRange wxExcelAllowEditRanges::GetItem(long index)
{
    wxASSERT( index > 0 );

    wxExcelAllowEditRange object;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET1("Item", index, object);
}

wxExcelAllowEditRange wxExcelAllowEditRanges::GetItem(const wxString& name)
{
     wxASSERT( !name.empty() );

    wxExcelAllowEditRange object;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET1("Item", name, object);
}

wxExcelAllowEditRange wxExcelAllowEditRanges::operator[](long index)
{
    return GetItem(index);
}

wxExcelAllowEditRange wxExcelAllowEditRanges::operator[](const wxString& name)
{
    return GetItem(name);
}


} // namespace wxAutoExcel
