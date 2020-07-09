/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Modified by:
// Copyright:   (c) 2020 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////

#include "wx/wxAutoExcel_prec.h"

#include "wx/wxAutoExcelTableStyle.h"

#include "wx/wxAutoExcelTableStyleElement.h"

#include "wx/wxAutoExcel_private.h"

namespace wxAutoExcel {

// ***** class wxExcelTableStyle METHODS *****

void wxExcelTableStyle::Delete()
{
    WXAUTOEXCEL_CALL_METHOD0_RET("Delete", "null");
}

wxExcelTableStyle wxExcelTableStyle::Duplicate(const wxString& newTableStyleName)
{
    wxExcelTableStyle style;

    WXAUTOEXCEL_OPTIONALCPPSTR_TO_OPTIONALVARIANT_NAME(NewTableStyleName, newTableStyleName);
    WXAUTOEXCEL_CALL_METHOD1_OBJECT("Duplicate", vNewTableStyleName, style);
}

// ***** class wxExcelTableStyle PROPERTIES *****

bool wxExcelTableStyle::GetBuiltIn()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("BuiltIn");
}

wxString wxExcelTableStyle::GetName()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("Name");
}

wxString wxExcelTableStyle::GetNameLocal()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("NameLocal");
}

bool wxExcelTableStyle::GetShowAsAvailablePivotTableStyle()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("ShowAsAvailablePivotTableStyle");
}

void wxExcelTableStyle::SetShowAsAvailablePivotTableStyle(bool showAsAvailablePivotTableStyle)
{
    InvokePutProperty(wxS("ShowAsAvailablePivotTableStyle"), showAsAvailablePivotTableStyle);
}

bool wxExcelTableStyle::GetShowAsAvailableSlicerStyle()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("ShowAsAvailableSlicerStyle");
}

void wxExcelTableStyle::SetShowAsAvailableSlicerStyle(bool showAsAvailableSlicerStyle)
{
    InvokePutProperty(wxS("ShowAsAvailableSlicerStyle"), showAsAvailableSlicerStyle);
}

bool wxExcelTableStyle::GetShowAsAvailableTableStyle()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("ShowAsAvailableTableStyle");
}

void wxExcelTableStyle::SetShowAsAvailableTableStyle(bool showAsAvailableTableStyle)
{
    InvokePutProperty(wxS("ShowAsAvailableTableStyle"), showAsAvailableTableStyle);
}

bool wxExcelTableStyle::GetShowAsAvailableTimelineStyle()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("ShowAsAvailableTimelineStyle");
}

void wxExcelTableStyle::SetShowAsAvailableTimelineStyle(bool showAsAvailableTimelineStyle)
{
    InvokePutProperty(wxS("ShowAsAvailableTimelineStyle"), showAsAvailableTimelineStyle);
}

wxExcelTableStyleElements wxExcelTableStyle::GetTableStyleElements()
{
    wxExcelTableStyleElements tableStyleElements;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("TableStyleElements", tableStyleElements);
}

// ***** class wxExcelTableStyles METHODS *****

wxExcelTableStyle wxExcelTableStyles::Add(const wxString& tableStyleName)
{
    wxASSERT( !tableStyleName.empty() );

    wxExcelTableStyle style;

    WXAUTOEXCEL_CALL_METHOD1_OBJECT("Add", tableStyleName, style);
}

wxExcelTableStyle wxExcelTableStyles::GetItem(long index)
{
    wxASSERT( index > 0 );

    wxExcelTableStyle style;

    WXAUTOEXCEL_CALL_METHOD1_OBJECT("Item", index, style);
}

wxExcelTableStyle wxExcelTableStyles::GetItem(const wxString& name)
{
     wxASSERT( !name.empty() );

    wxExcelTableStyle style;

    WXAUTOEXCEL_CALL_METHOD1_OBJECT("Item", name, style);
}

wxExcelTableStyle wxExcelTableStyles::operator[](long index)
{
    return GetItem(index);
}

wxExcelTableStyle wxExcelTableStyles::operator[](const wxString& name)
{
    return GetItem(name);
}

// ***** class wxExcelTableStyles PROPERTIES *****

long wxExcelTableStyles::GetCount()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("Count");
}

} // namespace wxAutoExcel
