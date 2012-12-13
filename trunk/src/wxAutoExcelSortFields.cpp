/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Modified by:
// RCS-ID:      $Id: $
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////


#include "wx/wxAutoExcel_prec.h"

#include "wx/wxAutoExcelSortFields.h"

#include "wx/wxAutoExcelRange.h"
#include "wx/wxAutoExcelIcon.h"

#include "wx/wxAutoExcelPrivate.h"

namespace wxAutoExcel {

// ***** class wxExcelSortField METHODS *****

void wxExcelSortField::Delete()
{
    WXAUTOEXCEL_CALL_METHOD0_RET("Delete", "null");
}

void wxExcelSortField::ModifyKey(wxExcelRange range)
{
    wxVariant vRange;

    if ( ObjectToVariant(&range, vRange) )
    {
        WXAUTOEXCEL_CALL_METHOD1_RET("ModifyKey", vRange, "null");
    }
}

void wxExcelSortField::SetIcon(wxExcelIcon icon)
{
    wxVariant vIcon;

    if ( ObjectToVariant(&icon, vIcon) )
    {
        WXAUTOEXCEL_CALL_METHOD1_RET("SetIcon", vIcon, "null");
    }
}

// ***** class wxExcelSortField PROPERTIES *****

wxString wxExcelSortField::GetCustomOrder()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("CustomOrder");
}

void wxExcelSortField::SetCustomOrder(const wxString& customOrder)
{
    InvokePutProperty(wxS("CustomOrder"), customOrder);
}

XlSortDataOption wxExcelSortField::GetDataOption()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("DataOption", XlSortDataOption, xlSortNormal);
}

void wxExcelSortField::SetDataOption(XlSortDataOption dataOption)
{
    InvokePutProperty(wxS("DataOption"), (long)dataOption);
}

wxExcelRange wxExcelSortField::GetKey()
{    
    wxExcelRange range;
    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("Key", range);
}

XlSortOrder wxExcelSortField::GetOrder()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("Order", XlSortOrder, xlAscending);
}

void wxExcelSortField::SetOrder(XlSortOrder order)
{
    InvokePutProperty(wxS("Order"), (long)order);
}

long wxExcelSortField::GetPriority()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("Priority");
}

void wxExcelSortField::SetPriority(long priority)
{
    InvokePutProperty(wxS("Priority"), priority);
}

XlSortOn wxExcelSortField::GetSortOn()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("SortOn", XlSortOn, SortOnValues);
}

void wxExcelSortField::SetSortOn(XlSortOn sortOn)
{
    InvokePutProperty(wxS("SortOn"), (long)sortOn);
}

wxExcelObject wxExcelSortField::GetSortOnValue()
{
    wxExcelObject object;
    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("SortOnValue", object);
}

// ***** class wxExcelSortFields METHODS *****

wxExcelSortField wxExcelSortFields::Add(wxExcelRange key, XlSortOn sortOn,
                            XlSortOrder* order, const wxString& customOrder,
                            XlSortDataOption* dataOption)
{
    wxVariant vKey, vSortOn;
    wxExcelSortField field;

    if ( ObjectToVariant(&key, vKey, wxS("Key")) )
    {    
        vSortOn = (long)sortOn;
        vSortOn.SetName(wxS("SortOn"));
        WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT_NAME(Order, ((long*)order));
        WXAUTOEXCEL_OPTIONALCPPSTR_TO_OPTIONALVARIANT_NAME(CustomOrder, customOrder);
        WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT_NAME(DataOption, ((long*)dataOption));        

        WXAUTOEXCEL_CALL_METHOD5("Add", vKey, vSortOn, vOrder, vCustomOrder, vDataOption, "void*", field);        
        VariantToObject(vResult, &field);
    }
    return field;
}

void wxExcelSortFields::Clear()
{
    WXAUTOEXCEL_CALL_METHOD0_RET("Clear", "null");
}

// ***** class wxExcelSortFields PROPERTIES *****


long wxExcelSortFields::GetCount()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("Count");
}


wxExcelSortField wxExcelSortFields::GetItem(long index)
{
    wxASSERT( index > 0 );

    wxExcelSortField sortField;
    WXAUTOEXCEL_PROPERTY_OBJECT_GET1("Item", index, sortField);
}

wxExcelSortField wxExcelSortFields::operator[](long index)
{
    return GetItem(index);
}



} // namespace wxAutoExcel
