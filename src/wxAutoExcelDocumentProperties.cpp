/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// License:     MIT license
/////////////////////////////////////////////////////////////////////////////


#include "wx/wxAutoExcel_prec.h"

#include "wx/wxAutoExcelDocumentProperties.h"


#include "wx/wxAutoExcel_private.h"

namespace wxAutoExcel {

// ***** class wxExcelDocumentProperty METHODS *****


void wxExcelDocumentProperty::Delete()
{
    WXAUTOEXCEL_CALL_METHOD0_RET("Delete", "null");
}

// ***** class wxExcelDocumentProperty PROPERTIES *****

wxString wxExcelDocumentProperty::GetLinkSource()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("LinkSource");
}

void wxExcelDocumentProperty::SetLinkSource(const wxString& linkSource)
{
    InvokePutProperty(wxS("LinkSource"), linkSource);
}

bool wxExcelDocumentProperty::GetLinkToContent()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("LinkToContent");
}

void wxExcelDocumentProperty::SetLinkToContent(bool linkToContent)
{
    InvokePutProperty(wxS("LinkToContent"), linkToContent);
}

wxString wxExcelDocumentProperty::GetName()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("Name");
}

void wxExcelDocumentProperty::SetName(const wxString& name)
{
    InvokePutProperty(wxS("Name"), name);
}

MsoDocProperties wxExcelDocumentProperty::GetType()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("Type", MsoDocProperties, msoPropertyTypeString);
}

void wxExcelDocumentProperty::SetType(MsoDocProperties propertyType)
{
    InvokePutProperty(wxS("Type"), (long)propertyType);
}

wxVariant wxExcelDocumentProperty::GetValue()
{
    wxVariant vResult;

    InvokeGetProperty(wxS("Value"), vResult);
    return vResult;
}

void wxExcelDocumentProperty::SetValue(const wxVariant& value)
{
    InvokePutProperty(wxS("Value"), value);
}


// ***** class wxExcelDocumentProperties METHODS *****


wxExcelDocumentProperty wxExcelDocumentProperties::Add(const wxString& name, bool linkToContent, MsoDocProperties propertyType,
                                                       wxVariant* value, const wxString& linkSource)
{
    wxExcelDocumentProperty prop;

    wxASSERT( !name.empty() );

    wxVariant vName(name, wxS("Name"));
    wxVariant vLinkToContent(linkToContent, wxS("LinkToContent"));
    wxVariant vPropertyType((long)propertyType, wxS("PropertyType"));

    WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT_NAME(Value, value);
    WXAUTOEXCEL_OPTIONALCPPSTR_TO_OPTIONALVARIANT_NAME(LinkSource, linkSource);

    WXAUTOEXCEL_CALL_METHOD5("Add", vResult, vLinkToContent, vPropertyType, vValue, vLinkSource, "void*", prop);
    VariantToObject(vResult, &prop);
    return prop;
}

// ***** class wxExcelDocumentProperties PROPERTIES *****

long wxExcelDocumentProperties::GetCount()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("Count");
}

wxExcelDocumentProperty wxExcelDocumentProperties::GetItem(long index)
{
    wxASSERT( index > 0 );

    wxExcelDocumentProperty item;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET1("Item", index, item);
}

wxExcelDocumentProperty wxExcelDocumentProperties::GetItem(const wxString& name)
{
    wxASSERT( !name.empty() );

    wxExcelDocumentProperty item;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET1("Item", name, item);
}

wxExcelDocumentProperty wxExcelDocumentProperties::operator[](long index)
{
    return GetItem(index);
}


wxExcelDocumentProperty wxExcelDocumentProperties::operator[](const wxString& name)
{
    return GetItem(name);
}


} // namespace wxAutoExcel

