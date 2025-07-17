/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////


#include "wx/wxAutoExcel_prec.h"

#include "wx/wxAutoExcelAddIns.h"

#include "wx/wxAutoExcel_private.h"


namespace wxAutoExcel {

// ***** class wxExcelAddIn PROPERTIES *****

wxString wxExcelAddIn::GetCLSID()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("CLSID");
}

wxString wxExcelAddIn::GetFullName()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("FullName");
}

bool wxExcelAddIn::GetInstalled()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("Installed");
}

void wxExcelAddIn::SetInstalled(bool installed)
{
    InvokePutProperty(wxS("Installed"), installed);
}

bool wxExcelAddIn::GetIsOpen()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("IsOpen");
}

wxString wxExcelAddIn::GetName()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("Name");
}

wxString wxExcelAddIn::GetPath()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("Path");
}

wxString wxExcelAddIn::GetprogID()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("progID");
}



// ***** class wxExcelAddInBase METHODS *****

wxExcelAddIn wxExcelAddInsBase::Add(const wxString& fileName, wxXlTribool copyFile)
{
    wxExcelAddIn addIn;

    WXAUTOEXCEL_OPTIONALCPPTBOOL_TO_OPTIONALVARIANT(copyFile, copyFile);
    WXAUTOEXCEL_CALL_METHOD2_OBJECT("Add", fileName, vcopyFile, addIn);
}

// ***** class wxExcelAddInBase PROPERTIES *****

long wxExcelAddInsBase::GetCount()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("Count");
}

wxExcelAddIn wxExcelAddInsBase::GetItem(long index)
{
    wxASSERT(index > 0);

    wxExcelAddIn addIn;
    WXAUTOEXCEL_PROPERTY_OBJECT_GET1("Item", index, addIn);
}

wxExcelAddIn wxExcelAddInsBase::operator[](long index)
{
    return GetItem(index);
}

wxExcelAddIn wxExcelAddInsBase::GetItem(const wxString& name)
{
    wxASSERT( !name.empty() );

    wxExcelAddIn addIn;
    WXAUTOEXCEL_PROPERTY_OBJECT_GET1("Item", name, addIn);
}

wxExcelAddIn wxExcelAddInsBase::operator[](const wxString& name)
{
    return GetItem(name);
}


} // namespace wxAutoExcel
