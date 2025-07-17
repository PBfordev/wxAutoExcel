/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// License:     MIT license
/////////////////////////////////////////////////////////////////////////////


#include "wx/wxAutoExcel_prec.h"

#include "wx/wxAutoExcelRecentFiles.h"

#include "wx/wxAutoExcelWorkbook.h"
#include "wx/wxAutoExcel_private.h"

namespace wxAutoExcel {

// ***** class wxExcelRecentFile METHODS *****

void wxExcelRecentFile::Delete()
{
    WXAUTOEXCEL_CALL_METHOD0_RET("Delete", "null");
}

wxExcelWorkbook wxExcelRecentFile::Open()
{
    wxExcelWorkbook book;

    WXAUTOEXCEL_CALL_METHOD0_OBJECT("Open", book);
}

// ***** class wxExcelRecentFile PROPERTIES *****

long wxExcelRecentFile::GetIndex()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("Index");
}

wxString wxExcelRecentFile::GetName()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("Name");
}

wxString wxExcelRecentFile::GetPath()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("Path");
}


// ***** class wxExcelRecentFiles METHODS *****

wxExcelRecentFile wxExcelRecentFiles::Add(const wxString& name)
{
    wxASSERT( !name.empty() );

    wxExcelRecentFile file;

    WXAUTOEXCEL_CALL_METHOD1_OBJECT("Add", name, file);
}

// ***** class wxExcelRecentFiles PROPERTIES *****

long wxExcelRecentFiles::GetCount()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("Count");
}

wxExcelRecentFile wxExcelRecentFiles::GetItem(long index)
{
    wxASSERT( index > 0 );

    wxExcelRecentFile item;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET1("Item", index, item);
}

wxExcelRecentFile wxExcelRecentFiles::operator[](long index)
{
    return GetItem(index);
}

long wxExcelRecentFiles::GetMaximum()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("Maximum");
}

void wxExcelRecentFiles::SetMaximum(long maximum)
{
    InvokePutProperty(wxS("Maximum"), maximum);
}








} // namespace wxAutoExcel

