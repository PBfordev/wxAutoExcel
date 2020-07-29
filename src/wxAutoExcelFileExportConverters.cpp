/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Modified by:
// Copyright:   (c) 2020 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////

#include "wx/wxAutoExcel_prec.h"

#include "wx/wxAutoExcelFileExportConverters.h"

#include "wx/wxAutoExcel_private.h"

namespace wxAutoExcel {

// ***** class wxExcelFileExportConverter PROPERTIES *****

wxString wxExcelFileExportConverter::GetDescription()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("Description");
}

wxString wxExcelFileExportConverter::GetExtensions()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("Extensions");
}

long wxExcelFileExportConverter::GetFileFormat()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("FileFormat");
}

// ***** class wxExcelFileExportConverters PROPERTIES *****

wxExcelFileExportConverter wxExcelFileExportConverters::GetItem(long index)
{
    wxASSERT( index > 0 );

    wxExcelFileExportConverter object;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET1("Item", index, object);
}

wxExcelFileExportConverter wxExcelFileExportConverters::operator[](long index)
{
    return GetItem(index);
}


long wxExcelFileExportConverters::GetCount()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("Count");
}

} // namespace wxAutoExcel
