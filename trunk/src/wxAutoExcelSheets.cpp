/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Modified by:
// RCS-ID:      $Id: $
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////


#include "wx/wxAutoExcel_prec.h"

#include "wx/wxAutoExcelSheets.h"
#include "wx/wxAutoExcelSheet.h"

#include "wx/wxAutoExcelPrivate.h"

namespace wxAutoExcel {

wxExcelSheet wxExcelSheets::Add(long count, XlSheetType type)
{
    return DoAdd(NULL, false, count, type);
}

wxExcelSheet wxExcelSheets::AddAfterOrBefore(wxExcelSheet sheetAfterOrBefore, bool after, long count, XlSheetType type)
{
    return DoAdd(&sheetAfterOrBefore, after, count, type);
}

wxExcelSheet wxExcelSheets::DoAdd(wxExcelSheet* sheetAfterOrBefore, bool after, long count, XlSheetType type)
{
    wxExcelSheet sheet;
    wxVariant vAfterOrBefore;

    if (sheetAfterOrBefore != NULL) {
        if ( !ObjectToVariant(sheetAfterOrBefore, vAfterOrBefore) )
            return sheet;
        vAfterOrBefore.SetName(after? wxS("After") : wxS("Before"));
    }

    wxVariant vCount(count, wxS("Count"));
    wxVariant vType(type, wxS("Type"));    
    
    WXAUTOEXCEL_CALL_METHOD3("Add", vAfterOrBefore, vCount, vType, "void*", sheet);   
    VariantToObject(vResult, &sheet);    
    return sheet;
}

long wxExcelSheets::GetCount()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("Count");
}

wxExcelSheet wxExcelSheets::GetItem(long index)
{
    wxASSERT(index > 0);

    wxExcelSheet worksheet;
    WXAUTOEXCEL_PROPERTY_OBJECT_GET1("Item", index, worksheet);
}

wxExcelSheet wxExcelSheets::GetItem(const wxString& name)
{
    wxExcelSheet worksheet;
    WXAUTOEXCEL_PROPERTY_OBJECT_GET1("Item", name, worksheet);
}

wxExcelSheet wxExcelSheets::operator[](long index)
{
    return GetItem(index);
}

wxExcelSheet wxExcelSheets::operator[](const wxString& name)
{
    return GetItem(name);
}

} // namespace wxAutoExcel
