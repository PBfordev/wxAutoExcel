/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////


#include "wx/wxAutoExcel_prec.h"

#include "wx/wxAutoExcelWorksheets.h"

#include "wx/wxAutoExcelApplication.h"
#include "wx/wxAutoExcelWorksheet.h"
#include "wx/wxAutoExcelSheet.h"
#include "wx/wxAutoExcelRange.h"
#include "wx/wxAutoExcelPageBreaks.h"

#include "wx/wxAutoExcel_private.h"

namespace wxAutoExcel {

// ***** class wxAutoExcelWorksheets METHODS *****

wxExcelWorksheet wxExcelWorksheets::Add(long count, const wxString& templateFile)
{
    return DoAdd(NULL, false, count, templateFile);
}

wxExcelWorksheet wxExcelWorksheets::AddAfterOrBefore(wxExcelSheet sheetAfterOrBefore, bool after, long count, const wxString& templateFile)
{
    return DoAdd(&sheetAfterOrBefore, after, count, templateFile);
}

wxExcelWorksheet wxExcelWorksheets::DoAdd(wxExcelSheet* sheetAfterOrBefore, bool after, long count, const wxString& templateFile)
{
    wxExcelWorksheet sheet;
    wxVariant vAfterOrBefore;

    if ( sheetAfterOrBefore != NULL) {
        if ( !ObjectToVariant(sheetAfterOrBefore, vAfterOrBefore) )
            return sheet;
        vAfterOrBefore.SetName(after? wxS("After") : wxS("Before"));
    }

    wxVariant vCount(count, wxS("Count"));
    WXAUTOEXCEL_OPTIONALCPPSTR_TO_OPTIONALVARIANT_NAME(Type, templateFile);


    WXAUTOEXCEL_CALL_METHOD3("Add", vAfterOrBefore, vCount, vType, "void*", sheet);
    VariantToObject(vResult, &sheet);
    return sheet;
}


void wxExcelWorksheets::FillAcrossSheets(wxExcelRange range, XlFillWith* type)
{
    wxVariant vRange;

    if ( !ObjectToVariant(&range, vRange, wxS("Range")) )
        return;

    WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT_NAME(Type, ((long*)type));
    WXAUTOEXCEL_CALL_METHOD2_RET("FillAcrossSheets", vRange, vType, "null");
}


void wxExcelWorksheets::Select(wxXlTribool replace)
{
    WXAUTOEXCEL_OPTIONALCPPTBOOL_TO_OPTIONALVARIANT(Replace, replace);
    WXAUTOEXCEL_CALL_METHOD1_RET("Select", vReplace, "null");
}

// ***** class wxAutoExcelWorksheets PROPERTIES *****


long wxExcelWorksheets::GetCount()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("Count");
}

wxExcelPageBreaks wxExcelWorksheets::GetHPageBreaks()
{
    wxExcelPageBreaks pageBreaks ;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("HPageBreaks", pageBreaks);
}

wxExcelWorksheet wxExcelWorksheets::GetItem(long index)
{
    wxASSERT(index > 0);

    wxExcelWorksheet worksheet;
    WXAUTOEXCEL_PROPERTY_OBJECT_GET1("Item", index, worksheet);
}

wxExcelWorksheet wxExcelWorksheets::GetItem(const wxString& name)
{
    wxExcelWorksheet worksheet;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET1("Item", name, worksheet);
}

wxExcelWorksheet wxExcelWorksheets::operator[](long index)
{
    return GetItem(index);
}

wxExcelWorksheet wxExcelWorksheets::operator[](const wxString& name)
{
    return GetItem(name);
}

bool wxExcelWorksheets::GetVisible()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("Visible");
}

wxExcelPageBreaks wxExcelWorksheets::GetVPageBreaks()
{
    wxExcelPageBreaks pageBreaks;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("VPageBreaks", pageBreaks);
}

} // namespace wxAutoExcel
