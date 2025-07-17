/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2020 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////

#include "wx/wxAutoExcel_prec.h"

#include "wx/wxAutoExcelTableObject.h"

#include "wx/wxAutoExcelListObject.h"
#include "wx/wxAutoExcelRange.h"

#include "wx/wxAutoExcel_private.h"

namespace wxAutoExcel {

// ***** class wxExcelTableObject METHODS *****

void wxExcelTableObject::Delete()
{
    WXAUTOEXCEL_CALL_METHOD0_RET("Delete", "null");
}

bool wxExcelTableObject::Refresh()
{
    WXAUTOEXCEL_CALL_METHOD0_BOOL("Refresh");
}

// ***** class wxExcelTableObject PROPERTIES *****

bool wxExcelTableObject::GetAdjustColumnWidth()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("AdjustColumnWidth");
}

void wxExcelTableObject::SetAdjustColumnWidth(bool adjustColumnWidth)
{
    InvokePutProperty(wxS("AdjustColumnWidth"), adjustColumnWidth);
}

wxExcelRange wxExcelTableObject::GetDestination()
{
    wxExcelRange range;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("Destination", range);
}

bool wxExcelTableObject::GetEnableEditing()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("EnableEditing");
}

void wxExcelTableObject::SetEnableEditing(bool enableEditing)
{
    InvokePutProperty(wxS("EnableEditing"), enableEditing);
}

bool wxExcelTableObject::GetEnableRefresh()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("EnableRefresh");
}

void wxExcelTableObject::SetEnableRefresh(bool enableRefresh)
{
    InvokePutProperty(wxS("EnableRefresh"), enableRefresh);
}

bool wxExcelTableObject::GetFetchedRowOverflow()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("FetchedRowOverflow");
}

wxExcelListObject wxExcelTableObject::GetListObject()
{
    wxExcelListObject listObject;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("ListObject", listObject);
}

bool wxExcelTableObject::GetPreserveColumnInfo()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("PreserveColumnInfo");
}

void wxExcelTableObject::SetPreserveColumnInfo(bool preserveColumnInfo)
{
    InvokePutProperty(wxS("PreserveColumnInfo"), preserveColumnInfo);
}

bool wxExcelTableObject::GetPreserveFormatting()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("PreserveFormatting");
}

void wxExcelTableObject::SetPreserveFormatting(bool preserveFormatting)
{
    InvokePutProperty(wxS("PreserveFormatting"), preserveFormatting);
}

XlCellInsertionMode wxExcelTableObject::GetRefreshStyle()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("RefreshStyle", XlCellInsertionMode, xlOverwriteCells);
}

void wxExcelTableObject::SetRefreshStyle(XlCellInsertionMode refreshStyle)
{
    InvokePutProperty(wxS("RefreshStyle"), (long)refreshStyle);
}

wxExcelRange wxExcelTableObject::GetResultRange()
{
    wxExcelRange range;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("ResultRange", range);
}

bool wxExcelTableObject::GetRowNumbers()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("RowNumbers");
}

void wxExcelTableObject::SetRowNumbers(bool rowNumbers)
{
    InvokePutProperty(wxS("RowNumbers"), rowNumbers);
}

} // namespace wxAutoExcel
