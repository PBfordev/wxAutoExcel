/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Modified by:
// RCS-ID:      $Id: $
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////


#include "wx/wxAutoExcel_prec.h"

#include "wx/wxAutoExcelControlFormat.h"

#if WXAUTOEXCEL_USE_SHAPES

#include "wx/wxAutoExcelPrivate.h"

namespace wxAutoExcel {

// ***** class wxExcelControlFormat METHODS *****

void wxExcelControlFormat::AddItem(const wxString& text, long* index)
{
    WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT(Index, index);
    WXAUTOEXCEL_CALL_METHOD2_RET("AddItem", text, vIndex, "null");
}

wxString wxExcelControlFormat::List(long index)
{
    WXAUTOEXCEL_CALL_METHOD1_STRING("List", index);    
}

wxArrayString wxExcelControlFormat::List()
{
    wxArrayString as;
    wxVariant vResult;

    if ( InvokeMethod(wxS("List"), vResult) )
    {
        wxString type = vResult.GetType();
        if ( type == wxS("arrstring") )
            return vResult.GetArrayString();
        if ( type == wxS("list") )
        {
            as.reserve(vResult.GetCount());
            for (size_t i = 0; i < vResult.GetCount(); i++)
                as.push_back(vResult[i].GetString());
        }
    }

    return as;
}

void wxExcelControlFormat::RemoveAllItems()
{
    WXAUTOEXCEL_CALL_METHOD0_RET("RemoveAllItems", "null");
}

void wxExcelControlFormat::RemoveItem(long index, long* count)
{
    WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT(Count, count);
    WXAUTOEXCEL_CALL_METHOD2_RET("RemoveItem", index, vCount, "null");    
}

// ***** class wxExcelControlFormat PROPERTIES *****


long wxExcelControlFormat::GetDropDownLines()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("DropDownLines");
}

void wxExcelControlFormat::SetDropDownLines(long dropDownLines)
{
    InvokePutProperty(wxS("DropDownLines"), dropDownLines);
}

bool wxExcelControlFormat::GetEnabled()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("Enabled");
}

void wxExcelControlFormat::SetEnabled(bool enabled)
{
    InvokePutProperty(wxS("Enabled"), enabled);
}

long wxExcelControlFormat::GetLargeChange()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("LargeChange");
}

void wxExcelControlFormat::SetLargeChange(long largeChange)
{
    InvokePutProperty(wxS("LargeChange"), largeChange);
}

wxString wxExcelControlFormat::GetLinkedCell()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("LinkedCell");
}

void wxExcelControlFormat::SetLinkedCell(const wxString& linkedCell)
{
    InvokePutProperty(wxS("LinkedCell"), linkedCell);
}

long wxExcelControlFormat::GetListCount()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("ListCount");
}

wxString wxExcelControlFormat::GetListFillRange()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("ListFillRange");
}

void wxExcelControlFormat::SetListFillRange(const wxString& listFillRange)
{
    InvokePutProperty(wxS("ListFillRange"), listFillRange);
}

long wxExcelControlFormat::GetListIndex()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("ListIndex");
}

void wxExcelControlFormat::SetListIndex(long listIndex)
{
    InvokePutProperty(wxS("ListIndex"), listIndex);
}

bool wxExcelControlFormat::GetLockedText()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("LockedText");
}

void wxExcelControlFormat::SetLockedText(bool lockedText)
{
    InvokePutProperty(wxS("LockedText"), lockedText);
}

long wxExcelControlFormat::GetMax()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("Max");
}

void wxExcelControlFormat::SetMax(long max)
{
    InvokePutProperty(wxS("Max"), max);
}

long wxExcelControlFormat::GetMin()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("Min");
}

void wxExcelControlFormat::SetMin(long min)
{
    InvokePutProperty(wxS("Min"), min);
}

long wxExcelControlFormat::GetMultiSelect()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("MultiSelect");
}

void wxExcelControlFormat::SetMultiSelect(long multiSelect)
{
    InvokePutProperty(wxS("MultiSelect"), multiSelect);
}

bool wxExcelControlFormat::GetPrintObject()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("PrintObject");
}

void wxExcelControlFormat::SetPrintObject(bool printObject)
{
    InvokePutProperty(wxS("PrintObject"), printObject);
}

long wxExcelControlFormat::GetSmallChange()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("SmallChange");
}

void wxExcelControlFormat::SetSmallChange(long smallChange)
{
    InvokePutProperty(wxS("SmallChange"), smallChange);
}

long wxExcelControlFormat::GetValue()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("Value");
}

void wxExcelControlFormat::SetValue(long value)
{
    InvokePutProperty(wxS("Value"), value);
}

} // namespace wxAutoExcel


#endif // #if WXAUTOEXCEL_USE_SHAPES
