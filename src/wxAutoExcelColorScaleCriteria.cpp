/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Modified by:
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////


#include "wx/wxAutoExcel_prec.h"

#include "wx/wxAutoExcelColorScaleCriteria.h"

#if WXAUTOEXCEL_USE_CONDFORMAT

#include "wx/wxAutoExcel_private.h"

#include "wx/wxAutoExcelFormatColor.h"

namespace wxAutoExcel {

// ***** class wxExcelColorScaleCriterion PROPERTIES *****

wxExcelFormatColor wxExcelColorScaleCriterion::GetFormatColor()
{
    wxExcelFormatColor formatColor;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("FormatColor", formatColor);
}

long wxExcelColorScaleCriterion::GetIndex()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("Index");
}

XlConditionValueTypes wxExcelColorScaleCriterion::GetType()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("Type", XlConditionValueTypes, xlConditionValueNone);
}

void wxExcelColorScaleCriterion::SetType(XlConditionValueTypes type)
{
    InvokePutProperty(wxS("Type"), (long)type);
}

wxVariant wxExcelColorScaleCriterion::GetValue()
{
    wxVariant vResult;

    InvokeGetProperty(wxS("Value"), vResult);
    return vResult;
}

void wxExcelColorScaleCriterion::SetValue(const wxVariant& value)
{
    InvokePutProperty(wxS("Value"), value);
}


// ***** class wxExcelColorScaleCriteria PROPERTIES *****

long wxExcelColorScaleCriteria::GetCount()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("Count");
}

wxExcelColorScaleCriterion wxExcelColorScaleCriteria::GetItem(long index)
{
    wxASSERT( index > 0 );

    wxExcelColorScaleCriterion colorScaleCriterion;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET1("Item", index, colorScaleCriterion);
}

wxExcelColorScaleCriterion wxExcelColorScaleCriteria::operator[](long index)
{
    return GetItem(index);
}


} // namespace wxAutoExcel

#endif // WXAUTOEXCEL_USE_CONDFORMAT
