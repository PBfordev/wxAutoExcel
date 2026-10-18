/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// License:     MIT license
/////////////////////////////////////////////////////////////////////////////


#include "wx/wxAutoExcel_prec.h"

#include "wx/wxAutoExcelIconCriteria.h"

#if WXAUTOEXCEL_USE_CONDFORMAT

#include "wx/wxAutoExcelIcon.h"

#include "wx/wxAutoExcel_private.h"

namespace wxAutoExcel {

// ***** class wxExcelIconCriterion PROPERTIES *****

XlIcon wxExcelIconCriterion::GetIcon()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("Icon", XlIcon, xlIconNoCellIcon);
}

void wxExcelIconCriterion::SetIcon(XlIcon icon)
{
    InvokePutProperty(wxS("Icon"), (long)icon);
}

long wxExcelIconCriterion::GetIndex()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("Index");
}

XlFormatConditionOperator wxExcelIconCriterion::GetOperator()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("Operator", XlFormatConditionOperator, xlBetween);
}

void wxExcelIconCriterion::SetOperator(XlFormatConditionOperator conditionOperator)
{
    InvokePutProperty(wxS("Operator"), (long)conditionOperator);
}

XlConditionValueTypes  wxExcelIconCriterion::GetType()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("Type", XlConditionValueTypes, xlConditionValueNone);
}

wxVariant wxExcelIconCriterion::GetValue()
{
    wxVariant vResult;

    InvokeGetProperty(wxS("Value"), vResult);
    return vResult;
}

void wxExcelIconCriterion::SetValue(const wxVariant& value)
{
    InvokePutProperty(wxS("Value"), value);
}

// ***** class wxExcelIconCriteria PROPERTIES *****

wxExcelIconCriterion wxExcelIconCriteria::GetItem(long index)
{
    wxASSERT( index > 0 );

    wxExcelIconCriterion criterion;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET1("Item", index, criterion);
}

wxExcelIconCriterion wxExcelIconCriteria::operator[](long index)
{
    return GetItem(index);
}


} // namespace wxAutoExcel

#endif // WXAUTOEXCEL_USE_CONDFORMAT
