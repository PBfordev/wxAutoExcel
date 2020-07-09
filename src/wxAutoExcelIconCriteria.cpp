/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Modified by:
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////


#include "wx/wxAutoExcel_prec.h"

#include "wx/wxAutoExcelIconCriteria.h"

#if WXAUTOEXCEL_USE_CONDFORMAT

#include "wx/wxAutoExcelIcon.h"

#include "wx/wxAutoExcel_private.h"

namespace wxAutoExcel {

// ***** class wxExcelIconCriterion PROPERTIES *****

wxExcelIcon wxExcelIconCriterion::GetIcon()
{
    wxExcelIcon icon;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("Icon", icon);
}

void wxExcelIconCriterion::SetIcon(const wxExcelIcon& icon)
{
    wxVariant vIcon;
    if ( ObjectToVariant(&icon, vIcon, wxS("Icon")) )
    {
        InvokePutProperty(wxS("Icon"), vIcon);
    }
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
