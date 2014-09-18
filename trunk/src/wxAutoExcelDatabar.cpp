/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Modified by:
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////


#include "wx/wxAutoExcel_prec.h"

#include "wx/wxAutoExcelDatabar.h"

#if WXAUTOEXCEL_USE_CONDFORMAT

#include "wx/wxAutoExcelRange.h"
#include "wx/wxAutoExcelFormatColor.h"
#include "wx/wxAutoExcelConditionValue.h"

#include "wx/wxAutoExcelPrivate.h"

namespace wxAutoExcel {


// ***** class wxExcelDatabar METHODS *****

void wxExcelDatabar::Delete()
{
    WXAUTOEXCEL_CALL_METHOD0_RET("Delete", "null");
}

void wxExcelDatabar::ModifyAppliesToRange(wxExcelRange range)
{    
    wxVariant vRange;

    if ( ObjectToVariant(&range, vRange, wxS("Range")) )
    {
        WXAUTOEXCEL_CALL_METHOD1_RET("ModifyAppliesToRange", vRange, "null");
    }
}

void wxExcelDatabar::SetFirstPriority()
{
    WXAUTOEXCEL_CALL_METHOD0_RET("SetFirstPriority", "null");
}

void wxExcelDatabar::SetLastPriority()
{
    WXAUTOEXCEL_CALL_METHOD0_RET("SetLastPriority", "null");
}

// ***** class wxExcelDatabar PROPERTIES *****


wxExcelRange wxExcelDatabar::GetAppliesTo()
{
    wxExcelRange range;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("AppliesTo", range);
}

wxExcelFormatColor wxExcelDatabar::GetAxisColor()
{
    wxExcelFormatColor formatColor;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("AxisColor", formatColor);
}


XlDataDataBarAxisPosition wxExcelDatabar::GetAxisPosition()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("AxisPosition", XlDataDataBarAxisPosition, xlDataBarAxisAutomatic);
}

void wxExcelDatabar::SetAxisPosition(XlDataDataBarAxisPosition axisPosition)
{
    InvokePutProperty(wxS("AxisPosition"), (long)axisPosition);
}


wxExcelFormatColor wxExcelDatabar::GetBarColor()
{
    wxExcelFormatColor formatColor;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("BarColor", formatColor);
}


XlDataBarFillType wxExcelDatabar::GetBarFillType()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("BarFillType", XlDataBarFillType, xlDataBarFillGradient);
}

void wxExcelDatabar::SetBarFillType(XlDataBarFillType barFillType)
{
    InvokePutProperty(wxS("BarFillType"), (long)barFillType);
}

wxString wxExcelDatabar::GetFormula()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("Formula");
}

void wxExcelDatabar::SetFormula(const wxString& formula)
{
    InvokePutProperty(wxS("Formula"), formula);
}

wxExcelConditionValue wxExcelDatabar::GetMaxPoint()
{
    wxExcelConditionValue conditionValue;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("MaxPoint", conditionValue);
}

wxExcelConditionValue wxExcelDatabar::GetMinPoint()
{
    wxExcelConditionValue conditionValue;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("MinPoint", conditionValue);
}


long wxExcelDatabar::GetPercentMax()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("PercentMax");
}

void wxExcelDatabar::SetPercentMax(long percentMax)
{
    InvokePutProperty(wxS("PercentMax"), percentMax);
}

long wxExcelDatabar::GetPercentMin()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("PercentMin");
}

void wxExcelDatabar::SetPercentMin(long percentMin)
{
    InvokePutProperty(wxS("PercentMin"), percentMin);
}

long wxExcelDatabar::GetPriority()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("Priority");
}

void wxExcelDatabar::SetPriority(long priority)
{
    InvokePutProperty(wxS("Priority"), priority);
}

bool wxExcelDatabar::GetPTCondition()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("PTCondition");
}

XlPivotConditionScope wxExcelDatabar::GetScopeType()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("ScopeType", XlPivotConditionScope, xlSelectionScope);
}

void wxExcelDatabar::SetScopeType(XlPivotConditionScope scopeType)
{
    InvokePutProperty(wxS("ScopeType"), (long)scopeType);
}

bool wxExcelDatabar::GetShowValue()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("ShowValue");
}

void wxExcelDatabar::SetShowValue(bool showValue)
{
    InvokePutProperty(wxS("ShowValue"), showValue);
}

bool wxExcelDatabar::GetStopIfTrue()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("StopIfTrue");
}

void wxExcelDatabar::SetStopIfTrue(bool stopIfTrue)
{
    InvokePutProperty(wxS("StopIfTrue"), stopIfTrue);
}

XlFormatConditionType wxExcelDatabar::GetType()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("Type", XlFormatConditionType, xlCellValue);
}



} // namespace wxAutoExcel

#endif // WXAUTOEXCEL_USE_CONDFORMAT