/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Modified by:
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////


#include "wx/wxAutoExcel_prec.h"

#include "wx/wxAutoExcelIconSetCondition.h"

#if WXAUTOEXCEL_USE_CONDFORMAT

#include "wx/wxAutoExcelPrivate.h"

#include "wx/wxAutoExcelRange.h"
#include "wx/wxAutoExcelIconCriteria.h"
#include "wx/wxAutoExcelIconSets.h"

namespace wxAutoExcel {

// ***** class wxExcelIconSetCondition METHODS *****

void wxExcelIconSetCondition::Delete()
{
    WXAUTOEXCEL_CALL_METHOD0_RET("Delete", "null");
}

void wxExcelIconSetCondition::ModifyAppliesToRange(wxExcelRange range)
{    
    wxVariant vRange;

    if ( ObjectToVariant(&range, vRange, wxS("Range")) )
    {
        WXAUTOEXCEL_CALL_METHOD1_RET("ModifyAppliesToRange", vRange, "null");
    }
}

void wxExcelIconSetCondition::SetFirstPriority()
{
    WXAUTOEXCEL_CALL_METHOD0_RET("SetFirstPriority", "null");
}

void wxExcelIconSetCondition::SetLastPriority()
{
    WXAUTOEXCEL_CALL_METHOD0_RET("SetLastPriority", "null");
}

// ***** class wxExcelIconSetCondition PROPERTIES *****

wxExcelRange wxExcelIconSetCondition::GetAppliesTo()
{
    wxExcelRange range;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("AppliesTo", range);
}

wxString wxExcelIconSetCondition::GetFormula()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("Formula");
}

void wxExcelIconSetCondition::SetFormula(const wxString& formula)
{
    InvokePutProperty(wxS("Formula"), formula);
}

wxExcelIconCriteria wxExcelIconSetCondition::GetIconCriteria()
{
    wxExcelIconCriteria iconCriteria;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("IconCriteria", iconCriteria);
}

wxExcelIconSets wxExcelIconSetCondition::GetIconSet()
{
    wxExcelIconSets iconSets;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("IconSet", iconSets);
}

void wxExcelIconSetCondition::SetIconSet(const wxExcelIconSets& iconSet)
{
    wxVariant vIconSet;

    if ( ObjectToVariant(&iconSet, vIconSet, wxS("IconSet")) )
    {
        InvokePutProperty("IconSet", vIconSet);
    }
}

bool wxExcelIconSetCondition::GetPercentileValues()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("PercentileValues");
}

void wxExcelIconSetCondition::SetPercentileValues(bool percentileValues)
{
    InvokePutProperty(wxS("PercentileValues"), percentileValues);
}

long wxExcelIconSetCondition::GetPriority()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("Priority");
}

void wxExcelIconSetCondition::SetPriority(long priority)
{
    InvokePutProperty(wxS("Priority"), priority);
}

bool wxExcelIconSetCondition::GetPTCondition()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("PTCondition");
}

bool wxExcelIconSetCondition::GetReverseOrder()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("ReverseOrder");
}

void wxExcelIconSetCondition::SetReverseOrder(bool reverseOrder)
{
    InvokePutProperty(wxS("ReverseOrder"), reverseOrder);
}

XlPivotConditionScope wxExcelIconSetCondition::GetScopeType()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("ScopeType", XlPivotConditionScope, xlSelectionScope);
}

void wxExcelIconSetCondition::SetScopeType(XlPivotConditionScope scopeType)
{
    InvokePutProperty(wxS("ScopeType"), (long)scopeType);
}

bool wxExcelIconSetCondition::GetShowIconOnly()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("ShowIconOnly");
}

void wxExcelIconSetCondition::SetShowIconOnly(bool showIconOnly)
{
    InvokePutProperty(wxS("ShowIconOnly"), showIconOnly);
}

bool wxExcelIconSetCondition::GetStopIfTrue()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("StopIfTrue");
}

void wxExcelIconSetCondition::SetStopIfTrue(bool stopIfTrue)
{
    InvokePutProperty(wxS("StopIfTrue"), stopIfTrue);
}

XlFormatConditionType  wxExcelIconSetCondition::GetType()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("Type", XlFormatConditionType, xlCellValue);
}



} // namespace wxAutoExcel


#endif // WXAUTOEXCEL_USE_CONDFORMAT