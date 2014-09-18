/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Modified by:
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////


#include "wx/wxAutoExcel_prec.h"

#include "wx/wxAutoExcelTop10.h"

#if WXAUTOEXCEL_USE_CONDFORMAT

#include "wx/wxAutoExcelPrivate.h"

#include "wx/wxAutoExcelRange.h"
#include "wx/wxAutoExcelFont.h"
#include "wx/wxAutoExcelInterior.h"
#include "wx/wxAutoExcelBorders.h"

namespace wxAutoExcel {

// ***** class wxExcelTop10 METHODS *****

void wxExcelTop10::Delete()
{
    WXAUTOEXCEL_CALL_METHOD0_RET("Delete", "null");
}

void wxExcelTop10::ModifyAppliesToRange(wxExcelRange range)
{    
    wxVariant vRange;

    if ( ObjectToVariant(&range, vRange, wxS("Range")) )
    {
        WXAUTOEXCEL_CALL_METHOD1_RET("ModifyAppliesToRange", vRange, "null");
    }
}

void wxExcelTop10::SetFirstPriority()
{
    WXAUTOEXCEL_CALL_METHOD0_RET("SetFirstPriority", "null");
}

void wxExcelTop10::SetLastPriority()
{
    WXAUTOEXCEL_CALL_METHOD0_RET("SetLastPriority", "null");
}

// ***** class wxExcelTop10 PROPERTIES *****

wxExcelRange wxExcelTop10::GetAppliesTo()
{
    wxExcelRange range;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("AppliesTo", range);
}

wxExcelBorders wxExcelTop10::GetBorders()
{
    wxExcelBorders borders;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("Borders", borders);
}

XlCalcFor wxExcelTop10::GetCalcFor()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("CalcFor", XlCalcFor, xlAllValues);
}

void wxExcelTop10::SetCalcFor(XlCalcFor calcFor)
{
    InvokePutProperty(wxS("CalcFor"), (long)calcFor);
}

wxExcelFont wxExcelTop10::GetFont()
{
    wxExcelFont font;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("Font", font);
}

wxExcelInterior wxExcelTop10::GetInterior()
{
    wxExcelInterior interior;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("Interior", interior);
}

wxString wxExcelTop10::GetNumberFormat()
{    
    WXAUTOEXCEL_PROPERTY_STRING_GET0("NumberFormat");
}

void wxExcelTop10::SetNumberFormat(const wxString& numberFormat)
{
   InvokePutProperty(wxS("NumberFormat"), numberFormat);
}

bool wxExcelTop10::GetPercent()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("Percent");
}

void wxExcelTop10::SetPercent(bool percent)
{
    InvokePutProperty(wxS("Percent"), percent);
}

long wxExcelTop10::GetPriority()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("Priority");
}

void wxExcelTop10::SetPriority(long priority)
{
    InvokePutProperty(wxS("Priority"), priority);
}

bool wxExcelTop10::GetPTCondition()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("PTCondition");
}

long wxExcelTop10::GetRank()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("Rank");
}

void wxExcelTop10::SetRank(long rank)
{
    InvokePutProperty(wxS("Rank"), rank);
}

XlPivotConditionScope wxExcelTop10::GetScopeType()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("ScopeType", XlPivotConditionScope, xlSelectionScope);
}

void wxExcelTop10::SetScopeType(XlPivotConditionScope scopeType)
{
    InvokePutProperty(wxS("ScopeType"), (long)scopeType);
}

bool wxExcelTop10::GetStopIfTrue()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("StopIfTrue");
}

void wxExcelTop10::SetStopIfTrue(bool stopIfTrue)
{
    InvokePutProperty(wxS("StopIfTrue"), stopIfTrue);
}

XlTopBottom wxExcelTop10::GetTopBottom()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("TopBottom", XlTopBottom, xlTop10Bottom);
}

void wxExcelTop10::SetTopBottom(XlTopBottom topBottom)
{
    InvokePutProperty(wxS("TopBottom"), (long)topBottom);
}

XlFormatConditionType wxExcelTop10::GetType()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("Type", XlFormatConditionType, xlCellValue);
}


} // namespace wxAutoExcel


#endif // WXAUTOEXCEL_USE_CONDFORMAT