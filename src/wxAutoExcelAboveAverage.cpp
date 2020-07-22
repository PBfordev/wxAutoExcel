/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Modified by:
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////


#include "wx/wxAutoExcel_prec.h"

#include "wx/wxAutoExcelAboveAverage.h"

#if WXAUTOEXCEL_USE_CONDFORMAT

#include "wx/wxAutoExcel_private.h"

#include "wx/wxAutoExcelRange.h"
#include "wx/wxAutoExcelFont.h"
#include "wx/wxAutoExcelInterior.h"
#include "wx/wxAutoExcelBorders.h"

namespace wxAutoExcel {

// ***** class wxExcelAboveAverage METHODS *****

void wxExcelAboveAverage::Delete()
{
    WXAUTOEXCEL_CALL_METHOD0_RET("Delete", "null");
}

void wxExcelAboveAverage::ModifyAppliesToRange(wxExcelRange range)
{
    wxVariant vRange;

    if ( ObjectToVariant(&range, vRange, wxS("Range")) )
    {
        WXAUTOEXCEL_CALL_METHOD1_RET("ModifyAppliesToRange", vRange, "null");
    }
}

void wxExcelAboveAverage::SetFirstPriority()
{
    WXAUTOEXCEL_CALL_METHOD0_RET("SetFirstPriority", "null");
}

void wxExcelAboveAverage::SetLastPriority()
{
    WXAUTOEXCEL_CALL_METHOD0_RET("SetLastPriority", "null");
}

// ***** class wxExcelAboveAverage PROPERTIES *****

XlAboveBelow wxExcelAboveAverage::GetAboveBelow()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("AboveBelow", XlAboveBelow, xlAboveAverage);
}

void wxExcelAboveAverage::SetAboveBelow(XlAboveBelow aboveBelow)
{
    InvokePutProperty(wxS("AboveBelow"), (long)aboveBelow);
}

wxExcelRange wxExcelAboveAverage::GetAppliesTo()
{
    wxExcelRange range;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("AppliesTo", range);
}

wxExcelBorders wxExcelAboveAverage::GetBorders()
{
    wxExcelBorders borders;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("Borders", borders);
}

XlCalcFor wxExcelAboveAverage::GetCalcFor()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("CalcFor", XlCalcFor, xlAllValues);
}

void wxExcelAboveAverage::SetCalcFor(XlCalcFor calcFor)
{
    InvokePutProperty(wxS("CalcFor"), (long)calcFor);
}

wxExcelFont wxExcelAboveAverage::GetFont()
{
    wxExcelFont font;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("Font", font);
}

wxExcelInterior wxExcelAboveAverage::GetInterior()
{
    wxExcelInterior interior;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("Interior", interior);
}

wxString wxExcelAboveAverage::GetNumberFormat()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("NumberFormat");
}

void wxExcelAboveAverage::SetNumberFormat(const wxString& numberFormat)
{
    InvokePutProperty(wxS("NumberFormat"), numberFormat);
}

long wxExcelAboveAverage::GetNumStDev()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("NumStDev");
}

void wxExcelAboveAverage::SetNumStDev(long numStDev)
{
    InvokePutProperty(wxS("NumStDev"), numStDev);
}

long wxExcelAboveAverage::GetPriority()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("Priority");
}

void wxExcelAboveAverage::SetPriority(long priority)
{
    InvokePutProperty(wxS("Priority"), priority);
}

bool wxExcelAboveAverage::GetPTCondition()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("PTCondition");
}

XlPivotConditionScope wxExcelAboveAverage::GetScopeType()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("ScopeType", XlPivotConditionScope, xlSelectionScope);
}

void wxExcelAboveAverage::SetScopeType(XlPivotConditionScope scopeType)
{
    InvokePutProperty(wxS("ScopeType"), (long)scopeType);
}

bool wxExcelAboveAverage::GetStopIfTrue()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("StopIfTrue");
}

void wxExcelAboveAverage::SetStopIfTrue(bool stopIfTrue)
{
    InvokePutProperty(wxS("StopIfTrue"), stopIfTrue);
}

XlFormatConditionType wxExcelAboveAverage::GetType()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("Type", XlFormatConditionType, xlCellValue);
}

void wxExcelAboveAverage::SetType(XlFormatConditionType type)
{
    InvokePutProperty(wxS("Type"), (long)type);
}


} // namespace wxAutoExcel


#endif // WXAUTOEXCEL_USE_CONDFORMAT