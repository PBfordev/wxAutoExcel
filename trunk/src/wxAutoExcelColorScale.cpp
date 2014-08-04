/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Modified by:
// RCS-ID:      $Id: $
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////


#include "wx/wxAutoExcel_prec.h"

#include "wx/wxAutoExcelColorScale.h"

#if WXAUTOEXCEL_USE_CONDFORMAT

#include "wx/wxAutoExcelPrivate.h"

#include "wx/wxAutoExcelRange.h"
#include "wx/wxAutoExcelColorScaleCriteria.h"

namespace wxAutoExcel {

// ***** class wxExcelColorScale METHODS *****

void wxExcelColorScale::Delete()
{
    WXAUTOEXCEL_CALL_METHOD0_RET("Delete", "null");
}


void wxExcelColorScale::ModifyAppliesToRange(wxExcelRange range)
{    
    wxVariant vRange;

    if ( ObjectToVariant(&range, vRange, wxS("Range")) )
    {
        WXAUTOEXCEL_CALL_METHOD1_RET("ModifyAppliesToRange", vRange, "null");
    }
}

void wxExcelColorScale::SetFirstPriority()
{
    WXAUTOEXCEL_CALL_METHOD0_RET("SetFirstPriority", "null");
}

void wxExcelColorScale::SetLastPriority()
{
    WXAUTOEXCEL_CALL_METHOD0_RET("SetLastPriority", "null");
}

// ***** class wxExcelColorScale PROPERTIES *****

wxExcelRange wxExcelColorScale::GetAppliesTo()
{
    wxExcelRange range;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("AppliesTo", range);
}

wxExcelColorScaleCriteria wxExcelColorScale::GetColorScaleCriteria()
{
    wxExcelColorScaleCriteria colorScaleCriteria;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("ColorScaleCriteria", colorScaleCriteria);
}

wxString wxExcelColorScale::GetFormula()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("Formula");
}

void wxExcelColorScale::SetFormula(const wxString& formula)
{
    InvokePutProperty(wxS("Formula"), formula);
}

long wxExcelColorScale::GetPriority()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("Priority");
}

void wxExcelColorScale::SetPriority(long priority)
{
    InvokePutProperty(wxS("Priority"), priority);
}

bool wxExcelColorScale::GetPTCondition()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("PTCondition");
}

XlPivotConditionScope wxExcelColorScale::GetScopeType()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("ScopeType", XlPivotConditionScope, xlSelectionScope);
}

bool wxExcelColorScale::GetStopIfTrue()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("StopIfTrue");
}

XlFormatConditionType wxExcelColorScale::GetType()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("Type", XlFormatConditionType, xlCellValue);
}


} // namespace wxAutoExcel

#endif // WXAUTOEXCEL_USE_CONDFORMAT