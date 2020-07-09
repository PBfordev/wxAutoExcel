/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Modified by:
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////


#include "wx/wxAutoExcel_prec.h"

#include "wx/wxAutoExcelUniqueValues.h"

#if WXAUTOEXCEL_USE_CONDFORMAT

#include "wx/wxAutoExcel_private.h"

#include "wx/wxAutoExcelRange.h"
#include "wx/wxAutoExcelFont.h"
#include "wx/wxAutoExcelInterior.h"
#include "wx/wxAutoExcelBorders.h"


namespace wxAutoExcel {

// ***** class wxExcelUniqueValues METHODS *****

void wxExcelUniqueValues::Delete()
{
    WXAUTOEXCEL_CALL_METHOD0_RET("Delete", "null");
}

void wxExcelUniqueValues::ModifyAppliesToRange(wxExcelRange range)
{
    wxVariant vRange;

    if ( ObjectToVariant(&range, vRange, wxS("Range")) )
    {
        WXAUTOEXCEL_CALL_METHOD1_RET("ModifyAppliesToRange", vRange, "null");
    }
}

void wxExcelUniqueValues::SetFirstPriority()
{
    WXAUTOEXCEL_CALL_METHOD0_RET("SetFirstPriority", "null");
}

void wxExcelUniqueValues::SetLastPriority()
{
    WXAUTOEXCEL_CALL_METHOD0_RET("SetLastPriority", "null");
}

// ***** class wxExcelUniqueValues PROPERTIES *****

wxExcelRange wxExcelUniqueValues::GetAppliesTo()
{
    wxExcelRange range;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("AppliesTo", range);
}

wxExcelBorders wxExcelUniqueValues::GetBorders()
{
    wxExcelBorders borders;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("Borders", borders);
}

XlDupeUnique wxExcelUniqueValues::GetDupeUnique()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("DupeUnique", XlDupeUnique, xlUnique);
}

void wxExcelUniqueValues::SetDupeUnique(XlDupeUnique  dupeUnique)
{
    InvokePutProperty(wxS("DupeUnique"), (long)dupeUnique);
}

wxExcelFont wxExcelUniqueValues::GetFont()
{
    wxExcelFont font;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("Font", font);
}

wxExcelInterior wxExcelUniqueValues::GetInterior()
{
    wxExcelInterior interior;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("Interior", interior);
}

wxString wxExcelUniqueValues::GetNumberFormat()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("NumberFormat");
}

void wxExcelUniqueValues::SetNumberFormat(const wxString& numberFormat)
{
   InvokePutProperty(wxS("NumberFormat"), numberFormat);
}

long wxExcelUniqueValues::GetPriority()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("Priority");
}

void wxExcelUniqueValues::SetPriority(long priority)
{
    InvokePutProperty(wxS("Priority"), priority);
}

bool wxExcelUniqueValues::GetPTCondition()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("PTCondition");
}

XlPivotConditionScope wxExcelUniqueValues::GetScopeType()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("ScopeType", XlPivotConditionScope, xlSelectionScope);
}

void wxExcelUniqueValues::SetScopeType(XlPivotConditionScope scopeType)
{
    InvokePutProperty(wxS("ScopeType"), (long)scopeType);
}

bool wxExcelUniqueValues::GetStopIfTrue()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("StopIfTrue");
}

void wxExcelUniqueValues::SetStopIfTrue(bool stopIfTrue)
{
    InvokePutProperty(wxS("StopIfTrue"), stopIfTrue);
}

XlFormatConditionType wxExcelUniqueValues::GetType()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("Type", XlFormatConditionType, xlCellValue);
}



} // namespace wxAutoExcel


#endif // WXAUTOEXCEL_USE_CONDFORMAT