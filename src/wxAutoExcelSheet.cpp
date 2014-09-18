/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Modified by:
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////


#include "wx/wxAutoExcel_prec.h"

#include "wx/wxAutoExcelSheet.h"

#include "wx/wxAutoExcelChart.h"
#include "wx/wxAutoExcelWorkSheet.h"

#include "wx/wxAutoExcelPrivate.h"

namespace wxAutoExcel {

wxString wxExcelSheet::GetName()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("Name");    
}

XlSheetType wxExcelSheet::GetType()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("Type", XlSheetType, xlWorksheet);
}

bool wxExcelSheet::IsWorksheet()
{
    return IsOk_() && GetAutomationObjectName_(true).IsSameAs(wxS("Worksheet"));    
}

bool wxExcelSheet::IsChart()
{
    return IsOk_() && GetAutomationObjectName_(true).IsSameAs(wxS("Chart"));
}

wxExcelWorksheet wxExcelSheet::ToWorksheet()
{
    wxExcelWorksheet worksheet;

    if ( IsWorksheet() )
        CloneDispatch(this, &worksheet);
    return worksheet;
}

#if WXAUTOEXCEL_USE_CHARTS

wxExcelChart wxExcelSheet::ToChart()
{
    wxExcelChart chart;

    if ( IsChart() )
        CloneDispatch(this, &chart);
    return chart;
}

#endif // #if WXAUTOEXCEL_USE_CHARTS

bool wxExcelSheet::Copy()
{
    WXAUTOEXCEL_CALL_METHOD0_BOOL("Copy");
}

bool wxExcelSheet::CopyAfterOrBefore(wxExcelSheet sheetAfterOrBefore, bool after)
{
    return DoOrderedCopyOrMove(true, sheetAfterOrBefore, after);
}

bool wxExcelSheet::Move()
{
    WXAUTOEXCEL_CALL_METHOD0_BOOL("Move");
}

bool wxExcelSheet::MoveAfterOrBefore(wxExcelSheet sheetAfterOrBefore, bool after)
{
    return DoOrderedCopyOrMove(false, sheetAfterOrBefore, after);
}

bool wxExcelSheet::Delete()
{
    WXAUTOEXCEL_CALL_METHOD0_BOOL("Delete");
}

bool wxExcelSheet::DoOrderedCopyOrMove(bool copy, wxExcelSheet sheetAfterOrBefore, bool after)
{
    wxVariant vAfterOrBefore;

    if ( ObjectToVariant(&sheetAfterOrBefore, vAfterOrBefore) )
    {
        vAfterOrBefore.SetName(after? wxS("After") : wxS("Before"));
        WXAUTOEXCEL_CALL_METHOD1_BOOL(copy ? "Copy" : "Move", vAfterOrBefore);
    }
    return false;
}


} // namespace wxAutoExcel
