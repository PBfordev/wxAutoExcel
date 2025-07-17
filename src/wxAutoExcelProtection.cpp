/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2020 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////

#include "wx/wxAutoExcel_prec.h"

#include "wx/wxAutoExcelProtection.h"

#include "wx/wxAutoExcelAllowEditRanges.h"

#include "wx/wxAutoExcel_private.h"

namespace wxAutoExcel {

// ***** class wxExcelProtection PROPERTIES *****

bool wxExcelProtection::GetAllowDeletingColumns()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("AllowDeletingColumns");
}

bool wxExcelProtection::GetAllowDeletingRows()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("AllowDeletingRows");
}

wxExcelAllowEditRanges wxExcelProtection::GetAllowEditRanges()
{
    wxExcelAllowEditRanges allowEditRanges;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("AllowEditRanges", allowEditRanges);
}

bool wxExcelProtection::GetAllowFiltering()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("AllowFiltering");
}

bool wxExcelProtection::GetAllowFormattingCells()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("AllowFormattingCells");
}

bool wxExcelProtection::GetAllowFormattingColumns()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("AllowFormattingColumns");
}

bool wxExcelProtection::GetAllowFormattingRows()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("AllowFormattingRows");
}

bool wxExcelProtection::GetAllowInsertingColumns()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("AllowInsertingColumns");
}

bool wxExcelProtection::GetAllowInsertingHyperlinks()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("AllowInsertingHyperlinks");
}

bool wxExcelProtection::GetAllowInsertingRows()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("AllowInsertingRows");
}

bool wxExcelProtection::GetAllowSorting()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("AllowSorting");
}

bool wxExcelProtection::GetAllowUsingPivotTables()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("AllowUsingPivotTables");
}

} // namespace wxAutoExcel
