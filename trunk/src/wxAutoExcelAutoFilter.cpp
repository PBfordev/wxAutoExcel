/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Modified by:
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////


#include "wx/wxAutoExcel_prec.h"

#include "wx/wxAutoExcelAutoFilter.h"
#include "wx/wxAutoExcelFilters.h"
#include "wx/wxAutoExcelSort.h"
#include "wx/wxAutoExcelRange.h"

#include "wx/wxAutoExcelPrivate.h"

namespace wxAutoExcel {

// ***** class wxExcelAutoFilter METHODS *****

void wxExcelAutoFilter::ApplyFilter()
{
    WXAUTOEXCEL_CALL_METHOD0_RET("ApplyFilter", "null");
}

void wxExcelAutoFilter::ShowAllData()
{
    WXAUTOEXCEL_CALL_METHOD0_RET("ShowAllData", "null");
}

// ***** class wxExcelAutoFilter PROPERTIES *****

bool wxExcelAutoFilter::GetFilterMode()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("FilterMode");
}

wxExcelFilters wxExcelAutoFilter::GetFilters()
{
    wxExcelFilters filters;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("Filters", filters);
}

wxExcelRange wxExcelAutoFilter::GetRange()
{
    wxExcelRange range;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("Range", range);
}

wxExcelSort wxExcelAutoFilter::GetSort()
{
    wxExcelSort sort;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("Sort", sort);
}



} // namespace wxAutoExcel
