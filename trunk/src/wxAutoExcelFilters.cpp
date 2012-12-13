/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Modified by:
// RCS-ID:      $Id: $
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////


#include "wx/wxAutoExcel_prec.h"

#include "wx/wxAutoExcelFilters.h"

#include "wx/wxAutoExcelPrivate.h"

namespace wxAutoExcel {

// ***** class wxExcelFilter PROPERTIES *****

long wxExcelFilter::GetCount()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("Count");
}

wxVariant wxExcelFilter::GetCriteria1()
{
    wxVariant vResult;

    InvokeGetProperty(wxS("Criteria1"), vResult);
    return vResult;
}

wxVariant wxExcelFilter::GetCriteria2()
{
    wxVariant vResult;

    InvokeGetProperty(wxS("Criteria2"), vResult);
    return vResult;
}

bool wxExcelFilter::GetOn()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("On");
}

XlAutoFilterOperator wxExcelFilter::GetOperator()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("Operator", XlAutoFilterOperator, xlAnd);
}


// ***** class wxExcelFilters PROPERTIES *****

long wxExcelFilters::GetCount()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("Count");
}

wxExcelFilter wxExcelFilters::GetItem(long index)
{
    wxASSERT( index > 0 );
        
    wxExcelFilter filter;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("Item", filter);
}

wxExcelFilter wxExcelFilters::operator[](long index)
{
    return GetItem(index);
}


} // namespace wxAutoExcel
