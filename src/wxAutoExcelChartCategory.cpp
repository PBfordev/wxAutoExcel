/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// License:     MIT license
/////////////////////////////////////////////////////////////////////////////


#include "wx/wxAutoExcel_prec.h"

#include "wx/wxAutoExcelChartCategory.h"

#if WXAUTOEXCEL_USE_CHARTS

#include "wx/wxAutoExcelRange.h"

#include "wx/wxAutoExcel_private.h"

namespace wxAutoExcel {

// ***** class  wxExcelChartCategory PROPERTIES *****

bool wxExcelChartCategory::GetIsFiltered()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("IsFiltered");
}

wxString wxExcelChartCategory::GetName()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("Name");
}

void wxExcelChartCategory::SetName(const wxString& name)
{
    InvokePutProperty(wxS("Name"), name);
}

// ***** class  wxExcelCategoryCollection  METHODS *****

wxExcelChartCategory wxExcelCategoryCollection::Item(long index)
{
    wxExcelChartCategory category;
    WXAUTOEXCEL_CALL_METHOD1_OBJECT("Item", index, category);
}

wxExcelChartCategory wxExcelCategoryCollection::operator[](long index)
{
    return Item(index);
}

// ***** class  wxExcelCategoryCollection PROPERTIES *****

long  wxExcelCategoryCollection::GetCount()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("Count");
}


} // namespace wxAutoExcel

#endif // #if WXAUTOEXCEL_USE_CHARTS
