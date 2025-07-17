/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////


#include "wx/wxAutoExcel_prec.h"

#include "wx/wxAutoExcelDataTable.h"

#if WXAUTOEXCEL_USE_CHARTS

#include "wx/wxAutoExcelChartFormat.h"
#include "wx/wxAutoExcelBorders.h"
#include "wx/wxAutoExcelFont.h"

#include "wx/wxAutoExcel_private.h"

namespace wxAutoExcel {

// ***** class wxExcelDataTable METHODS *****

void wxExcelDataTable::Delete()
{
    WXAUTOEXCEL_CALL_METHOD0_RET("Delete", "null");
}

void wxExcelDataTable::Select()
{
    WXAUTOEXCEL_CALL_METHOD0_RET("Select", "null");
}

// ***** class wxExcelDataTable PROPERTIES *****

wxExcelBorder wxExcelDataTable::GetBorder()
{
    wxExcelBorder border;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("Border", border);
}

wxExcelFont wxExcelDataTable::GetFont()
{
    wxExcelFont font;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("Font", font);
}

wxExcelChartFormat wxExcelDataTable::GetFormat()
{
    wxExcelChartFormat chartFormat;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("Format", chartFormat);
}

bool wxExcelDataTable::GetHasBorderHorizontal()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("HasBorderHorizontal");
}

void wxExcelDataTable::SetHasBorderHorizontal(bool hasBorderHorizontal)
{
    InvokePutProperty(wxS("HasBorderHorizontal"), hasBorderHorizontal);
}

bool wxExcelDataTable::GetHasBorderOutline()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("HasBorderOutline");
}

void wxExcelDataTable::SetHasBorderOutline(bool hasBorderOutline)
{
    InvokePutProperty(wxS("HasBorderOutline"), hasBorderOutline);
}

bool wxExcelDataTable::GetHasBorderVertical()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("HasBorderVertical");
}

void wxExcelDataTable::SetHasBorderVertical(bool hasBorderVertical)
{
    InvokePutProperty(wxS("HasBorderVertical"), hasBorderVertical);
}

bool wxExcelDataTable::GetShowLegendKey()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("ShowLegendKey");
}

void wxExcelDataTable::SetShowLegendKey(bool showLegendKey)
{
    InvokePutProperty(wxS("ShowLegendKey"), showLegendKey);
}


} // namespace wxAutoExcel


#endif // #if WXAUTOEXCEL_USE_CHARTS
