/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Modified by:
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////


#include "wx/wxAutoExcel_prec.h"

#include "wx/wxAutoExcelLegendEntries.h"

#if WXAUTOEXCEL_USE_CHARTS

#include "wx/wxAutoExcelChartFormat.h"
#include "wx/wxAutoExcelFont.h"
#include "wx/wxAutoExcelLegendKey.h"

#include "wx/wxAutoExcelPrivate.h"

namespace wxAutoExcel {

// ***** class wxExcelLegendEntry METHODS *****

bool wxExcelLegendEntry::Delete()
{
    WXAUTOEXCEL_CALL_METHOD0_BOOL("Delete");
}

bool wxExcelLegendEntry::Select()
{
    WXAUTOEXCEL_CALL_METHOD0_BOOL("Select");
}

// ***** class wxExcelLegendEntry PROPERTIES *****


bool wxExcelLegendEntry::GetAutoScaleFont()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("AutoScaleFont");
}

void wxExcelLegendEntry::SetAutoScaleFont(bool autoScaleFont)
{
    InvokePutProperty(wxS("AutoScaleFont"), autoScaleFont);
}

wxExcelFont wxExcelLegendEntry::GetFont()
{
    wxExcelFont font;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("Font", font);
}

wxExcelChartFormat wxExcelLegendEntry::GetFormat()
{
    wxExcelChartFormat chartFormat;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("Format", chartFormat);
}

double wxExcelLegendEntry::GetHeight()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("Height");
}

long wxExcelLegendEntry::GetIndex()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("Index");
}

double wxExcelLegendEntry::GetLeft()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("Left");
}

wxExcelLegendKey wxExcelLegendEntry::GetLegendKey()
{
    wxExcelLegendKey legendKey;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("LegendKey", legendKey);
}

double wxExcelLegendEntry::GetTop()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("Top");
}

double wxExcelLegendEntry::GetWidth()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("Width");
}

// ***** class wxExcelLegendEntries METHODS *****

wxExcelLegendEntry wxExcelLegendEntries::Item(long index)
{
    wxASSERT( index > 0 );
    
    wxExcelLegendEntry entry;
    WXAUTOEXCEL_CALL_METHOD1_OBJECT("Item", index, entry);
}

wxExcelLegendEntry wxExcelLegendEntries::operator[](long index)
{
    return Item(index);
}


// ***** class wxExcelLegendEntries PROPERTIES *****

long wxExcelLegendEntries::GetCount()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("Count");
}

} // namespace wxAutoExcel

#endif // #if WXAUTOEXCEL_USE_CHARTS
