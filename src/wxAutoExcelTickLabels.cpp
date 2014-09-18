/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Modified by:
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////


#include "wx/wxAutoExcel_prec.h"

#include "wx/wxAutoExcelTickLabels.h"

#if WXAUTOEXCEL_USE_CHARTS

#include "wx/wxAutoExcelChartFormat.h"
#include "wx/wxAutoExcelFont.h"

#include "wx/wxAutoExcelPrivate.h"

namespace wxAutoExcel {

// ***** class wxExcelTickLabels METHODS *****

bool wxExcelTickLabels::Delete()
{
    WXAUTOEXCEL_CALL_METHOD0_BOOL("Delete");
}

bool wxExcelTickLabels::Select()
{
    WXAUTOEXCEL_CALL_METHOD0_BOOL("Select");
}

// ***** class wxExcelTickLabels PROPERTIES *****

long wxExcelTickLabels::GetAlignment()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("Alignment");
}

void wxExcelTickLabels::SetAlignment(long alignment)
{
    InvokePutProperty(wxS("Alignment"), alignment);
}

bool wxExcelTickLabels::GetAutoScaleFont()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("AutoScaleFont");
}

void wxExcelTickLabels::SetAutoScaleFont(bool autoScaleFont)
{
    InvokePutProperty(wxS("AutoScaleFont"), autoScaleFont);
}

long wxExcelTickLabels::GetDepth()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("Depth");
}

wxExcelFont wxExcelTickLabels::GetFont()
{
    wxExcelFont font;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("Font", font);
}

wxExcelChartFormat wxExcelTickLabels::GetFormat()
{
    wxExcelChartFormat chartFormat;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("Format", chartFormat);
}

bool wxExcelTickLabels::GetMultiLevel()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("MultiLevel");
}

void wxExcelTickLabels::SetMultiLevel(bool multiLevel)
{
    InvokePutProperty(wxS("MultiLevel"), multiLevel);
}

wxString wxExcelTickLabels::GetName()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("Name");
}

wxString wxExcelTickLabels::GetNumberFormat()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("NumberFormat");
}

void wxExcelTickLabels::SetNumberFormat(const wxString& numberFormat)
{
    InvokePutProperty(wxS("NumberFormat"), numberFormat);
}

bool wxExcelTickLabels::GetNumberFormatLinked()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("NumberFormatLinked");
}

void wxExcelTickLabels::SetNumberFormatLinked(bool numberFormatLinked)
{
    InvokePutProperty(wxS("NumberFormatLinked"), numberFormatLinked);
}

wxString wxExcelTickLabels::GetNumberFormatLocal()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("NumberFormatLocal");
}

void wxExcelTickLabels::SetNumberFormatLocal(const wxString& numberFormatLocal)
{
    InvokePutProperty(wxS("NumberFormatLocal"), numberFormatLocal);
}

long wxExcelTickLabels::GetOffset()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("Offset");
}

void wxExcelTickLabels::SetOffset(long offset)
{
    InvokePutProperty(wxS("Offset"), offset);
}

long wxExcelTickLabels::GetOrientation()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("Orientation");
}

void wxExcelTickLabels::SetOrientation(long orientation)
{
    InvokePutProperty(wxS("Orientation"), orientation);
}

long wxExcelTickLabels::GetReadingOrder()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("ReadingOrder");
}

void wxExcelTickLabels::SetReadingOrder(long readingOrder)
{
    InvokePutProperty(wxS("ReadingOrder"), readingOrder);
}

} // namespace wxAutoExcel


#endif // #if WXAUTOEXCEL_USE_CHARTS
