/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Modified by:
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////


#include "wx/wxAutoExcel_prec.h"

#include "wx/wxAutoExcelDisplayUnitLabel.h"

#if WXAUTOEXCEL_USE_CHARTS

#include "wx/wxAutoExcelCharacters.h"
#include "wx/wxAutoExcelChartFormat.h"

#include "wx/wxAutoExcel_private.h"

namespace wxAutoExcel {
// ***** class wxExcelDisplayUnitLabel METHODS *****

bool wxExcelDisplayUnitLabel::Delete()
{
    WXAUTOEXCEL_CALL_METHOD0_BOOL("Delete");
}

bool wxExcelDisplayUnitLabel::Select()
{
    WXAUTOEXCEL_CALL_METHOD0_BOOL("Select");
}

// ***** class wxExcelDisplayUnitLabel PROPERTIES *****

bool wxExcelDisplayUnitLabel::GetAutoScaleFont()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("AutoScaleFont");
}

void wxExcelDisplayUnitLabel::SetAutoScaleFont(bool autoScaleFont)
{
    InvokePutProperty(wxS("AutoScaleFont"), autoScaleFont);
}

wxString wxExcelDisplayUnitLabel::GetCaption()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("Caption");
}

void wxExcelDisplayUnitLabel::SetCaption(const wxString& caption)
{
    InvokePutProperty(wxS("Caption"), caption);
}

wxExcelCharacters wxExcelDisplayUnitLabel::GetCharacters(long start, long* length)
{
   wxExcelCharacters characters;
   WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT(Length, length);
   WXAUTOEXCEL_PROPERTY_OBJECT_GET2("Characters", start, vLength, characters);
}

wxExcelChartFormat wxExcelDisplayUnitLabel::GetFormat()
{
    wxExcelChartFormat chartFormat;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("Format", chartFormat);
}

long wxExcelDisplayUnitLabel::GetHorizontalAlignment()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("HorizontalAlignment");
}

void wxExcelDisplayUnitLabel::SetHorizontalAlignment(long horizontalAlignment)
{
    InvokePutProperty(wxS("HorizontalAlignment"), horizontalAlignment);
}

double wxExcelDisplayUnitLabel::GetLeft()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("Left");
}

void wxExcelDisplayUnitLabel::SetLeft(double left)
{
    InvokePutProperty(wxS("Left"), left);
}

wxString wxExcelDisplayUnitLabel::GetName()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("Name");
}

long wxExcelDisplayUnitLabel::GetOrientation()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("Orientation");
}

void wxExcelDisplayUnitLabel::SetOrientation(long orientation)
{
    InvokePutProperty(wxS("Orientation"), orientation);
}

XlChartElementPosition wxExcelDisplayUnitLabel::GetPosition()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("Position", XlChartElementPosition, xlChartElementPositionAutomatic);
}

void wxExcelDisplayUnitLabel::SetPosition(XlChartElementPosition position)
{
    InvokePutProperty(wxS("Position"), (long)position);
}

long wxExcelDisplayUnitLabel::GetReadingOrder()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("ReadingOrder");
}

void wxExcelDisplayUnitLabel::SetReadingOrder(long readingOrder)
{
    InvokePutProperty(wxS("ReadingOrder"), readingOrder);
}

bool wxExcelDisplayUnitLabel::GetShadow()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("Shadow");
}

void wxExcelDisplayUnitLabel::SetShadow(bool shadow)
{
    InvokePutProperty(wxS("Shadow"), shadow);
}

wxString wxExcelDisplayUnitLabel::GetText()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("Text");
}

void wxExcelDisplayUnitLabel::SetText(const wxString& text)
{
    InvokePutProperty(wxS("Text"), text);
}

double wxExcelDisplayUnitLabel::GetTop()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("Top");
}

void wxExcelDisplayUnitLabel::SetTop(double top)
{
    InvokePutProperty(wxS("Top"), top);
}

long wxExcelDisplayUnitLabel::GetVerticalAlignment()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("VerticalAlignment");
}

void wxExcelDisplayUnitLabel::SetVerticalAlignment(long verticalAlignment)
{
    InvokePutProperty(wxS("VerticalAlignment"), verticalAlignment);
}

} // namespace wxAutoExcel

#endif // #if WXAUTOEXCEL_USE_CHARTS
