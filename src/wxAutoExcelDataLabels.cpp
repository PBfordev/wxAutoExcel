/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Modified by:
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////


#include "wx/wxAutoExcel_prec.h"

#include "wx/wxAutoExcelDataLabels.h"

#if WXAUTOEXCEL_USE_CHARTS

#include "wx/wxAutoExcelCharacters.h"
#include "wx/wxAutoExcelChartFormat.h"

#include "wx/wxAutoExcel_private.h"

namespace wxAutoExcel {

// ***** class wxExcelDataLabel METHODS *****

bool wxExcelDataLabel::Delete()
{
    WXAUTOEXCEL_CALL_METHOD0_BOOL("Delete");
}

bool wxExcelDataLabel::Select()
{
    WXAUTOEXCEL_CALL_METHOD0_BOOL("Select");
}

// ***** class wxExcelDataLabel PROPERTIES *****

bool wxExcelDataLabel::GetAutoScaleFont()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("AutoScaleFont");
}

void wxExcelDataLabel::SetAutoScaleFont(bool autoScaleFont)
{
    InvokePutProperty(wxS("AutoScaleFont"), autoScaleFont);
}

bool wxExcelDataLabel::GetAutoText()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("AutoText");
}

void wxExcelDataLabel::SetAutoText(bool autoText)
{
    InvokePutProperty(wxS("AutoText"), autoText);
}

wxString wxExcelDataLabel::GetCaption()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("Caption");
}

void wxExcelDataLabel::SetCaption(const wxString& caption)
{
    InvokePutProperty(wxS("Caption"), caption);
}

wxExcelCharacters wxExcelDataLabel::GetCharacters(long start, long* length)
{
   wxExcelCharacters characters;
   WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT(Length, length);
   WXAUTOEXCEL_PROPERTY_OBJECT_GET2("Characters", start, vLength, characters);
}

wxExcelChartFormat wxExcelDataLabel::GetFormat()
{
    wxExcelChartFormat chartFormat;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("Format", chartFormat);
}

long wxExcelDataLabel::GetHorizontalAlignment()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("HorizontalAlignment");
}

void wxExcelDataLabel::SetHorizontalAlignment(long horizontalAlignment)
{
    InvokePutProperty(wxS("HorizontalAlignment"), horizontalAlignment);
}

double wxExcelDataLabel::GetLeft()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("Left");
}

void wxExcelDataLabel::SetLeft(double left)
{
    InvokePutProperty(wxS("Left"), left);
}

wxString wxExcelDataLabel::GetName()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("Name");
}

wxString wxExcelDataLabel::GetNumberFormat()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("NumberFormat");
}

void wxExcelDataLabel::SetNumberFormat(const wxString& numberFormat)
{
    InvokePutProperty(wxS("NumberFormat"), numberFormat);
}

bool wxExcelDataLabel::GetNumberFormatLinked()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("NumberFormatLinked");
}

void wxExcelDataLabel::SetNumberFormatLinked(bool numberFormatLinked)
{
    InvokePutProperty(wxS("NumberFormatLinked"), numberFormatLinked);
}

wxString wxExcelDataLabel::GetNumberFormatLocal()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("NumberFormatLocal");
}

void wxExcelDataLabel::SetNumberFormatLocal(const wxString& numberFormatLocal)
{
    InvokePutProperty(wxS("NumberFormatLocal"), numberFormatLocal);
}

long wxExcelDataLabel::GetOrientation()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("Orientation");
}

void wxExcelDataLabel::SetOrientation(long orientation)
{
    InvokePutProperty(wxS("Orientation"), orientation);
}

XlDataLabelPosition wxExcelDataLabel::GetPosition()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("Position", XlDataLabelPosition, xlLabelPositionAbove);
}

void wxExcelDataLabel::SetPosition(XlDataLabelPosition position)
{
    InvokePutProperty(wxS("Position"), (long)position);
}

long wxExcelDataLabel::GetReadingOrder()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("ReadingOrder");
}

void wxExcelDataLabel::SetReadingOrder(long readingOrder)
{
    InvokePutProperty(wxS("ReadingOrder"), readingOrder);
}

wxString wxExcelDataLabel::GetSeparator()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("Separator");
}

void wxExcelDataLabel::SetSeparator(const wxString& separator)
{
    InvokePutProperty(wxS("Separator"), separator);
}

bool wxExcelDataLabel::GetShadow()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("Shadow");
}

void wxExcelDataLabel::SetShadow(bool shadow)
{
    InvokePutProperty(wxS("Shadow"), shadow);
}

bool wxExcelDataLabel::GetShowBubbleSize()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("ShowBubbleSize");
}

void wxExcelDataLabel::SetShowBubbleSize(bool showBubbleSize)
{
    InvokePutProperty(wxS("ShowBubbleSize"), showBubbleSize);
}

bool wxExcelDataLabel::GetShowCategoryName()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("ShowCategoryName");
}

void wxExcelDataLabel::SetShowCategoryName(bool showCategoryName)
{
    InvokePutProperty(wxS("ShowCategoryName"), showCategoryName);
}

bool wxExcelDataLabel::GetShowLegendKey()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("ShowLegendKey");
}

void wxExcelDataLabel::SetShowLegendKey(bool showLegendKey)
{
    InvokePutProperty(wxS("ShowLegendKey"), showLegendKey);
}

bool wxExcelDataLabel::GetShowPercentage()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("ShowPercentage");
}

void wxExcelDataLabel::SetShowPercentage(bool showPercentage)
{
    InvokePutProperty(wxS("ShowPercentage"), showPercentage);
}

bool wxExcelDataLabel::GetShowRange()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("ShowRange");
}

void wxExcelDataLabel::SetShowRange(bool showRange)
{
    InvokePutProperty(wxS("ShowRange"), showRange);
}

bool wxExcelDataLabel::GetShowSeriesName()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("ShowSeriesName");
}

void wxExcelDataLabel::SetShowSeriesName(bool showSeriesName)
{
    InvokePutProperty(wxS("ShowSeriesName"), showSeriesName);
}

bool wxExcelDataLabel::GetShowValue()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("ShowValue");
}

void wxExcelDataLabel::SetShowValue(bool showValue)
{
    InvokePutProperty(wxS("ShowValue"), showValue);
}

wxString wxExcelDataLabel::GetText()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("Text");
}

void wxExcelDataLabel::SetText(const wxString& text)
{
    InvokePutProperty(wxS("Text"), text);
}

double wxExcelDataLabel::GetTop()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("Top");
}

void wxExcelDataLabel::SetTop(double top)
{
    InvokePutProperty(wxS("Top"), top);
}

long wxExcelDataLabel::GetVerticalAlignment()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("VerticalAlignment");
}

void wxExcelDataLabel::SetVerticalAlignment(long verticalAlignment)
{
    InvokePutProperty(wxS("VerticalAlignment"), verticalAlignment);
}

// ***** class wxExcelDataLabels METHODS *****

bool wxExcelDataLabels::Delete()
{
    WXAUTOEXCEL_CALL_METHOD0_BOOL("Delete");
}

wxExcelDataLabel wxExcelDataLabels::Item(long index)
{
    wxExcelDataLabel label;
    WXAUTOEXCEL_CALL_METHOD1_OBJECT("Item", index, label);
}

wxExcelDataLabel wxExcelDataLabels::operator[](long index)
{
    return Item(index);
}

void wxExcelDataLabels::Propagate(long index)
{
    WXAUTOEXCEL_CALL_METHOD1_RET("Propagate", index, "null");
}

bool wxExcelDataLabels::Select()
{
    WXAUTOEXCEL_CALL_METHOD0_BOOL("Select");
}

// ***** class wxExcelDataLabels PROPERTIES *****


bool wxExcelDataLabels::GetAutoScaleFont()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("AutoScaleFont");
}

void wxExcelDataLabels::SetAutoScaleFont(bool autoScaleFont)
{
    InvokePutProperty(wxS("AutoScaleFont"), autoScaleFont);
}

bool wxExcelDataLabels::GetAutoText()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("AutoText");
}

void wxExcelDataLabels::SetAutoText(bool autoText)
{
    InvokePutProperty(wxS("AutoText"), autoText);
}

long wxExcelDataLabels::GetCount()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("Count");
}

wxExcelChartFormat wxExcelDataLabels::GetFormat()
{
    wxExcelChartFormat chartFormat;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("Format", chartFormat);
}

long wxExcelDataLabels::GetHorizontalAlignment()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("HorizontalAlignment");
}

void wxExcelDataLabels::SetHorizontalAlignment(long horizontalAlignment)
{
    InvokePutProperty(wxS("HorizontalAlignment"), horizontalAlignment);
}

wxString wxExcelDataLabels::GetName()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("Name");
}

wxString wxExcelDataLabels::GetNumberFormat()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("NumberFormat");
}

void wxExcelDataLabels::SetNumberFormat(const wxString& numberFormat)
{
    InvokePutProperty(wxS("NumberFormat"), numberFormat);
}

bool wxExcelDataLabels::GetNumberFormatLinked()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("NumberFormatLinked");
}

void wxExcelDataLabels::SetNumberFormatLinked(bool numberFormatLinked)
{
    InvokePutProperty(wxS("NumberFormatLinked"), numberFormatLinked);
}

wxString wxExcelDataLabels::GetNumberFormatLocal()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("NumberFormatLocal");
}

void wxExcelDataLabels::SetNumberFormatLocal(const wxString& numberFormatLocal)
{
    InvokePutProperty(wxS("NumberFormatLocal"), numberFormatLocal);
}

long wxExcelDataLabels::GetOrientation()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("Orientation");
}

void wxExcelDataLabels::SetOrientation(long orientation)
{
    InvokePutProperty(wxS("Orientation"), orientation);
}


XlDataLabelPosition wxExcelDataLabels::GetPosition()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("Position", XlDataLabelPosition, xlLabelPositionAbove);
}

void wxExcelDataLabels::SetPosition(XlDataLabelPosition position)
{
    InvokePutProperty(wxS("Position"), (long)position);
}

long wxExcelDataLabels::GetReadingOrder()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("ReadingOrder");
}

void wxExcelDataLabels::SetReadingOrder(long readingOrder)
{
    InvokePutProperty(wxS("ReadingOrder"), readingOrder);
}

wxString wxExcelDataLabels::GetSeparator()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("Separator");
}

void wxExcelDataLabels::SetSeparator(const wxString& separator)
{
    InvokePutProperty(wxS("Separator"), separator);
}

bool wxExcelDataLabels::GetShadow()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("Shadow");
}

void wxExcelDataLabels::SetShadow(bool shadow)
{
    InvokePutProperty(wxS("Shadow"), shadow);
}

bool wxExcelDataLabels::GetShowBubbleSize()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("ShowBubbleSize");
}

void wxExcelDataLabels::SetShowBubbleSize(bool showBubbleSize)
{
    InvokePutProperty(wxS("ShowBubbleSize"), showBubbleSize);
}

bool wxExcelDataLabels::GetShowCategoryName()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("ShowCategoryName");
}

void wxExcelDataLabels::SetShowCategoryName(bool showCategoryName)
{
    InvokePutProperty(wxS("ShowCategoryName"), showCategoryName);
}

bool wxExcelDataLabels::GetShowLegendKey()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("ShowLegendKey");
}

void wxExcelDataLabels::SetShowLegendKey(bool showLegendKey)
{
    InvokePutProperty(wxS("ShowLegendKey"), showLegendKey);
}

bool wxExcelDataLabels::GetShowPercentage()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("ShowPercentage");
}

void wxExcelDataLabels::SetShowPercentage(bool showPercentage)
{
    InvokePutProperty(wxS("ShowPercentage"), showPercentage);
}

bool wxExcelDataLabels::GetShowSeriesName()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("ShowSeriesName");
}

bool wxExcelDataLabels::GetShowRange()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("ShowRange");
}

void wxExcelDataLabels::SetShowRange(bool showRange)
{
    InvokePutProperty(wxS("ShowRange"), showRange);
}

void wxExcelDataLabels::SetShowSeriesName(bool showSeriesName)
{
    InvokePutProperty(wxS("ShowSeriesName"), showSeriesName);
}

bool wxExcelDataLabels::GetShowValue()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("ShowValue");
}

void wxExcelDataLabels::SetShowValue(bool showValue)
{
    InvokePutProperty(wxS("ShowValue"), showValue);
}

long wxExcelDataLabels::GetVerticalAlignment()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("VerticalAlignment");
}

void wxExcelDataLabels::SetVerticalAlignment(long verticalAlignment)
{
    InvokePutProperty(wxS("VerticalAlignment"), verticalAlignment);
}

} // namespace wxAutoExcel

#endif // #if WXAUTOEXCEL_USE_CHARTS
