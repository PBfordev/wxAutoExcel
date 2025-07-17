/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// License:     MIT license
/////////////////////////////////////////////////////////////////////////////


#include "wx/wxAutoExcel_prec.h"

#include "wx/wxAutoExcelChartTitle.h"

#if WXAUTOEXCEL_USE_CHARTS

#include "wx/wxAutoExcelCharacters.h"
#include "wx/wxAutoExcelChartFormat.h"

#include "wx/wxAutoExcel_private.h"

namespace wxAutoExcel {

// ***** class wxExcelChartTitle METHODS *****

bool wxExcelChartTitle::Delete()
{
    WXAUTOEXCEL_CALL_METHOD0_BOOL("Delete");
}

bool wxExcelChartTitle::Select()
{
    WXAUTOEXCEL_CALL_METHOD0_BOOL("Select");
}

// ***** class wxExcelChartTitle PROPERTIES *****

bool wxExcelChartTitle::GetAutoScaleFont()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("AutoScaleFont");
}

void wxExcelChartTitle::SetAutoScaleFont(bool autoScaleFont)
{
    InvokePutProperty(wxS("AutoScaleFont"), autoScaleFont);
}

wxString wxExcelChartTitle::GetCaption()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("Caption");
}

void wxExcelChartTitle::SetCaption(const wxString& caption)
{
    InvokePutProperty(wxS("Caption"), caption);
}

wxExcelCharacters wxExcelChartTitle::GetCharacters(long start, long* length)
{
   wxExcelCharacters characters;
   WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT(Length, length);
   WXAUTOEXCEL_PROPERTY_OBJECT_GET2("Characters", start, vLength, characters);
}

wxExcelChartFormat wxExcelChartTitle::GetFormat()
{
    wxExcelChartFormat chartFormat;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("Format", chartFormat);
}

wxString wxExcelChartTitle::GetFormula()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("Formula");
}

void wxExcelChartTitle::SetFormula(const wxString& formula)
{
    InvokePutProperty(wxS("Formula"), formula);
}

wxString wxExcelChartTitle::GetFormulaLocal()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("FormulaLocal");
}

void wxExcelChartTitle::SetFormulaLocal(const wxString& formulaLocal)
{
    InvokePutProperty(wxS("FormulaLocal"), formulaLocal);
}

wxString wxExcelChartTitle::GetFormulaR1C1()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("FormulaR1C1");
}

void wxExcelChartTitle::SetFormulaR1C1(const wxString& formulaR1C1)
{
    InvokePutProperty(wxS("FormulaR1C1"), formulaR1C1);
}

wxString wxExcelChartTitle::GetFormulaR1C1Local()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("FormulaR1C1Local");
}

void wxExcelChartTitle::SetFormulaR1C1Local(const wxString& formulaR1C1Local)
{
    InvokePutProperty(wxS("FormulaR1C1Local"), formulaR1C1Local);
}

double wxExcelChartTitle::GetHeight()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("Height");
}


long wxExcelChartTitle::GetHorizontalAlignment()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("HorizontalAlignment");
}

void wxExcelChartTitle::SetHorizontalAlignment(long horizontalAlignment)
{
    InvokePutProperty(wxS("HorizontalAlignment"), horizontalAlignment);
}

bool wxExcelChartTitle::GetIncludeInLayout()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("IncludeInLayout");
}

void wxExcelChartTitle::SetIncludeInLayout(bool includeInLayout)
{
    InvokePutProperty(wxS("IncludeInLayout"), includeInLayout);
}

double wxExcelChartTitle::GetLeft()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("Left");
}

void wxExcelChartTitle::SetLeft(double left)
{
    InvokePutProperty(wxS("Left"), left);
}

wxString wxExcelChartTitle::GetName()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("Name");
}

long wxExcelChartTitle::GetOrientation()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("Orientation");
}

void wxExcelChartTitle::SetOrientation(long orientation)
{
    InvokePutProperty(wxS("Orientation"), orientation);
}


XlChartElementPosition wxExcelChartTitle::GetPosition()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("Position",  XlChartElementPosition, xlChartElementPositionAutomatic);
}

void wxExcelChartTitle::SetPosition(XlChartElementPosition position)
{
    InvokePutProperty(wxS("Position"), (long)position);
}

long wxExcelChartTitle::GetReadingOrder()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("ReadingOrder");
}

void wxExcelChartTitle::SetReadingOrder(long readingOrder)
{
    InvokePutProperty(wxS("ReadingOrder"), readingOrder);
}

bool wxExcelChartTitle::GetShadow()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("Shadow");
}

void wxExcelChartTitle::SetShadow(bool shadow)
{
    InvokePutProperty(wxS("Shadow"), shadow);
}

wxString wxExcelChartTitle::GetText()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("Text");
}

void wxExcelChartTitle::SetText(const wxString& text)
{
    InvokePutProperty(wxS("Text"), text);
}

double wxExcelChartTitle::GetTop()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("Top");
}

void wxExcelChartTitle::SetTop(double top)
{
    InvokePutProperty(wxS("Top"), top);
}

long wxExcelChartTitle::GetVerticalAlignment()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("VerticalAlignment");
}

void wxExcelChartTitle::SetVerticalAlignment(long verticalAlignment)
{
    InvokePutProperty(wxS("VerticalAlignment"), verticalAlignment);
}

double wxExcelChartTitle::GetWidth()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("Width");
}

} // namespace wxAutoExcel


#endif // #if WXAUTOEXCEL_USE_CHARTS
