/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Modified by:
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////


#include "wx/wxAutoExcel_prec.h"

#include "wx/wxAutoExcelAxisTitle.h"

#if WXAUTOEXCEL_USE_CHARTS

#include "wx/wxAutoExcelCharacters.h"
#include "wx/wxAutoExcelChartFormat.h"

#include "wx/wxAutoExcel_private.h"

namespace wxAutoExcel {

// ***** class wxExcelAxisTitle METHODS *****

bool wxExcelAxisTitle::Delete()
{
    WXAUTOEXCEL_CALL_METHOD0_BOOL("Delete");
}

bool wxExcelAxisTitle::Select()
{
    WXAUTOEXCEL_CALL_METHOD0_BOOL("Select");
}

// ***** class wxExcelAxisTitle PROPERTIES *****


bool wxExcelAxisTitle::GetAutoScaleFont()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("AutoScaleFont");
}

void wxExcelAxisTitle::SetAutoScaleFont(bool autoScaleFont)
{
    InvokePutProperty(wxS("AutoScaleFont"), autoScaleFont);
}

wxString wxExcelAxisTitle::GetCaption()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("Caption");
}

void wxExcelAxisTitle::SetCaption(const wxString& caption)
{
    InvokePutProperty(wxS("Caption"), caption);
}

wxExcelCharacters wxExcelAxisTitle::GetCharacters()
{
    wxExcelCharacters characters;
    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("Characters", characters);
}

wxExcelChartFormat wxExcelAxisTitle::GetFormat()
{
    wxExcelChartFormat chartFormat;
    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("Format", chartFormat);
}

wxString wxExcelAxisTitle::GetFormula()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("Formula");
}

void wxExcelAxisTitle::SetFormula(const wxString& formula)
{
    InvokePutProperty(wxS("Formula"), formula);
}

wxString wxExcelAxisTitle::GetFormulaLocal()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("FormulaLocal");
}

void wxExcelAxisTitle::SetFormulaLocal(const wxString& formulaLocal)
{
    InvokePutProperty(wxS("FormulaLocal"), formulaLocal);
}

wxString wxExcelAxisTitle::GetFormulaR1C1()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("FormulaR1C1");
}

void wxExcelAxisTitle::SetFormulaR1C1(const wxString& formulaR1C1)
{
    InvokePutProperty(wxS("FormulaR1C1"), formulaR1C1);
}

wxString wxExcelAxisTitle::GetFormulaR1C1Local()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("FormulaR1C1Local");
}

void wxExcelAxisTitle::SetFormulaR1C1Local(const wxString& formulaR1C1Local)
{
    InvokePutProperty(wxS("FormulaR1C1Local"), formulaR1C1Local);
}

long wxExcelAxisTitle::GetHorizontalAlignment()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("HorizontalAlignment");
}

void wxExcelAxisTitle::SetHorizontalAlignment(long horizontalAlignment)
{
    InvokePutProperty(wxS("HorizontalAlignment"), horizontalAlignment);
}

bool wxExcelAxisTitle::GetIncludeInLayout()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("IncludeInLayout");
}

void wxExcelAxisTitle::SetIncludeInLayout(bool includeInLayout)
{
    InvokePutProperty(wxS("IncludeInLayout"), includeInLayout);
}

double wxExcelAxisTitle::GetLeft()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("Left");
}

void wxExcelAxisTitle::SetLeft(double left)
{
    InvokePutProperty(wxS("Left"), left);
}

wxString wxExcelAxisTitle::GetName()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("Name");
}

long wxExcelAxisTitle::GetOrientation()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("Orientation");
}

void wxExcelAxisTitle::SetOrientation(long orientation)
{
    InvokePutProperty(wxS("Orientation"), orientation);
}


XlChartElementPosition wxExcelAxisTitle::GetPosition()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("Position", XlChartElementPosition, xlChartElementPositionAutomatic);
}

void wxExcelAxisTitle::SetPosition(XlChartElementPosition position)
{
    InvokePutProperty(wxS("Position"), (long)position);
}

long wxExcelAxisTitle::GetReadingOrder()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("ReadingOrder");
}

void wxExcelAxisTitle::SetReadingOrder(long readingOrder)
{
    InvokePutProperty(wxS("ReadingOrder"), readingOrder);
}

bool wxExcelAxisTitle::GetShadow()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("Shadow");
}

void wxExcelAxisTitle::SetShadow(bool shadow)
{
    InvokePutProperty(wxS("Shadow"), shadow);
}

wxString wxExcelAxisTitle::GetText()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("Text");
}

void wxExcelAxisTitle::SetText(const wxString& text)
{
    InvokePutProperty(wxS("Text"), text);
}

double wxExcelAxisTitle::GetTop()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("Top");
}

void wxExcelAxisTitle::SetTop(double top)
{
    InvokePutProperty(wxS("Top"), top);
}

long wxExcelAxisTitle::GetVerticalAlignment()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("VerticalAlignment");
}

void wxExcelAxisTitle::SetVerticalAlignment(long verticalAlignment)
{
    InvokePutProperty(wxS("VerticalAlignment"), verticalAlignment);
}

double wxExcelAxisTitle::GetWidth()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("Width");
}


} // namespace wxAutoExcel

#endif // #if WXAUTOEXCEL_USE_CHARTS
