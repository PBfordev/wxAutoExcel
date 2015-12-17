/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Modified by:
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////


#include "wx/wxAutoExcel_prec.h"

#include "wx/wxAutoExcelTextFrame.h"

#if WXAUTOEXCEL_USE_SHAPES || WXAUTOEXCEL_USE_CHARTS

#include "wx/wxAutoExcelCharacters.h"

#include "wx/wxAutoExcel_private.h"

namespace wxAutoExcel {

// ***** class wxExcelTextFrame METHODS *****

wxExcelCharacters wxExcelTextFrame::Characters(long* start, long* length)
{
    wxExcelCharacters characters;
    WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT_NAME(Start, start);
    WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT_NAME(Length, length);

    WXAUTOEXCEL_CALL_METHOD2("Characters", vStart, vLength, "void*", characters);    
    VariantToObject(vResult, &characters);
    return characters;
}

bool wxExcelTextFrame::GetAutoSize()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("AutoSize");
}

void wxExcelTextFrame::SetAutoSize(bool autoSize)
{
    InvokePutProperty(wxS("AutoSize"), autoSize);
}

XlHAlign wxExcelTextFrame::GetHorizontalAlignment()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("HorizontalAlignment", XlHAlign, xlHAlignGeneral);
}

void wxExcelTextFrame::SetHorizontalAlignment(XlHAlign horizontalAlignment)
{
    InvokePutProperty(wxS("HorizontalAlignment"), (long)horizontalAlignment);
}

double wxExcelTextFrame::GetMarginBottom()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("MarginBottom");
}

void wxExcelTextFrame::SetMarginBottom(double marginBottom)
{
    InvokePutProperty(wxS("MarginBottom"), marginBottom);
}

double wxExcelTextFrame::GetMarginLeft()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("MarginLeft");
}

void wxExcelTextFrame::SetMarginLeft(double marginLeft)
{
    InvokePutProperty(wxS("MarginLeft"), marginLeft);
}

double wxExcelTextFrame::GetMarginRight()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("MarginRight");
}

void wxExcelTextFrame::SetMarginRight(double marginRight)
{
    InvokePutProperty(wxS("MarginRight"), marginRight);
}

double wxExcelTextFrame::GetMarginTop()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("MarginTop");
}

void wxExcelTextFrame::SetMarginTop(double marginTop)
{
    InvokePutProperty(wxS("MarginTop"), marginTop);
}

long wxExcelTextFrame::GetOrientation()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("Orientation");
}

void wxExcelTextFrame::SetOrientation(long orientation)
{
    InvokePutProperty(wxS("Orientation"), orientation);
}

long wxExcelTextFrame::GetReadingOrder()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("ReadingOrder");
}

void wxExcelTextFrame::SetReadingOrder(long readingOrder)
{
    InvokePutProperty(wxS("ReadingOrder"), readingOrder);
}

XlVAlign  wxExcelTextFrame::GetVerticalAlignment()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("VerticalAlignment", XlVAlign, xlVAlignTop);
}

void wxExcelTextFrame::SetVerticalAlignment(XlVAlign  verticalAlignment)
{
    InvokePutProperty(wxS("VerticalAlignment"), (long)verticalAlignment);
}

} // namespace wxAutoExcel

#endif // #if WXAUTOEXCEL_USE_SHAPES || WXAUTOEXCEL_USE_CHARTS
