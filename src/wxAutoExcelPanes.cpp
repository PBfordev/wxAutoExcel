/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// License:     MIT license
/////////////////////////////////////////////////////////////////////////////


#include "wx/wxAutoExcel_prec.h"

#include "wx/wxAutoExcelPanes.h"

#include "wx/wxAutoExcelRange.h"

#include "wx/wxAutoExcel_private.h"

namespace wxAutoExcel {

// ***** class wxExcelPane METHODS *****

bool wxExcelPane::Activate()
{
    WXAUTOEXCEL_CALL_METHOD0_BOOL("Activate");
}

bool wxExcelPane::LargeScroll(long* down, long* up, long* toRight, long* toLeft)
{
    WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT_NAME(Down, down);
    WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT_NAME(Up, up);
    WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT_NAME(ToRight, toRight);
    WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT_NAME(ToLeft, toLeft);

    WXAUTOEXCEL_CALL_METHOD4("LargeScroll", vDown, vUp, vToRight, vToLeft, "bool", false);
    return vResult.GetBool();
}

long wxExcelPane::PointsToScreenPixelsX(double points)
{
    WXAUTOEXCEL_CALL_METHOD1_LONG("PointsToScreenPixelsX", points, 0);
}

long wxExcelPane::PointsToScreenPixelsY(double points)
{
   WXAUTOEXCEL_CALL_METHOD1_LONG("PointsToScreenPixelsY", points, 0);
}

void wxExcelPane::ScrollIntoView(long left, long top, long width, long height, wxXlTribool start)
{
    WXAUTOEXCEL_OPTIONALCPPTBOOL_TO_OPTIONALVARIANT(Start, start);
    WXAUTOEXCEL_CALL_METHOD5_RET("ScrollIntoView", left, top, width, height, vStart, "null");
}

bool wxExcelPane::SmallScroll(long* down, long* up, long* toRight, long* toLeft)
{
    WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT_NAME(Down, down);
    WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT_NAME(Up, up);
    WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT_NAME(ToRight, toRight);
    WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT_NAME(ToLeft, toLeft);

    WXAUTOEXCEL_CALL_METHOD4("SmallScroll", vDown, vUp, vToRight, vToLeft, "bool", false);
    return vResult.GetBool();
}

// ***** class wxExcelPane PROPERTIES *****

long wxExcelPane::GetIndex()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("Index");
}


long wxExcelPane::GetScrollColumn()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("ScrollColumn");
}

void wxExcelPane::SetScrollColumn(long scrollColumn)
{
    InvokePutProperty(wxS("ScrollColumn"), scrollColumn);
}

long wxExcelPane::GetScrollRow()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("ScrollRow");
}

void wxExcelPane::SetScrollRow(long scrollRow)
{
    InvokePutProperty(wxS("ScrollRow"), scrollRow);
}

wxExcelRange wxExcelPane::GetVisibleRange()
{
    wxExcelRange range;
    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("VisibleRange", range);
}


// ***** class wxExcelPanes PROPERTIES *****

long wxExcelPanes::GetCount()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("Count");
}

wxExcelPane wxExcelPanes::GetItem(long index)
{
    wxASSERT( index > 0 );

    wxExcelPane pane;
    WXAUTOEXCEL_PROPERTY_OBJECT_GET1("Item", index, pane);
}

wxExcelPane wxExcelPanes::operator[](long index)
{
    return GetItem(index);
}

} // namespace wxAutoExcel