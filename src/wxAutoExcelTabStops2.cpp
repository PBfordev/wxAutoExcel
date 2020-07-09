/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Modified by:
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////


#include "wx/wxAutoExcel_prec.h"

#include "wx/wxAutoExcelTabStops2.h"

#if WXAUTOEXCEL_USE_SHAPES || WXAUTOEXCEL_USE_CHARTS

#include "wx/wxAutoExcel_private.h"

namespace wxAutoExcel {

// ***** class wxExcelTabStop2 METHODS *****

void wxExcelTabStop2::Clear()
{
    WXAUTOEXCEL_CALL_METHOD0_RET("Clear", "null");
}

// ***** class wxExcelTabStop2 PROPERTIES *****


double wxExcelTabStop2::GetPosition()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("Position");
}

void wxExcelTabStop2::SetPosition(double position)
{
    InvokePutProperty(wxS("Position"), position);
}

MsoTabStopType wxExcelTabStop2::GetType()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("Type", MsoTabStopType, msoTabStopLeft);
}

void wxExcelTabStop2::SetType(MsoTabStopType type)
{
    InvokePutProperty(wxS("Type"), (long)type);
}

// ***** class wxExcelTabStops2 METHODS *****

wxExcelTabStop2 wxExcelTabStops2::Add(MsoTabStopType type, double position)
{
    wxExcelTabStop2 tabStop;

    WXAUTOEXCEL_CALL_METHOD2("Add", (long)type, position, "void*", tabStop);
    VariantToObject(vResult, &tabStop);
    return tabStop;
}

wxExcelTabStop2 wxExcelTabStops2::Item(long index)
{
    wxASSERT( index > 0 );

    wxExcelTabStop2 tabStop;
    WXAUTOEXCEL_CALL_METHOD1_OBJECT("Item", index, tabStop);
}

wxExcelTabStop2 wxExcelTabStops2::operator[](long index)
{
    return Item(index);
}

// ***** class wxExcelTabStops2 PROPERTIES *****

long wxExcelTabStops2::GetCount()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("Count");
}


double wxExcelTabStops2::GetDefaultSpacing()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("DefaultSpacing");
}

void wxExcelTabStops2::SetDefaultSpacing(double defaultSpacing)
{
    InvokePutProperty(wxS("DefaultSpacing"), defaultSpacing);
}

} // namespace wxAutoExcel

#endif // #if WXAUTOEXCEL_USE_SHAPES || WXAUTOEXCEL_USE_CHARTS
