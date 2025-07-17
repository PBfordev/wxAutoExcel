/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// License:     MIT license
/////////////////////////////////////////////////////////////////////////////


#include "wx/wxAutoExcel_prec.h"

#include "wx/wxAutoExcelCalloutFormat.h"

#if WXAUTOEXCEL_USE_SHAPES

#include "wx/wxAutoExcel_private.h"

namespace wxAutoExcel {

    // ***** class wxExcelCalloutFormat METHODS *****

void wxExcelCalloutFormat::AutomaticLength()
{
    WXAUTOEXCEL_CALL_METHOD0_RET("AutomaticLength", "null");
}

void wxExcelCalloutFormat::CustomDrop(double drop)
{
    WXAUTOEXCEL_CALL_METHOD1_RET("CustomDrop", drop, "null");
}

void wxExcelCalloutFormat::CustomLength(double length)
{
    WXAUTOEXCEL_CALL_METHOD1_RET("CustomLength", length, "null");
}

void wxExcelCalloutFormat::PresetDrop(MsoCalloutDropType dropType)
{
    WXAUTOEXCEL_CALL_METHOD1_RET("PresetDrop", (long)dropType, "null");
}

// ***** class wxExcelCalloutFormat PROPERTIES *****

MsoTriState wxExcelCalloutFormat::GetAccent()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("Accent", MsoTriState, msoFalse);
}

void wxExcelCalloutFormat::SetAccent(MsoTriState accent)
{
    InvokePutProperty(wxS("Accent"), (long)accent);
}

MsoCalloutAngleType wxExcelCalloutFormat::GetAngle()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("Angle", MsoCalloutAngleType, msoCalloutAngleAutomatic);
}

void wxExcelCalloutFormat::SetAngle(MsoCalloutAngleType angle)
{
    InvokePutProperty(wxS("Angle"), (long)angle);
}


MsoTriState wxExcelCalloutFormat::GetAutoAttach()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("AutoAttach", MsoTriState, msoFalse);
}

void wxExcelCalloutFormat::SetAutoAttach(MsoTriState autoAttach)
{
    InvokePutProperty(wxS("AutoAttach"), (long)autoAttach);
}

MsoTriState wxExcelCalloutFormat::GetAutoLength()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("AutoLength", MsoTriState, msoFalse);
}

void wxExcelCalloutFormat::SetAutoLength(MsoTriState autoLength)
{
    InvokePutProperty(wxS("AutoLength"), (long)autoLength);
}

MsoTriState wxExcelCalloutFormat::GetBorder()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("Border", MsoTriState, msoFalse);
}

void wxExcelCalloutFormat::SetBorder(MsoTriState border)
{
    InvokePutProperty(wxS("Border"), (long)border);
}


double wxExcelCalloutFormat::GetDrop()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("Drop");
}

MsoCalloutDropType wxExcelCalloutFormat::GetDropType()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("DropType", MsoCalloutDropType, msoCalloutDropCenter);
}

double wxExcelCalloutFormat::GetGap()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("Gap");
}

void wxExcelCalloutFormat::SetGap(double gap)
{
    InvokePutProperty(wxS("Gap"), gap);
}

double wxExcelCalloutFormat::GetLength()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("Length");
}


MsoCalloutType  wxExcelCalloutFormat::GetType()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("Type", MsoCalloutType, msoCalloutOne);
}

void wxExcelCalloutFormat::SetType(MsoCalloutType  type)
{
    InvokePutProperty(wxS("Type"), (long)type);
}

} // namespace wxAutoExcel

#endif // #if WXAUTOEXCEL_USE_SHAPES
