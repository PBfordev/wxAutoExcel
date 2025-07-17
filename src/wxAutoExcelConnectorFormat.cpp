/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////


#include "wx/wxAutoExcel_prec.h"

#include "wx/wxAutoExcelConnectorFormat.h"

#if WXAUTOEXCEL_USE_SHAPES

#include "wx/wxAutoExcelShape.h"

#include "wx/wxAutoExcel_private.h"

namespace wxAutoExcel {

// ***** class wxExcelConnectorFormat METHODS *****

void wxExcelConnectorFormat::BeginConnect(wxExcelShape connectedShape, long connectionSite)
{
    wxVariant vShape;
    if ( !wxExcelObject::ObjectToVariant(&connectedShape, vShape) )
        return;

    WXAUTOEXCEL_CALL_METHOD2_RET("BeginConnect", vShape, connectionSite, "null");
}

void wxExcelConnectorFormat::BeginDisconnect()
{
    WXAUTOEXCEL_CALL_METHOD0_RET("BeginDisconnect", "null");
}

void wxExcelConnectorFormat::EndConnect(wxExcelShape connectedShape, long connectionSite)
{
    wxVariant vShape;
    if ( !wxExcelObject::ObjectToVariant(&connectedShape, vShape) )
        return;

    WXAUTOEXCEL_CALL_METHOD2_RET("EndConnect", vShape, connectionSite, "null");
}

void wxExcelConnectorFormat::EndDisconnect()
{
    WXAUTOEXCEL_CALL_METHOD0_RET("EndDisconnect", "null");
}

// ***** class wxExcelConnectorFormat PROPERTIES *****

MsoTriState wxExcelConnectorFormat::GetBeginConnected()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("BeginConnected", MsoTriState, msoFalse);
}

wxExcelShape wxExcelConnectorFormat::GetBeginConnectedShape()
{
    wxExcelShape shape;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("BeginConnectedShape", shape);
}

long wxExcelConnectorFormat::GetBeginConnectionSite()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("BeginConnectionSite");
}

MsoTriState wxExcelConnectorFormat::GetEndConnected()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("EndConnected", MsoTriState, msoFalse);
}

wxExcelShape wxExcelConnectorFormat::GetEndConnectedShape()
{
    wxExcelShape shape;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("EndConnectedShape", shape);
}

long wxExcelConnectorFormat::GetEndConnectionSite()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("EndConnectionSite");
}

MsoConnectorType wxExcelConnectorFormat::GetType()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("Type", MsoConnectorType , msoConnectorStraight);
}

} // namespace wxAutoExcel

#endif // #if WXAUTOEXCEL_USE_SHAPES
