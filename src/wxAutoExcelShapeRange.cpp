/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Modified by:
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////


#include "wx/wxAutoExcel_prec.h"

#include "wx/wxAutoExcelShapeRange.h"

#if WXAUTOEXCEL_USE_SHAPES

#include <wx/vector.h>

#include "wx/wxAutoExcelShape.h"
#include "wx/wxAutoExcelGroupShapes.h"
#include "wx/wxAutoExcelShapeNodes.h"
#include "wx/wxAutoExcelAdjustments.h"
#include "wx/wxAutoExcelCalloutFormat.h"
#include "wx/wxAutoExcelConnectorFormat.h"
#include "wx/wxAutoExcelControlFormat.h"
#include "wx/wxAutoExcelFillFormat.h"
#include "wx/wxAutoExcelGlowFormat.h"
#include "wx/wxAutoExcelLineFormat.h"
#include "wx/wxAutoExcelLinkFormat.h"
#include "wx/wxAutoExcelOLEFormat.h"
#include "wx/wxAutoExcelPictureFormat.h"
#include "wx/wxAutoExcelReflectionFormat.h"
#include "wx/wxAutoExcelShadowFormat.h"
#include "wx/wxAutoExcelSoftEdgeFormat.h"
#include "wx/wxAutoExcelColorFormat.h"
#include "wx/wxAutoExcelTextEffectFormat.h"
#include "wx/wxAutoExcelThreeDFormat.h"

#include "wx/wxAutoExcelTextFrame.h"
#include "wx/wxAutoExcelTextFrame2.h"

#include "wx/wxAutoExcelChart.h"

#include "wx/wxAutoExcel_private.h"

namespace wxAutoExcel {
// ***** class wxExcelShapeRange METHODS *****

void wxExcelShapeRange::Align(MsoAlignCmd alignCmd)
{
    WXAUTOEXCEL_CALL_METHOD2_RET("Align", (long)alignCmd, (long)msoFalse, "null");
}

void wxExcelShapeRange::Apply()
{
    WXAUTOEXCEL_CALL_METHOD0_RET("Apply", "null");
}

void wxExcelShapeRange::Delete()
{
    WXAUTOEXCEL_CALL_METHOD0_RET("Delete", "null");
}

void wxExcelShapeRange::Distribute(MsoDistributeCmd distributeCmd)
{
    WXAUTOEXCEL_CALL_METHOD2_RET("Distribute", (long)distributeCmd, (long)msoFalse, "null");
}

wxExcelShapeRange wxExcelShapeRange::Duplicate()
{    
    wxExcelShapeRange range;
    
    WXAUTOEXCEL_CALL_METHOD0_OBJECT("Duplicate", range);    
}

void wxExcelShapeRange::Flip(MsoFlipCmd flipCmd)
{
    WXAUTOEXCEL_CALL_METHOD1_RET("Flip", (long)flipCmd, "null");
}

wxExcelShape wxExcelShapeRange::Group()
{    
    wxExcelShape shape;

    WXAUTOEXCEL_CALL_METHOD0_OBJECT("Group", shape); 
}

void wxExcelShapeRange::IncrementLeft(double increment)
{
    WXAUTOEXCEL_CALL_METHOD1_RET("IncrementLeft", increment, "null");
}

void wxExcelShapeRange::IncrementRotation(double increment)
{
    WXAUTOEXCEL_CALL_METHOD1_RET("IncrementRotation", increment, "null");
}

void wxExcelShapeRange::IncrementTop(double increment)
{
    WXAUTOEXCEL_CALL_METHOD1_RET("IncrementTop", increment, "null");
}

wxExcelShape wxExcelShapeRange::Item(long index)
{    
    wxASSERT( index > 0 );

    wxExcelShape shape;
    WXAUTOEXCEL_CALL_METHOD1_OBJECT("Item", index, shape);
}

wxExcelShape wxExcelShapeRange::operator[](long index)
{
    return Item(index);
}

wxExcelShape wxExcelShapeRange::Item(const wxString& name)
{ 
    wxExcelShape shape;

    WXAUTOEXCEL_CALL_METHOD1_OBJECT("Item", name, shape);
}

wxExcelShape wxExcelShapeRange::operator[](const wxString& name)
{
    return Item(name);
}

void wxExcelShapeRange::PickUp()
{
    WXAUTOEXCEL_CALL_METHOD0_RET("PickUp", "null");
}

wxExcelShape wxExcelShapeRange::Regroup()
{    
    wxExcelShape shape;

    WXAUTOEXCEL_CALL_METHOD0_OBJECT("Regroup", shape);    
}

void wxExcelShapeRange::RerouteConnections()
{
    WXAUTOEXCEL_CALL_METHOD0_RET("RerouteConnections", "null");
}

void wxExcelShapeRange::ScaleHeight(double factor, MsoTriState relativeToOriginalSize,
                                    MsoScaleFrom* scale)
{

    WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT(Scale, (long*)scale);

    WXAUTOEXCEL_CALL_METHOD3_RET("ScaleHeight", factor, (long)relativeToOriginalSize, vScale, "null");
}

void wxExcelShapeRange::ScaleWidth(double factor, MsoTriState relativeToOriginalSize,
                                    MsoScaleFrom* scale)
{
    WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT(Scale, (long*)scale);

    WXAUTOEXCEL_CALL_METHOD3_RET("ScaleWidth", factor, (long)relativeToOriginalSize, vScale, "null");    
}

void wxExcelShapeRange::Select()
{
    WXAUTOEXCEL_CALL_METHOD0_RET("Select", "null");
}

void wxExcelShapeRange::SetShapesDefaultProperties()
{
    WXAUTOEXCEL_CALL_METHOD0_RET("SetShapesDefaultProperties", "null");
}

wxExcelShapeRange wxExcelShapeRange::Ungroup()
{    
    wxExcelShapeRange range;
    
    WXAUTOEXCEL_CALL_METHOD0_OBJECT("Ungroup", range);
}

void wxExcelShapeRange::ZOrder(MsoZOrderCmd ZOrderCmd)
{
    WXAUTOEXCEL_CALL_METHOD1_RET("ZOrder", (long)ZOrderCmd, "null");
}

// ***** class wxExcelShapeRange PROPERTIES *****

wxExcelAdjustments wxExcelShapeRange::GetAdjustments()
{
    wxExcelAdjustments adjustments;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("Adjustments", adjustments);

}

wxString wxExcelShapeRange::GetAlternativeText()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("AlternativeText");
}


MsoAutoShapeType wxExcelShapeRange::GetAutoShapeType()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("AutoShapeType", MsoAutoShapeType, msoShapeRectangle);
}

void wxExcelShapeRange::SetAutoShapeType(MsoAutoShapeType autoShapeType)
{
    InvokePutProperty(wxS("AutoShapeType"), (long)autoShapeType);
}

MsoBackgroundStyleIndex wxExcelShapeRange::GetBackgroundStyle()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("BackgroundStyle", MsoBackgroundStyleIndex, msoBackgroundStyleNone);
}

void wxExcelShapeRange::SetBackgroundStyle(MsoBackgroundStyleIndex backgroundStyle)
{
    InvokePutProperty(wxS("BackgroundStyle"), (long)backgroundStyle);
}

MsoBlackWhiteMode wxExcelShapeRange::GetBlackWhiteMode()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("BlackWhiteMode", MsoBlackWhiteMode, msoBlackWhiteAutomatic);
}

wxExcelCalloutFormat wxExcelShapeRange::GetCallout()
{
    wxExcelCalloutFormat calloutFormat;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("Callout", calloutFormat);
}

#if WXAUTOEXCEL_USE_CHARTS

wxExcelChart wxExcelShapeRange::GetChart()
{
    wxExcelChart chart;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("Chart", chart);
}

#endif // #if WXAUTOEXCEL_USE_CHARTS

MsoTriState wxExcelShapeRange::GetChild()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("Child", MsoTriState, msoFalse);
}

long wxExcelShapeRange::GetConnectionSiteCount()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("ConnectionSiteCount");
}

MsoTriState wxExcelShapeRange::GetConnector()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("Connector", MsoTriState, msoFalse);
}

wxExcelConnectorFormat  wxExcelShapeRange::GetConnectorFormat()
{
    wxExcelConnectorFormat connectorFormat;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("ConnectorFormat", connectorFormat);
}

long wxExcelShapeRange::GetCount()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("Count");
}

wxExcelFillFormat wxExcelShapeRange::GetFill()
{
    wxExcelFillFormat fillFormat;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("Fill", fillFormat);
}

wxExcelGlowFormat wxExcelShapeRange::GetGlow()
{
    wxExcelGlowFormat glowFormat;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("Glow", glowFormat);
}

wxExcelGroupShapes wxExcelShapeRange::GetGroupItems()
{
    wxExcelGroupShapes groupShapes;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("GroupItems", groupShapes);
}

bool wxExcelShapeRange::GetHasChart()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("HasChart");
}

double wxExcelShapeRange::GetHeight()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("Height");
}

MsoTriState wxExcelShapeRange::GetHorizontalFlip()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("HorizontalFlip", MsoTriState, msoFalse);
}

long wxExcelShapeRange::GetID()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("ID");
}

double wxExcelShapeRange::GetLeft()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("Left");
}

void wxExcelShapeRange::SetLeft(double left)
{
    InvokePutProperty(wxS("Left"), left);
}

wxExcelLineFormat wxExcelShapeRange::GetLine()
{
    wxExcelLineFormat lineFormat;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("Line", lineFormat);
}

MsoTriState wxExcelShapeRange::GetLockAspectRatio()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("LockAspectRatio", MsoTriState, msoFalse);
}

void wxExcelShapeRange::SetLockAspectRatio(MsoTriState lockAspectRatio)
{
    InvokePutProperty(wxS("LockAspectRatio"), (long)lockAspectRatio);
}

wxString wxExcelShapeRange::GetName()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("Name");
}

void wxExcelShapeRange::SetName(const wxString& name)
{
    InvokePutProperty(wxS("Name"), name);
}

wxExcelShapeNodes wxExcelShapeRange::GetNodes()
{
    wxExcelShapeNodes shapeNodes;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("Nodes", shapeNodes);
}


wxExcelShape wxExcelShapeRange::GetParentGroup()
{
    wxExcelShape shape;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("ParentGroup", shape);
}

wxExcelPictureFormat wxExcelShapeRange::GetPictureFormat()
{
    wxExcelPictureFormat pictureFormat;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("PictureFormat", pictureFormat);
}

wxExcelReflectionFormat wxExcelShapeRange::GetReflection()
{
    wxExcelReflectionFormat reflectionFormat;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("Reflection", reflectionFormat);
}

double wxExcelShapeRange::GetRotation()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("Rotation");
}

wxExcelShadowFormat wxExcelShapeRange::GetShadow()
{
    wxExcelShadowFormat shadowFormat;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("Shadow", shadowFormat);
}

MsoShapeStyleIndex wxExcelShapeRange::GetShapeStyle()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("ShapeStyle", MsoShapeStyleIndex, msoShapeStyle1);
}

void wxExcelShapeRange::SetShapeStyle(MsoShapeStyleIndex shapeStyle)
{
    InvokePutProperty(wxS("ShapeStyle"), (long)shapeStyle);
}

wxExcelSoftEdgeFormat wxExcelShapeRange::GetSoftEdge()
{
    wxExcelSoftEdgeFormat softEdgeFormat;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("SoftEdge", softEdgeFormat);

}

wxExcelTextEffectFormat wxExcelShapeRange::GetTextEffect()
{
    wxExcelTextEffectFormat textEffectFormat;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("TextEffect", textEffectFormat);

}

wxExcelTextFrame wxExcelShapeRange::GetTextFrame()
{
    wxExcelTextFrame textFrame;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("TextFrame", textFrame);

}

wxExcelTextFrame2 wxExcelShapeRange::GetTextFrame2()
{
    wxExcelTextFrame2 textFrame2;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("TextFrame2", textFrame2);
}

wxExcelThreeDFormat wxExcelShapeRange::GetThreeD()
{
    wxExcelThreeDFormat threeDFormat;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("ThreeD", threeDFormat);

}

double wxExcelShapeRange::GetTop()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("Top");
}

MsoShapeType  wxExcelShapeRange::GetType()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("Type", MsoShapeType , msoAutoShape);
}

MsoTriState wxExcelShapeRange::GetVerticalFlip()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("VerticalFlip", MsoTriState, msoFalse);
}

wxVector<wxPoint2DDouble> wxExcelShapeRange::GetVertices()
{
    wxVariant vResult;
    wxVector<wxPoint2DDouble> points;

    if ( InvokeGetProperty(wxS("Vertices"), vResult) )
    {
        if ( vResult.GetType() == wxS("list") && vResult.GetCount() % 2 == 0 )
        {
            size_t pointCount = vResult.GetCount() / 2;
            for ( size_t i = 0; i < pointCount; i++ )            
            {
                wxPoint2DDouble point;
                point.m_x = vResult[i];
                point.m_y = vResult[i+pointCount]; // we get the array with all x coordinates and then all y coordinates
                points.push_back(point);
            }
        }
    }
    return points;
}

MsoTriState wxExcelShapeRange::GetVisible()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("Visible", MsoTriState, msoFalse);
}

double wxExcelShapeRange::GetWidth()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("Width");
}

long wxExcelShapeRange::GetZOrderPosition()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("ZOrderPosition");
}

} // namespace wxAutoExcel

#endif // #if WXAUTOEXCEL_USE_SHAPES
