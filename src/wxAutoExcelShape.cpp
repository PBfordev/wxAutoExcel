/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Modified by:
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////

#include "wx/wxAutoExcel_prec.h"

#include "wx/wxAutoExcelShape.h"

#if WXAUTOEXCEL_USE_SHAPES

#include <wx/vector.h>

#include "wx/wxAutoExcelAdjustments.h"
#include "wx/wxAutoExcelRange.h"
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
#include "wx/wxAutoExcelTextEffectFormat.h"
#include "wx/wxAutoExcelThreeDFormat.h"
#include "wx/wxAutoExcelGroupShapes.h"
#include "wx/wxAutoExcelShapeRange.h"
#include "wx/wxAutoExcelShapeNodes.h"
#include "wx/wxAutoExcelHyperlinks.h"
#include "wx/wxAutoExcelTextFrame.h"
#include "wx/wxAutoExcelTextFrame2.h"
#include "wx/wxAutoExcelChart.h"

#include "wx/wxAutoExcelPrivate.h"

namespace wxAutoExcel {

// ***** class wxExcelShape METHODS *****

void wxExcelShape::Apply()
{
    WXAUTOEXCEL_CALL_METHOD0_RET("Apply", "null");
}

void wxExcelShape::Copy()
{
    WXAUTOEXCEL_CALL_METHOD0_RET("Copy", "null");
}

void wxExcelShape::CopyPicture(XlPictureAppearance* appearance, XlCopyPictureFormat* format)
{
    WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT_NAME(Appearance, ((long*)appearance));
    WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT_NAME(Format, ((long*)format));

    WXAUTOEXCEL_CALL_METHOD2_RET("CopyPicture", vAppearance, vFormat, "null");
}

void wxExcelShape::Cut()
{
    WXAUTOEXCEL_CALL_METHOD0_RET("Cut", "null");
}

void wxExcelShape::Delete()
{
    WXAUTOEXCEL_CALL_METHOD0_RET("Delete", "null");
}

void wxExcelShape::Duplicate()
{
    WXAUTOEXCEL_CALL_METHOD0_RET("Duplicate", "null");
}

void wxExcelShape::Flip(MsoFlipCmd flipCmd)
{    
    WXAUTOEXCEL_CALL_METHOD1_RET("Flip", (long)flipCmd, "null");
}

void wxExcelShape::IncrementLeft(double increment)
{
    WXAUTOEXCEL_CALL_METHOD1_RET("IncrementLeft", increment, "null");
}

void wxExcelShape::IncrementRotation(double increment)
{
    WXAUTOEXCEL_CALL_METHOD1_RET("IncrementRotation", increment, "null");
}

void wxExcelShape::IncrementTop(double increment)
{
    WXAUTOEXCEL_CALL_METHOD1_RET("IncrementTop", increment, "null");
}

void wxExcelShape::PickUp()
{
    WXAUTOEXCEL_CALL_METHOD0_RET("PickUp", "null");
}

void wxExcelShape::RerouteConnections()
{
    WXAUTOEXCEL_CALL_METHOD0_RET("RerouteConnections", "null");
}

void wxExcelShape::ScaleHeight(double factor, MsoTriState relativeToOriginalSize, MsoScaleFrom* scale)
{
    WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT(Scale, ((long*)scale));
    WXAUTOEXCEL_CALL_METHOD3_RET("ScaleHeight", factor, (long)relativeToOriginalSize, vScale, "null");
}

void wxExcelShape::ScaleWidth(double factor, MsoTriState relativeToOriginalSize, MsoScaleFrom* scale)
{
    WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT(Scale, ((long*)scale));
    WXAUTOEXCEL_CALL_METHOD3_RET("ScaleWidth", factor, (long)relativeToOriginalSize, vScale, "null");    
}

void wxExcelShape::Select(wxXlTribool replace)
{
    WXAUTOEXCEL_OPTIONALCPPTBOOL_TO_OPTIONALVARIANT(Replace, replace);

    WXAUTOEXCEL_CALL_METHOD1_RET("Select", vReplace, "null");
}

void wxExcelShape::SetShapesDefaultProperties()
{
    WXAUTOEXCEL_CALL_METHOD0_RET("SetShapesDefaultProperties", "null");
}

wxExcelShapeRange wxExcelShape::Ungroup()
{    
    wxExcelShapeRange range;

    WXAUTOEXCEL_CALL_METHOD0_OBJECT("Ungroup", range);    
}

void wxExcelShape::ZOrder(MsoZOrderCmd ZOrderCmd)
{
    WXAUTOEXCEL_CALL_METHOD1_RET("ZOrder", (long)ZOrderCmd, "null");
}

// ***** class wxExcelShape PROPERTIES *****

wxExcelAdjustments wxExcelShape::GetAdjustments()
{
    wxExcelAdjustments adjustments;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("Adjustments", adjustments);
}

wxString wxExcelShape::GetAlternativeText()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("AlternativeText");
}

void wxExcelShape::SetAlternativeText(const wxString& alternativeText)
{
    InvokePutProperty(wxS("AlternativeText"), alternativeText);
}

MsoAutoShapeType wxExcelShape::GetAutoShapeType()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("AutoShapeType", MsoAutoShapeType, msoShapeRectangle);
}

void wxExcelShape::SetAutoShapeType(MsoAutoShapeType autoShapeType)
{
    InvokePutProperty(wxS("AutoShapeType"), (long)autoShapeType);
}

MsoBackgroundStyleIndex wxExcelShape::GetBackgroundStyle()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("BackgroundStyle", MsoBackgroundStyleIndex,  msoBackgroundStyleNone);
}

MsoBlackWhiteMode wxExcelShape::GetBlackWhiteMode()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("BlackWhiteMode", MsoBlackWhiteMode, msoBlackWhiteAutomatic);
}

void wxExcelShape::SetBlackWhiteMode(MsoBlackWhiteMode blackWhiteMode)
{
    InvokePutProperty(wxS("BlackWhiteMode"), (long)blackWhiteMode);
}

wxExcelRange wxExcelShape::GetBottomRightCell()
{
    wxExcelRange range;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("BottomRightCell", range);
}

wxExcelCalloutFormat wxExcelShape::GetCallout()
{
    wxExcelCalloutFormat calloutFormat;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("Callout", calloutFormat);
}

#if WXAUTOEXCEL_USE_CHARTS

wxExcelChart wxExcelShape::GetChart()
{
    wxExcelChart chart;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("Chart", chart);
}

#endif // #if WXAUTOEXCEL_USE_CHARTS

MsoTriState wxExcelShape::GetChild()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("Child", MsoTriState, msoFalse);
}

long wxExcelShape::GetConnectionSiteCount()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("ConnectionSiteCount");
}

MsoTriState wxExcelShape::GetConnector()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("Connector", MsoTriState, msoFalse);
}

wxExcelConnectorFormat  wxExcelShape::GetConnectorFormat()
{
    wxExcelConnectorFormat  connectorFormat ;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("ConnectorFormat", connectorFormat);
}

wxExcelControlFormat  wxExcelShape::GetControlFormat()
{
    wxExcelControlFormat controlFormat;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("ControlFormat", controlFormat);
}

wxExcelFillFormat wxExcelShape::GetFill()
{
    wxExcelFillFormat fillFormat;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("Fill", fillFormat);
}

XlFormControl wxExcelShape::GetFormControlType()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("FormControlType", XlFormControl, xlButtonControl);
}

wxExcelGlowFormat wxExcelShape::GetGlow()
{
    wxExcelGlowFormat glowFormat;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("Glow", glowFormat);
}

wxExcelGroupShapes wxExcelShape::GetGroupItems()
{
    wxExcelGroupShapes groupShapes;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("GroupItems", groupShapes);
}

MsoTriState wxExcelShape::GetHasChart()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("HasChart", MsoTriState, msoFalse);
}

double wxExcelShape::GetHeight()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("Height");
}

MsoTriState wxExcelShape::GetHorizontalFlip()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("HorizontalFlip", MsoTriState, msoFalse);
}

wxExcelHyperlink wxExcelShape::GetHyperlink()
{
    wxExcelHyperlink hyperlink;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("Hyperlink", hyperlink);
}

long wxExcelShape::GetID()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("ID");
}

double wxExcelShape::GetLeft()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("Left");
}

wxExcelLineFormat wxExcelShape::GetLine()
{
    wxExcelLineFormat lineFormat;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("Line", lineFormat);
}

wxExcelLinkFormat wxExcelShape::GetLinkFormat()
{
    wxExcelLinkFormat linkFormat;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("LinkFormat", linkFormat);
}

MsoTriState wxExcelShape::GetLockAspectRatio()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("LockAspectRatio", MsoTriState, msoFalse);
}

void wxExcelShape::SetLockAspectRatio(MsoTriState lockAspectRatio)
{
    InvokePutProperty(wxS("LockAspectRatio"), (long)lockAspectRatio);
}

bool wxExcelShape::GetLocked()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("Locked");
}

void wxExcelShape::SetLocked(bool locked)
{
    InvokePutProperty(wxS("Locked"), locked);
}

wxString wxExcelShape::GetName()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("Name");
}

wxExcelShapeNodes wxExcelShape::GetNodes()
{
    wxExcelShapeNodes shapeNodes;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("Nodes", shapeNodes);
}

wxExcelOLEFormat wxExcelShape::GetOLEFormat()
{
    wxExcelOLEFormat OLEFormat;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("OLEFormat", OLEFormat);

}

wxString wxExcelShape::GetOnAction()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("OnAction");
}

void wxExcelShape::SetOnAction(const wxString& onAction)
{
    InvokePutProperty(wxS("OnAction"), onAction);
}


wxExcelShape wxExcelShape::GetParentGroup()
{
    wxExcelShape shape;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("ParentGroup", shape);
}

wxExcelPictureFormat wxExcelShape::GetPictureFormat()
{
    wxExcelPictureFormat pictureFormat ;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("PictureFormat", pictureFormat);
}

XlPlacement wxExcelShape::GetPlacement()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("Placement", XlPlacement, xlMoveAndSize);
}

void wxExcelShape::SetPlacement(XlPlacement  placement)
{
    InvokePutProperty(wxS("Placement"), (long)placement);
}

wxExcelReflectionFormat wxExcelShape::GetReflection()
{
    wxExcelReflectionFormat reflectionFormat;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("Reflection", reflectionFormat);
}

double wxExcelShape::GetRotation()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("Rotation");
}

void wxExcelShape::SetRotation(double rotation)
{
    InvokePutProperty(wxS("Rotation"), rotation);
}

wxExcelShadowFormat wxExcelShape::GetShadow()
{
    wxExcelShadowFormat shadowFormat;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("Shadow", shadowFormat);
}

MsoShapeStyleIndex wxExcelShape::GetShapeStyle()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("ShapeStyle", MsoShapeStyleIndex, msoShapeStyle1);
}

void wxExcelShape::SetShapeStyle(MsoShapeStyleIndex shapeStyle)
{
    InvokePutProperty(wxS("ShapeStyle"), (long)shapeStyle);
}

wxExcelSoftEdgeFormat wxExcelShape::GetSoftEdge()
{
    wxExcelSoftEdgeFormat softEdgeFormat;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("SoftEdge", softEdgeFormat);
}

wxExcelTextEffectFormat wxExcelShape::GetTextEffect()
{
    wxExcelTextEffectFormat textEffectFormat;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("TextEffect", textEffectFormat);
}

wxExcelTextFrame wxExcelShape::GetTextFrame()
{
    wxExcelTextFrame textFrame;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("TextFrame", textFrame);
}

wxExcelTextFrame2  wxExcelShape::GetTextFrame2()
{
    wxExcelTextFrame2 textFrame2;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("TextFrame2", textFrame2);
}

wxExcelThreeDFormat wxExcelShape::GetThreeD()
{
    wxExcelThreeDFormat threeDFormat;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("ThreeD", threeDFormat);
}

double wxExcelShape::GetTop()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("Top");
}

void wxExcelShape::SetTop(double top)
{
    InvokePutProperty(wxS("Top"), top);
}

wxExcelRange wxExcelShape::GetTopLeftCell()
{
    wxExcelRange range;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("TopLeftCell", range);
}

MsoShapeType  wxExcelShape::GetType()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("Type", MsoShapeType, msoAutoShape);
}

MsoTriState wxExcelShape::GetVerticalFlip()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("VerticalFlip", MsoTriState, msoFalse);
}


wxVector<wxPoint2DDouble> wxExcelShape::GetVertices()
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

MsoTriState wxExcelShape::GetVisible()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("Visible", MsoTriState, msoFalse);
}

void wxExcelShape::SetVisible(MsoTriState visible)
{
    InvokePutProperty(wxS("Visible"), (long)visible);
}

double wxExcelShape::GetWidth()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("Width");
}

void wxExcelShape::SetWidth(double width)
{
    InvokePutProperty(wxS("Width"), width);
}

long wxExcelShape::GetZOrderPosition()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("ZOrderPosition");
}

} // namespace wxAutoExcel

#endif // #if WXAUTOEXCEL_USE_SHAPES
