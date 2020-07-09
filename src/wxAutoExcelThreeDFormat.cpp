/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Modified by:
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////


#include "wx/wxAutoExcel_prec.h"

#include "wx/wxAutoExcelThreeDFormat.h"

#if WXAUTOEXCEL_USE_SHAPES || WXAUTOEXCEL_USE_CHARTS

#include "wx/wxAutoExcelApplication.h"
#include "wx/wxAutoExcelColorFormat.h"

#include "wx/wxAutoExcel_private.h"

namespace wxAutoExcel {

// ***** class wxExcelThreeDFormat METHODS *****

void wxExcelThreeDFormat::IncrementRotationHorizontal(double increment)
{
    WXAUTOEXCEL_CALL_METHOD1_RET("IncrementRotationHorizontal", increment, "null");
}

void wxExcelThreeDFormat::IncrementRotationVertical(double increment)
{
    WXAUTOEXCEL_CALL_METHOD1_RET("IncrementRotationVertical", increment, "null");
}

void wxExcelThreeDFormat::IncrementRotationX(double increment)
{
    WXAUTOEXCEL_CALL_METHOD1_RET("IncrementRotationX", increment, "null");
}

void wxExcelThreeDFormat::IncrementRotationY(double increment)
{
    WXAUTOEXCEL_CALL_METHOD1_RET("IncrementRotationY", increment, "null");
}

void wxExcelThreeDFormat::IncrementRotationZ(double increment)
{
    WXAUTOEXCEL_CALL_METHOD1_RET("IncrementRotationZ", increment, "null");
}

void wxExcelThreeDFormat::ResetRotation()
{
    WXAUTOEXCEL_CALL_METHOD0_RET("ResetRotation", "null");
}

void wxExcelThreeDFormat::SetExtrusionDirection(MsoPresetExtrusionDirection presetExtrusionDirection)
{
    WXAUTOEXCEL_CALL_METHOD1_RET("SetExtrusionDirection", (long)presetExtrusionDirection, "null");
}

void wxExcelThreeDFormat::SetPresetCamera(MsoPresetCamera presetCamera)
{
    WXAUTOEXCEL_CALL_METHOD1_RET("SetPresetCamera", (long)presetCamera, "null");
}

void wxExcelThreeDFormat::SetThreeDFormat(MsoPresetThreeDFormat presetThreeDFormat)
{
    WXAUTOEXCEL_CALL_METHOD1_RET("SetThreeDFormat", (long)presetThreeDFormat, "null");
}

// ***** class wxExcelThreeDFormat PROPERTIES *****


double wxExcelThreeDFormat::GetBevelBottomDepth()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("BevelBottomDepth");
}

void wxExcelThreeDFormat::SetBevelBottomDepth(double bevelBottomDepth)
{
    InvokePutProperty(wxS("BevelBottomDepth"), bevelBottomDepth);
}

double wxExcelThreeDFormat::GetBevelBottomInset()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("BevelBottomInset");
}

void wxExcelThreeDFormat::SetBevelBottomInset(double bevelBottomInset)
{
    InvokePutProperty(wxS("BevelBottomInset"), bevelBottomInset);
}

MsoBevelType wxExcelThreeDFormat::GetBevelBottomType()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("BevelBottomType", MsoBevelType, msoBevelNone);
}

void wxExcelThreeDFormat::SetBevelBottomType(MsoBevelType bevelBottomType)
{
    InvokePutProperty(wxS("BevelBottomType"), (long)bevelBottomType);
}

double wxExcelThreeDFormat::GetBevelTopDepth()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("BevelTopDepth");
}

void wxExcelThreeDFormat::SetBevelTopDepth(double bevelTopDepth)
{
    InvokePutProperty(wxS("BevelTopDepth"), bevelTopDepth);
}

double wxExcelThreeDFormat::GetBevelTopInset()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("BevelTopInset");
}

void wxExcelThreeDFormat::SetBevelTopInset(double bevelTopInset)
{
    InvokePutProperty(wxS("BevelTopInset"), bevelTopInset);
}

MsoBevelType wxExcelThreeDFormat::GetBevelTopType()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("BevelTopType", MsoBevelType, msoBevelNone);
}

void wxExcelThreeDFormat::SetBevelTopType(MsoBevelType bevelTopType)
{
    InvokePutProperty(wxS("BevelTopType"), (long)bevelTopType);
}

wxExcelColorFormat wxExcelThreeDFormat::GetContourColor()
{
    wxExcelColorFormat color;
    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("ContourColor", color);
}

double wxExcelThreeDFormat::GetContourWidth()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("ContourWidth");
}

void wxExcelThreeDFormat::SetContourWidth(double contourWidth)
{
    InvokePutProperty(wxS("ContourWidth"), contourWidth);
}

double wxExcelThreeDFormat::GetDepth()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("Depth");
}

void wxExcelThreeDFormat::SetDepth(double depth)
{
    InvokePutProperty(wxS("Depth"), depth);
}

wxExcelColorFormat wxExcelThreeDFormat::GetExtrusionColor()
{
    wxExcelColorFormat color;
    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("ExtrusionColor", color);
}

MsoExtrusionColorType wxExcelThreeDFormat::GetExtrusionColorType()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("ExtrusionColorType", MsoExtrusionColorType, msoExtrusionColorAutomatic);
}

void wxExcelThreeDFormat::SetExtrusionColorType(MsoExtrusionColorType extrusionColorType)
{
    InvokePutProperty(wxS("ExtrusionColorType"), (long)extrusionColorType);
}

double wxExcelThreeDFormat::GetFieldOfView()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("FieldOfView");
}

void wxExcelThreeDFormat::SetFieldOfView(double fieldOfView)
{
    InvokePutProperty(wxS("FieldOfView"), fieldOfView);
}

double wxExcelThreeDFormat::GetLightAngle()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("LightAngle");
}

void wxExcelThreeDFormat::SetLightAngle(double lightAngle)
{
    InvokePutProperty(wxS("LightAngle"), lightAngle);
}


MsoTriState wxExcelThreeDFormat::GetPerspective()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("Perspective", MsoTriState, msoFalse);
}

void wxExcelThreeDFormat::SetPerspective(MsoTriState perspective)
{
    InvokePutProperty(wxS("Perspective"), (long)perspective);
}

MsoPresetCamera wxExcelThreeDFormat::GetPresetCamera()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("PresetCamera", MsoPresetCamera, msoCameraLegacyObliqueTopLeft);
}

MsoPresetExtrusionDirection wxExcelThreeDFormat::GetPresetExtrusionDirection()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("PresetExtrusionDirection", MsoPresetExtrusionDirection, msoExtrusionBottomRight);
}

MsoLightRigType wxExcelThreeDFormat::GetPresetLighting()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("PresetLighting", MsoLightRigType, msoLightRigLegacyFlat1);
}

void wxExcelThreeDFormat::SetPresetLighting(MsoLightRigType presetLighting)
{
    InvokePutProperty(wxS("PresetLighting"), (long)presetLighting);
}

MsoPresetLightingDirection wxExcelThreeDFormat::GetPresetLightingDirection()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("PresetLightingDirection", MsoPresetLightingDirection, msoLightingTopLeft);
}

void wxExcelThreeDFormat::SetPresetLightingDirection(MsoPresetLightingDirection presetLightingDirection)
{
    InvokePutProperty(wxS("PresetLightingDirection"), (long)presetLightingDirection);
}

MsoPresetLightingSoftness wxExcelThreeDFormat::GetPresetLightingSoftness()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("PresetLightingSoftness", MsoPresetLightingSoftness, msoLightingNormal);
}

void wxExcelThreeDFormat::SetPresetLightingSoftness(MsoPresetLightingSoftness presetLightingSoftness)
{
    InvokePutProperty(wxS("PresetLightingSoftness"), (long)presetLightingSoftness);
}

MsoPresetMaterial wxExcelThreeDFormat::GetPresetMaterial()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("PresetMaterial", MsoPresetMaterial, msoMaterialMatte);
}

void wxExcelThreeDFormat::SetPresetMaterial(MsoPresetMaterial presetMaterial)
{
    InvokePutProperty(wxS("PresetMaterial"), (long)presetMaterial);
}

MsoPresetThreeDFormat wxExcelThreeDFormat::GetPresetThreeDFormat()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("PresetThreeDFormat", MsoPresetThreeDFormat, msoThreeD1);
}

MsoTriState wxExcelThreeDFormat::GetProjectText()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("ProjectText", MsoTriState, msoFalse);
}

void wxExcelThreeDFormat::SetProjectText(MsoTriState projectText)
{
    InvokePutProperty(wxS("ProjectText"), (long)projectText);
}

double wxExcelThreeDFormat::GetRotationX()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("RotationX");
}

void wxExcelThreeDFormat::SetRotationX(double rotationX)
{
    InvokePutProperty(wxS("RotationX"), rotationX);
}

double wxExcelThreeDFormat::GetRotationY()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("RotationY");
}

void wxExcelThreeDFormat::SetRotationY(double rotationY)
{
    InvokePutProperty(wxS("RotationY"), rotationY);
}

double wxExcelThreeDFormat::GetRotationZ()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("RotationZ");
}

void wxExcelThreeDFormat::SetRotationZ(double rotationZ)
{
    InvokePutProperty(wxS("RotationZ"), rotationZ);
}

MsoTriState wxExcelThreeDFormat::GetVisible()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("Visible", MsoTriState, msoFalse);
}

void wxExcelThreeDFormat::SetVisible(MsoTriState visible)
{
    InvokePutProperty(wxS("Visible"), (long)visible);
}

double wxExcelThreeDFormat::GetZ()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("Z");
}

void wxExcelThreeDFormat::SetZ(double z)
{
    InvokePutProperty(wxS("Z"), z);
}

} // namespace wxAutoExcel


#endif // #if WXAUTOEXCEL_USE_SHAPES || WXAUTOEXCEL_USE_CHARTS
