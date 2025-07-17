/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2020 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////

#include "wx/wxAutoExcel_prec.h"

#include "wx/wxAutoExcelModel3DFormat.h"

#include "wx/wxAutoExcel_private.h"

namespace wxAutoExcel {

// ***** class wxExcelModel3DFormat METHODS *****

void wxExcelModel3DFormat::IncrementRotationX(double increment)
{
    WXAUTOEXCEL_CALL_METHOD1_RET("IncrementRotationX", increment, "null");
}

void wxExcelModel3DFormat::IncrementRotationY(double increment)
{
    WXAUTOEXCEL_CALL_METHOD1_RET("IncrementRotationY", increment, "null");
}

void wxExcelModel3DFormat::IncrementRotationZ(double increment)
{
    WXAUTOEXCEL_CALL_METHOD1_RET("IncrementRotationZ", increment, "null");
}

void wxExcelModel3DFormat::ResetModel(bool resetSize)
{
    WXAUTOEXCEL_CALL_METHOD1_RET("ResetModel", resetSize, "null");
}

// ***** class wxExcelModel3DFormat PROPERTIES *****

MsoTriState wxExcelModel3DFormat::GetAutoFit()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("AutoFit", MsoTriState, msoFalse);
}

double wxExcelModel3DFormat::GetCameraPositionX()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("CameraPositionX");
}

double wxExcelModel3DFormat::GetCameraPositionY()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("CameraPositionY");
}

double wxExcelModel3DFormat::GetCameraPositionZ()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("CameraPositionZ");
}

double wxExcelModel3DFormat::GetFieldOfView()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("FieldOfView");
}

double wxExcelModel3DFormat::GetLookAtPointX()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("LookAtPointX");
}

double wxExcelModel3DFormat::GetLookAtPointY()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("LookAtPointY");
}

double wxExcelModel3DFormat::GetLookAtPointZ()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("LookAtPointZ");
}

double wxExcelModel3DFormat::GetRotationX()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("RotationX");
}

double wxExcelModel3DFormat::GetRotationY()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("RotationY");
}

double wxExcelModel3DFormat::GetRotationZ()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("RotationZ");
}

} // namespace wxAutoExcel 
