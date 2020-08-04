/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Modified by:
// Copyright:   (c) 2020 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////

#ifndef _WXAUTOEXCEL_MODEL3DFORMAT_H
#define _WXAUTOEXCEL_MODEL3DFORMAT_H

#include "wx/wxAutoExcel_defs.h"

#include "wx/wxAutoExcel_object.h"
#include "wx/wxAutoExcel_enums.h"

namespace wxAutoExcel
{

/**
    @brief Represents the properties of a 3D model shape.
*/
class WXDLLIMPEXP_WXAUTOEXCEL wxExcelModel3DFormat : public wxExcelObject
{
public:
    // ***** METHODS *****

    /**
    Changes the rotation of the specified shape around the x-axis by the specified number of degrees.

    [Excel VBA documentation for Model3DFormat.IncrementRotationX](https://docs.microsoft.com/en-us/office/vba/api/excel.model3dformat.incrementrotationx)
    */
    void IncrementRotationX(double increment);

    /**
    Changes the rotation of the specified shape around the y-axis by the specified number of degrees.

    [Excel VBA documentation for Model3DFormat.IncrementRotationY](https://docs.microsoft.com/en-us/office/vba/api/excel.model3dformat.incrementrotationy)
    */
    void IncrementRotationY(double increment);

    /**
    Changes the rotation of the specified shape around the z-axis by the specified number of degrees.

    [Excel VBA documentation for Model3DFormat.IncrementRotationZ](https://docs.microsoft.com/en-us/office/vba/api/excel.model3dformat.incrementrotationz)
    */
    void IncrementRotationZ(double increment);

    /**
    Restores 3D model properties back to default settings. 

    [Excel VBA documentation for Model3DFormat.ResetModel](https://docs.microsoft.com/en-us/office/vba/api/excel.model3dformat.resetmodel)
    */
    void ResetModel(bool resetSize = false);

    // ***** PROPERTIES *****

    /**
    Whether AutoFit is enabled for the model.

    [Excel VBA documentation for Model3DFormat.AutoFit](https://docs.microsoft.com/en-us/office/vba/api/excel.model3dformat.autofit)
    */
    MsoTriState GetAutoFit();

    /**
    The x-coordinate of a 3D model object's camera position.

    [Excel VBA documentation for Model3DFormat.CameraPositionX](https://docs.microsoft.com/en-us/office/vba/api/excel.model3dformat.camerapositionx)
    */
    double GetCameraPositionX();

    /**
    The y-coordinate of a 3D model object's camera position.

    [Excel VBA documentation for Model3DFormat.CameraPositionY](https://docs.microsoft.com/en-us/office/vba/api/excel.model3dformat.camerapositiony)
    */
    double GetCameraPositionY();

    /**
    The z-coordinate of a 3D model object's camera position.

    [Excel VBA documentation for Model3DFormat.CameraPositionZ](https://docs.microsoft.com/en-us/office/vba/api/excel.model3dformat.camerapositionz)
    */
    double GetCameraPositionZ();

    /**
    The field-of-view angle of a 3D model object's camera, expressed in degrees.

    [Excel VBA documentation for Model3DFormat.FieldOfView](https://docs.microsoft.com/en-us/office/vba/api/excel.model3dformat.fieldofview)
    */
    double GetFieldOfView();

    /**
    The x-coordinate of a 3D model object's camera look-at position.

    [Excel VBA documentation for Model3DFormat.LookAtPointX](https://docs.microsoft.com/en-us/office/vba/api/excel.model3dformat.lookatpointx)
    */
    double GetLookAtPointX();

    /**
    The y-coordinate of a 3D model object's camera look-at position.

    [Excel VBA documentation for Model3DFormat.LookAtPointY](https://docs.microsoft.com/en-us/office/vba/api/excel.model3dformat.lookatpointy)
    */
    double GetLookAtPointY();

    /**
    The z-coordinate of a 3D model object's camera look-at position.

    [Excel VBA documentation for Model3DFormat.LookAtPointZ](https://docs.microsoft.com/en-us/office/vba/api/excel.model3dformat.lookatpointz)
    */
    double GetLookAtPointZ();

    /**
    The x-angle of a 3D model object's rotation.

    [Excel VBA documentation for Model3DFormat.RotationX](https://docs.microsoft.com/en-us/office/vba/api/excel.model3dformat.rotationx)
    */
    double GetRotationX();

    /**
    The z-angle of a 3D model object's rotation.

    [Excel VBA documentation for Model3DFormat.RotationY](https://docs.microsoft.com/en-us/office/vba/api/excel.model3dformat.rotationy)
    */
    double GetRotationY();

    /**
    The z-angle of a 3D model object's rotation.

    [Excel VBA documentation for Model3DFormat.RotationZ](https://docs.microsoft.com/en-us/office/vba/api/excel.model3dformat.rotationz)
    */
    double GetRotationZ();

    /**
    Returns "Model3DFormat".
    */
    virtual wxString GetAutoExcelObjectName_() const { return wxS("Model3DFormat"); }

}; // class wxExcelModel3DFormat

} // namespace wxAutoExcel 

#endif // #ifndef _WXAUTOEXCEL_MODEL3DFORMAT_H
