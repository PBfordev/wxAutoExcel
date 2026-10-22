/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// License:     MIT license
/////////////////////////////////////////////////////////////////////////////


#ifndef _WXAUTOEXCEL_THREEDFORMAT_H
#define _WXAUTOEXCEL_THREEDFORMAT_H

#include "wx/wxAutoExcel_defs.h"

#if WXAUTOEXCEL_USE_SHAPES || WXAUTOEXCEL_USE_CHARTS

#include "wx/wxAutoExcel_object.h"
#include "wx/wxAutoExcel_enums.h"

namespace wxAutoExcel {


    /**
    @brief Represents Microsoft Excel ThreeDFormat object.
    */

   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelThreeDFormat : public wxExcelObject
    {
    public:
        // ***** METHODS *****

        /**
        Changes the rotation of the specified shape horizontally by the specified number of degrees.

        [MSDN documentation for ThreeDFormat.IncrementRotationHorizontal](https://learn.microsoft.com/en-us/office/vba/api/excel.threedformat.incrementrotationhorizontal).
        */
        void IncrementRotationHorizontal(double increment);

        /**
        Changes the rotation of the specified shape vertically by the specified number of degrees.

        [MSDN documentation for ThreeDFormat.IncrementRotationVertical](https://learn.microsoft.com/en-us/office/vba/api/excel.threedformat.incrementrotationvertical).
        */
        void IncrementRotationVertical(double increment);

        /**
        Changes the rotation of the specified shape around the x-axis by the specified number of degrees. Use the RotationX property to set the absolute rotation of the shape around the x-axis.

        [MSDN documentation for ThreeDFormat.IncrementRotationX](https://learn.microsoft.com/en-us/office/vba/api/excel.threedformat.incrementrotationx).
        */
        void IncrementRotationX(double increment);

        /**
        Changes the rotation of the specified shape around the y-axis by the specified number of degrees. Use the RotationY property to set the absolute rotation of the shape around the y-axis.

        [MSDN documentation for ThreeDFormat.IncrementRotationY](https://learn.microsoft.com/en-us/office/vba/api/excel.threedformat.incrementrotationy).
        */
        void IncrementRotationY(double increment);

        /**
        Changes the rotation of the specified shape around the z-axis by the specified number of degrees.

        [MSDN documentation for ThreeDFormat.IncrementRotationZ](https://learn.microsoft.com/en-us/office/vba/api/excel.threedformat.incrementrotationz).
        */
        void IncrementRotationZ(double increment);

        /**
        Resets the extrusion rotation around the x-axis and the y-axis to 0 (zero) so that the front of the extrusion faces forward. This method doesn’t reset the rotation around the z-axis.

        [MSDN documentation for ThreeDFormat.ResetRotation](https://learn.microsoft.com/en-us/office/vba/api/excel.threedformat.resetrotation).
        */
        void ResetRotation();

        /**
        Sets the direction that the extrusion's sweep path takes away from the extruded shape.

        [MSDN documentation for ThreeDFormat.SetExtrusionDirection](https://learn.microsoft.com/en-us/office/vba/api/excel.threedformat.setextrusiondirection).
        */
        void SetExtrusionDirection(MsoPresetExtrusionDirection presetExtrusionDirection);

        /**
        Sets the camera for the specified ThreeDFormat object.

        [MSDN documentation for ThreeDFormat.SetPresetCamera](https://learn.microsoft.com/en-us/office/vba/api/excel.threedformat.setpresetcamera).
        */
        void SetPresetCamera(MsoPresetCamera presetCamera);

        /**
        Sets the preset extrusion format. Each preset extrusion format contains a set of preset values for the various properties of the extrusion.

        [MSDN documentation for ThreeDFormat.SetThreeDFormat](https://learn.microsoft.com/en-us/office/vba/api/excel.threedformat.setthreedformat).
        */
        void SetThreeDFormat(MsoPresetThreeDFormat presetThreeDFormat);

        // ***** PROPERTIES *****


        /**
        Returns the bottom depth when using the bevel effect on a ThreeDFormat object. Read/write Single. Since Excel 2007.

        [MSDN documentation for ThreeDFormat.BevelBottomDepth](https://learn.microsoft.com/en-us/office/vba/api/excel.threedformat.bevelbottomdepth).
        */
        double GetBevelBottomDepth();

        /**
        Sets the bottom depth when using the bevel effect on a ThreeDFormat object. Read/write Single.

        [MSDN documentation for ThreeDFormat.BevelBottomDepth](https://learn.microsoft.com/en-us/office/vba/api/excel.threedformat.bevelbottomdepth).
        */
        void SetBevelBottomDepth(double bevelBottomDepth);

        /**
        Returns a value indicating whether the bottom insert bevel should be raised for a ThreeDFormat object. Read/write Single. Since Excel 2007.

        [MSDN documentation for ThreeDFormat.BevelBottomInset](https://learn.microsoft.com/en-us/office/vba/api/excel.threedformat.bevelbottominset).
        */
        double GetBevelBottomInset();

        /**
        Sets a value indicating whether the bottom insert bevel should be raised for a ThreeDFormat object. Read/write Single.

        [MSDN documentation for ThreeDFormat.BevelBottomInset](https://learn.microsoft.com/en-us/office/vba/api/excel.threedformat.bevelbottominset).
        */
        void SetBevelBottomInset(double bevelBottomInset);

        /**
        Returns the bottom bevel type for a ThreeDFormat object. Read/write MsoBevelType. Since Excel 2007.

        [MSDN documentation for ThreeDFormat.BevelBottomType](https://learn.microsoft.com/en-us/office/vba/api/excel.threedformat.bevelbottomtype).
        */
        MsoBevelType GetBevelBottomType();

        /**
        Sets the bottom bevel type for a ThreeDFormat object. Read/write MsoBevelType.

        [MSDN documentation for ThreeDFormat.BevelBottomType](https://learn.microsoft.com/en-us/office/vba/api/excel.threedformat.bevelbottomtype).
        */
        void SetBevelBottomType(MsoBevelType bevelBottomType);

        /**
        Returns the top depth when using the bevel effect on a ThreeDFormat object. Read/write Single. Since Excel 2007.

        [MSDN documentation for ThreeDFormat.BevelTopDepth](https://learn.microsoft.com/en-us/office/vba/api/excel.threedformat.beveltopdepth).
        */
        double GetBevelTopDepth();

        /**
        Sets the top depth when using the bevel effect on a ThreeDFormat object. Read/write Single.

        [MSDN documentation for ThreeDFormat.BevelTopDepth](https://learn.microsoft.com/en-us/office/vba/api/excel.threedformat.beveltopdepth).
        */
        void SetBevelTopDepth(double bevelTopDepth);

        /**
        Returns a value indicating whether the top insert bevel should be raised for a ThreeDFormat object. Read/write Single. Since Excel 2007.

        [MSDN documentation for ThreeDFormat.BevelTopInset](https://learn.microsoft.com/en-us/office/vba/api/excel.threedformat.beveltopinset).
        */
        double GetBevelTopInset();

        /**
        Sets a value indicating whether the top insert bevel should be raised for a ThreeDFormat object. Read/write Single.

        [MSDN documentation for ThreeDFormat.BevelTopInset](https://learn.microsoft.com/en-us/office/vba/api/excel.threedformat.beveltopinset).
        */
        void SetBevelTopInset(double bevelTopInset);

        /**
        Returns the top Bevel type for a ThreeDFormat object. Read/write MsoBevelType. Since Excel 2007.

        [MSDN documentation for ThreeDFormat.BevelTopType](https://learn.microsoft.com/en-us/office/vba/api/excel.threedformat.beveltoptype).
        */
        MsoBevelType GetBevelTopType();

        /**
        Sets the top Bevel type for a ThreeDFormat object. Read/write MsoBevelType.

        [MSDN documentation for ThreeDFormat.BevelTopType](https://learn.microsoft.com/en-us/office/vba/api/excel.threedformat.beveltoptype).
        */
        void SetBevelTopType(MsoBevelType bevelTopType);

        /**
        Returns the contour color for a ThreeDFormat object. Read-only ColorFormat. Since Excel 2007.

        [MSDN documentation for ThreeDFormat.ContourColor](https://learn.microsoft.com/en-us/office/vba/api/excel.threedformat.contourcolor).
        */
        wxExcelColorFormat GetContourColor();

        /**
        Returns the contour width for a ThreeDFormat object. Read/write Single. Since Excel 2007.

        [MSDN documentation for ThreeDFormat.ContourWidth](https://learn.microsoft.com/en-us/office/vba/api/excel.threedformat.contourwidth).
        */
        double GetContourWidth();

        /**
        Sets the contour width for a ThreeDFormat object. Read/write Single.

        [MSDN documentation for ThreeDFormat.ContourWidth](https://learn.microsoft.com/en-us/office/vba/api/excel.threedformat.contourwidth).
        */
        void SetContourWidth(double contourWidth);

        /**
        Returns a Single value that represents the depth of the shape's extrusion.

        [MSDN documentation for ThreeDFormat.Depth](https://learn.microsoft.com/en-us/office/vba/api/excel.threedformat.depth).
        */
        double GetDepth();

        /**
        Sets a Single value that represents the depth of the shape's extrusion.

        [MSDN documentation for ThreeDFormat.Depth](https://learn.microsoft.com/en-us/office/vba/api/excel.threedformat.depth).
        */
        void SetDepth(double depth);

        /**
        Returns a ColorFormat Represents the color of the shape's extrusion.

        [MSDN documentation for ThreeDFormat.ExtrusionColor](https://learn.microsoft.com/en-us/office/vba/api/excel.threedformat.extrusioncolor).
        */
        wxExcelColorFormat GetExtrusionColor();

        /**
        Returns a value that indicates whether the extrusion color is based on the extruded shape’s fill (the front face of the extrusion) and automatically changes when the shape’s fill changes, or whether the extrusion color is independent of the shape’s fill. Read/write MsoExtrusionColorType.

        [MSDN documentation for ThreeDFormat.ExtrusionColorType](https://learn.microsoft.com/en-us/office/vba/api/excel.threedformat.extrusioncolortype).
        */
        MsoExtrusionColorType GetExtrusionColorType();

        /**
        Sets a value that indicates whether the extrusion color is based on the extruded shape’s fill (the front face of the extrusion) and automatically changes when the shape’s fill changes, or whether the extrusion color is independent of the shape’s fill. Read/write MsoExtrusionColorType.

        [MSDN documentation for ThreeDFormat.ExtrusionColorType](https://learn.microsoft.com/en-us/office/vba/api/excel.threedformat.extrusioncolortype).
        */
        void SetExtrusionColorType(MsoExtrusionColorType extrusionColorType);

        /**
        Returns the angle at which a ThreeDFormat object can be viewed. Read/write Single. Since Excel 2007.

        [MSDN documentation for ThreeDFormat.FieldOfView](https://learn.microsoft.com/en-us/office/vba/api/excel.threedformat.fieldofview).
        */
        double GetFieldOfView();

        /**
        Sets the angle at which a ThreeDFormat object can be viewed. Read/write Single.

        [MSDN documentation for ThreeDFormat.FieldOfView](https://learn.microsoft.com/en-us/office/vba/api/excel.threedformat.fieldofview).
        */
        void SetFieldOfView(double fieldOfView);

        /**
        Returns the angel of the extrusion lights set on a ThreeDFormat object. Read/write Single. Since Excel 2007.

        [MSDN documentation for ThreeDFormat.LightAngle](https://learn.microsoft.com/en-us/office/vba/api/excel.threedformat.lightangle).
        */
        double GetLightAngle();

        /**
        Sets the angel of the extrusion lights set on a ThreeDFormat object. Read/write Single.

        [MSDN documentation for ThreeDFormat.LightAngle](https://learn.microsoft.com/en-us/office/vba/api/excel.threedformat.lightangle).
        */
        void SetLightAngle(double lightAngle);

        /**
        Returns an MsoTriState value that determines whether the extrusion appears in perspective.

        [MSDN documentation for ThreeDFormat.Perspective](https://learn.microsoft.com/en-us/office/vba/api/excel.threedformat.perspective).
        */
        MsoTriState GetPerspective();

        /**
        Sets an MsoTriState value that determines whether the extrusion appears in perspective.

        [MSDN documentation for ThreeDFormat.Perspective](https://learn.microsoft.com/en-us/office/vba/api/excel.threedformat.perspective).
        */
        void SetPerspective(MsoTriState perspective);

        /**
        Returns the extrusion preset camera for a ThreeDFormat object. Read-only MsoPresetCamera. Since Excel 2007.

        [MSDN documentation for ThreeDFormat.PresetCamera](https://learn.microsoft.com/en-us/office/vba/api/excel.threedformat.presetcamera).
        */
        MsoPresetCamera GetPresetCamera();

        /**
        Returns the direction that the extrusion's sweep path takes away from the extruded shape (the front face of the extrusion). Read-only MsoPresetExtrusionDirection.

        [MSDN documentation for ThreeDFormat.PresetExtrusionDirection](https://learn.microsoft.com/en-us/office/vba/api/excel.threedformat.presetextrusiondirection).
        */
        MsoPresetExtrusionDirection GetPresetExtrusionDirection();

        /**
        Returns the extrusion preset lighting for a ThreeDFormat object. Read-only MsoLightRigType. Since Excel 2007.

        [MSDN documentation for ThreeDFormat.PresetLighting](https://learn.microsoft.com/en-us/office/vba/api/excel.threedformat.presetlighting).
        */
        MsoLightRigType GetPresetLighting();

        /**
        Sets the extrusion preset lighting for a ThreeDFormat object. Read-only MsoLightRigType.

        [MSDN documentation for ThreeDFormat.PresetLighting](https://learn.microsoft.com/en-us/office/vba/api/excel.threedformat.presetlighting).
        */
        void SetPresetLighting(MsoLightRigType presetLighting);

        /**
        Returns the position of the light source relative to the extrusion. Read/write MsoPresetLightingDirection.

        [MSDN documentation for ThreeDFormat.PresetLightingDirection](https://learn.microsoft.com/en-us/office/vba/api/excel.threedformat.presetlightingdirection).
        */
        MsoPresetLightingDirection GetPresetLightingDirection();

        /**
        Sets the position of the light source relative to the extrusion. Read/write MsoPresetLightingDirection.

        [MSDN documentation for ThreeDFormat.PresetLightingDirection](https://learn.microsoft.com/en-us/office/vba/api/excel.threedformat.presetlightingdirection).
        */
        void SetPresetLightingDirection(MsoPresetLightingDirection presetLightingDirection);

        /**
        Returns the intensity of the extrusion lighting. Read/write MsoPresetLightingSoftness.

        [MSDN documentation for ThreeDFormat.PresetLightingSoftness](https://learn.microsoft.com/en-us/office/vba/api/excel.threedformat.presetlightingsoftness).
        */
        MsoPresetLightingSoftness GetPresetLightingSoftness();

        /**
        Sets the intensity of the extrusion lighting. Read/write MsoPresetLightingSoftness.

        [MSDN documentation for ThreeDFormat.PresetLightingSoftness](https://learn.microsoft.com/en-us/office/vba/api/excel.threedformat.presetlightingsoftness).
        */
        void SetPresetLightingSoftness(MsoPresetLightingSoftness presetLightingSoftness);

        /**
        Returns the extrusion surface material. Read/write MsoPresetMaterial.

        [MSDN documentation for ThreeDFormat.PresetMaterial](https://learn.microsoft.com/en-us/office/vba/api/excel.threedformat.presetmaterial).
        */
        MsoPresetMaterial GetPresetMaterial();

        /**
        Sets the extrusion surface material. Read/write MsoPresetMaterial.

        [MSDN documentation for ThreeDFormat.PresetMaterial](https://learn.microsoft.com/en-us/office/vba/api/excel.threedformat.presetmaterial).
        */
        void SetPresetMaterial(MsoPresetMaterial presetMaterial);

        /**
        Returns the preset extrusion format. Each preset extrusion format contains a set of preset values for the various properties of the extrusion. Read-only MsoPresetThreeDFormat.

        [MSDN documentation for ThreeDFormat.PresetThreeDFormat](https://learn.microsoft.com/en-us/office/vba/api/excel.threedformat.presetthreedformat).
        */
        MsoPresetThreeDFormat GetPresetThreeDFormat();

        /**
        Returns the project text state for the specified ThreeDFormat object. Read/write MsoTriState. Since Excel 2007.

        [MSDN documentation for ThreeDFormat.ProjectText](https://learn.microsoft.com/en-us/office/vba/api/excel.threedformat.projecttext).
        */
        MsoTriState GetProjectText();

        /**
        Sets the project text state for the specified ThreeDFormat object. Read/write MsoTriState.

        [MSDN documentation for ThreeDFormat.ProjectText](https://learn.microsoft.com/en-us/office/vba/api/excel.threedformat.projecttext).
        */
        void SetProjectText(MsoTriState projectText);

        /**
        Returns the rotation of the extruded shape around the x-axis in degrees. Can be a value from – 90 through 90. A positive value indicates upward rotation; a negative value indicates downward rotation. Read/write Single.

        [MSDN documentation for ThreeDFormat.RotationX](https://learn.microsoft.com/en-us/office/vba/api/excel.threedformat.rotationx).
        */
        double GetRotationX();

        /**
        Sets the rotation of the extruded shape around the x-axis in degrees. Can be a value from – 90 through 90. A positive value indicates upward rotation; a negative value indicates downward rotation. Read/write Single.

        [MSDN documentation for ThreeDFormat.RotationX](https://learn.microsoft.com/en-us/office/vba/api/excel.threedformat.rotationx).
        */
        void SetRotationX(double rotationX);

        /**
        Returns the rotation of the extruded shape around the y-axis in degrees. Can be a value from – 90 through 90. A positive value indicates rotation to the left; a negative value indicates rotation to the right. Read/write Single.

        [MSDN documentation for ThreeDFormat.RotationY](https://learn.microsoft.com/en-us/office/vba/api/excel.threedformat.rotationy).
        */
        double GetRotationY();

        /**
        Sets the rotation of the extruded shape around the y-axis in degrees. Can be a value from – 90 through 90. A positive value indicates rotation to the left; a negative value indicates rotation to the right. Read/write Single.

        [MSDN documentation for ThreeDFormat.RotationY](https://learn.microsoft.com/en-us/office/vba/api/excel.threedformat.rotationy).
        */
        void SetRotationY(double rotationY);

        /**
        Returns the rotation of the extruded shape around the z-axis in degrees. Read/write Single. Since Excel 2007.

        [MSDN documentation for ThreeDFormat.RotationZ](https://learn.microsoft.com/en-us/office/vba/api/excel.threedformat.rotationz).
        */
        double GetRotationZ();

        /**
        Sets the rotation of the extruded shape around the z-axis in degrees. Read/write Single.

        [MSDN documentation for ThreeDFormat.RotationZ](https://learn.microsoft.com/en-us/office/vba/api/excel.threedformat.rotationz).
        */
        void SetRotationZ(double rotationZ);

        /**
        Returns a MsoTriState value that determines whether the object is visible.

        [MSDN documentation for ThreeDFormat.Visible](https://learn.microsoft.com/en-us/office/vba/api/excel.threedformat.visible).
        */
        MsoTriState GetVisible();

        /**
        Sets a MsoTriState value that determines whether the object is visible.

        [MSDN documentation for ThreeDFormat.Visible](https://learn.microsoft.com/en-us/office/vba/api/excel.threedformat.visible).
        */
        void SetVisible(MsoTriState visible);

        /**
        Returns the Z order of the specified ThreeDFormat object. Read/write Single. Since Excel 2007.

        [MSDN documentation for ThreeDFormat.Z](https://learn.microsoft.com/en-us/office/vba/api/excel.threedformat.z).
        */
        double GetZ();

        /**
        Returns the Z order of the specified ThreeDFormat object. Read/write Single.

        [MSDN documentation for ThreeDFormat.Z](https://learn.microsoft.com/en-us/office/vba/api/excel.threedformat.z).
        */
        void SetZ(double z);

        /**
        Returns "ThreeDFormat".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("ThreeDFormat"); }
    };


} // namespace wxAutoExcel

#endif // #if WXAUTOEXCEL_USE_SHAPES || WXAUTOEXCEL_USE_CHARTS

#endif //_WXAUTOEXCEL_THREEDFORMAT_H
