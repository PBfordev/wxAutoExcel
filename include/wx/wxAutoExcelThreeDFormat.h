/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Modified by:
// RCS-ID:      $Id: $
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////


#ifndef _WXAUTOEXCEL_THREEDFORMAT_H
#define _WXAUTOEXCEL_THREEDFORMAT_H

#include "wx/wxAutoExcel_defs.h"

#if WXAUTOEXCEL_USE_SHAPES || WXAUTOEXCEL_USE_CHARTS

#include "wx/wxAutoExcelObject.h"
#include "wx/wxAutoExcel_enums.h"

namespace wxAutoExcel {


    /**
    @brief Represents Microsoft Excel ThreeDFormat object.
    */

   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelThreeDFormat : public wxExcelObject
    {
        // ***** METHODS *****

        /**
        Changes the rotation of the specified shape horizontally by the specified number of degrees.

        [MSDN documentation for ThreeDFormat.IncrementRotationHorizontal](http://msdn.microsoft.com/en-us/library/bb238895).
        */
        void IncrementRotationHorizontal(double increment);

        /**
        Changes the rotation of the specified shape vertically by the specified number of degrees.

        [MSDN documentation for ThreeDFormat.IncrementRotationVertical](http://msdn.microsoft.com/en-us/library/bb238896).
        */
        void IncrementRotationVertical(double increment);

        /**
        Changes the rotation of the specified shape around the x-axis by the specified number of degrees. Use the RotationX property to set the absolute rotation of the shape around the x-axis.

        [MSDN documentation for ThreeDFormat.IncrementRotationX](http://msdn.microsoft.com/en-us/library/bb209944).
        */
        void IncrementRotationX(double increment);

        /**
        Changes the rotation of the specified shape around the y-axis by the specified number of degrees. Use the RotationY property to set the absolute rotation of the shape around the y-axis.

        [MSDN documentation for ThreeDFormat.IncrementRotationY](http://msdn.microsoft.com/en-us/library/bb209947).
        */
        void IncrementRotationY(double increment);

        /**
        Changes the rotation of the specified shape around the z-axis by the specified number of degrees.

        [MSDN documentation for ThreeDFormat.IncrementRotationZ](http://msdn.microsoft.com/en-us/library/bb238897).
        */
        void IncrementRotationZ(double increment);

        /**
        Resets the extrusion rotation around the x-axis and the y-axis to 0 (zero) so that the front of the extrusion faces forward. This method doesn’t reset the rotation around the z-axis.

        [MSDN documentation for ThreeDFormat.ResetRotation](http://msdn.microsoft.com/en-us/library/bb177970).
        */
        void ResetRotation();

        /**
        Sets the direction that the extrusion's sweep path takes away from the extruded shape.

        [MSDN documentation for ThreeDFormat.SetExtrusionDirection](http://msdn.microsoft.com/en-us/library/bb178054).
        */
        void SetExtrusionDirection(MsoPresetExtrusionDirection presetExtrusionDirection);

        /**
        Sets the camera for the specified ThreeDFormat object.

        [MSDN documentation for ThreeDFormat.SetPresetCamera](http://msdn.microsoft.com/en-us/library/bb238898).
        */
        void SetPresetCamera(MsoPresetCamera presetCamera);

        /**
        Sets the preset extrusion format. Each preset extrusion format contains a set of preset values for the various properties of the extrusion.

        [MSDN documentation for ThreeDFormat.SetThreeDFormat](http://msdn.microsoft.com/en-us/library/bb178096).
        */
        void SetThreeDFormat(MsoPresetThreeDFormat presetThreeDFormat);

        // ***** PROPERTIES *****
        

        /**
        Returns the bottom depth when using the bevel effect on a ThreeDFormat object. Read/write Single. Since Excel 2007.

        [MSDN documentation for ThreeDFormat.BevelBottomDepth](http://msdn.microsoft.com/en-us/library/bb215880).
        */
        double GetBevelBottomDepth();

        /**
        Sets the bottom depth when using the bevel effect on a ThreeDFormat object. Read/write Single.

        [MSDN documentation for ThreeDFormat.BevelBottomDepth](http://msdn.microsoft.com/en-us/library/bb215880).
        */
        void SetBevelBottomDepth(double bevelBottomDepth);

        /**
        Returns a value indicating whether the bottom insert bevel should be raised for a ThreeDFormat object. Read/write Single. Since Excel 2007.

        [MSDN documentation for ThreeDFormat.BevelBottomInset](http://msdn.microsoft.com/en-us/library/bb215883).
        */
        double GetBevelBottomInset();

        /**
        Sets a value indicating whether the bottom insert bevel should be raised for a ThreeDFormat object. Read/write Single.

        [MSDN documentation for ThreeDFormat.BevelBottomInset](http://msdn.microsoft.com/en-us/library/bb215883).
        */
        void SetBevelBottomInset(double bevelBottomInset);

        /**
        Returns the bottom bevel type for a ThreeDFormat object. Read/write MsoBevelType. Since Excel 2007.

        [MSDN documentation for ThreeDFormat.BevelBottomType](http://msdn.microsoft.com/en-us/library/bb215887).
        */
        MsoBevelType GetBevelBottomType();

        /**
        Sets the bottom bevel type for a ThreeDFormat object. Read/write MsoBevelType.

        [MSDN documentation for ThreeDFormat.BevelBottomType](http://msdn.microsoft.com/en-us/library/bb215887).
        */
        void SetBevelBottomType(MsoBevelType bevelBottomType);

        /**
        Returns the top depth when using the bevel effect on a ThreeDFormat object. Read/write Single. Since Excel 2007.

        [MSDN documentation for ThreeDFormat.BevelTopDepth](http://msdn.microsoft.com/en-us/library/bb215892).
        */
        double GetBevelTopDepth();

        /**
        Sets the top depth when using the bevel effect on a ThreeDFormat object. Read/write Single.

        [MSDN documentation for ThreeDFormat.BevelTopDepth](http://msdn.microsoft.com/en-us/library/bb215892).
        */
        void SetBevelTopDepth(double bevelTopDepth);

        /**
        Returns a value indicating whether the top insert bevel should be raised for a ThreeDFormat object. Read/write Single. Since Excel 2007.

        [MSDN documentation for ThreeDFormat.BevelTopInset](http://msdn.microsoft.com/en-us/library/bb215895).
        */
        double GetBevelTopInset();

        /**
        Sets a value indicating whether the top insert bevel should be raised for a ThreeDFormat object. Read/write Single.

        [MSDN documentation for ThreeDFormat.BevelTopInset](http://msdn.microsoft.com/en-us/library/bb215895).
        */
        void SetBevelTopInset(double bevelTopInset);

        /**
        Returns the top Bevel type for a ThreeDFormat object. Read/write MsoBevelType. Since Excel 2007.

        [MSDN documentation for ThreeDFormat.BevelTopType](http://msdn.microsoft.com/en-us/library/bb215898).
        */
        MsoBevelType GetBevelTopType();

        /**
        Sets the top Bevel type for a ThreeDFormat object. Read/write MsoBevelType.

        [MSDN documentation for ThreeDFormat.BevelTopType](http://msdn.microsoft.com/en-us/library/bb215898).
        */
        void SetBevelTopType(MsoBevelType bevelTopType);

        /**
        Returns the contour color for a ThreeDFormat object. Read-only ColorFormat. Since Excel 2007.

        [MSDN documentation for ThreeDFormat.ContourColor](http://msdn.microsoft.com/en-us/library/bb215899).
        */
        wxExcelColorFormat GetContourColor();

        /**
        Returns the contour width for a ThreeDFormat object. Read/write Single. Since Excel 2007.

        [MSDN documentation for ThreeDFormat.ContourWidth](http://msdn.microsoft.com/en-us/library/bb215905).
        */
        double GetContourWidth();

        /**
        Sets the contour width for a ThreeDFormat object. Read/write Single.

        [MSDN documentation for ThreeDFormat.ContourWidth](http://msdn.microsoft.com/en-us/library/bb215905).
        */
        void SetContourWidth(double contourWidth);

        /**
        Returns a Single value that represents the depth of the shape's extrusion.

        [MSDN documentation for ThreeDFormat.Depth](http://msdn.microsoft.com/en-us/library/bb238247).
        */
        double GetDepth();

        /**
        Sets a Single value that represents the depth of the shape's extrusion.

        [MSDN documentation for ThreeDFormat.Depth](http://msdn.microsoft.com/en-us/library/bb238247).
        */
        void SetDepth(double depth);

        /**
        Returns a ColorFormat Represents the color of the shape's extrusion.

        [MSDN documentation for ThreeDFormat.ExtrusionColor](http://msdn.microsoft.com/en-us/library/bb208490).
        */
        wxExcelColorFormat GetExtrusionColor();

        /**
        Returns a value that indicates whether the extrusion color is based on the extruded shape’s fill (the front face of the extrusion) and automatically changes when the shape’s fill changes, or whether the extrusion color is independent of the shape’s fill. Read/write MsoExtrusionColorType.

        [MSDN documentation for ThreeDFormat.ExtrusionColorType](http://msdn.microsoft.com/en-us/library/bb208492).
        */
        MsoExtrusionColorType GetExtrusionColorType();

        /**
        Sets a value that indicates whether the extrusion color is based on the extruded shape’s fill (the front face of the extrusion) and automatically changes when the shape’s fill changes, or whether the extrusion color is independent of the shape’s fill. Read/write MsoExtrusionColorType.

        [MSDN documentation for ThreeDFormat.ExtrusionColorType](http://msdn.microsoft.com/en-us/library/bb208492).
        */
        void SetExtrusionColorType(MsoExtrusionColorType extrusionColorType);

        /**
        Returns the angle at which a ThreeDFormat object can be viewed. Read/write Single. Since Excel 2007.

        [MSDN documentation for ThreeDFormat.FieldOfView](http://msdn.microsoft.com/en-us/library/bb215913).
        */
        double GetFieldOfView();

        /**
        Sets the angle at which a ThreeDFormat object can be viewed. Read/write Single.

        [MSDN documentation for ThreeDFormat.FieldOfView](http://msdn.microsoft.com/en-us/library/bb215913).
        */
        void SetFieldOfView(double fieldOfView);

        /**
        Returns the angel of the extrusion lights set on a ThreeDFormat object. Read/write Single. Since Excel 2007.

        [MSDN documentation for ThreeDFormat.LightAngle](http://msdn.microsoft.com/en-us/library/bb242114).
        */
        double GetLightAngle();

        /**
        Sets the angel of the extrusion lights set on a ThreeDFormat object. Read/write Single.

        [MSDN documentation for ThreeDFormat.LightAngle](http://msdn.microsoft.com/en-us/library/bb242114).
        */
        void SetLightAngle(double lightAngle);        

        /**
        Returns an MsoTriState value that determines whether the extrusion appears in perspective.

        [MSDN documentation for ThreeDFormat.Perspective](http://msdn.microsoft.com/en-us/library/bb238256).
        */
        MsoTriState GetPerspective();

        /**
        Sets an MsoTriState value that determines whether the extrusion appears in perspective.

        [MSDN documentation for ThreeDFormat.Perspective](http://msdn.microsoft.com/en-us/library/bb238256).
        */
        void SetPerspective(MsoTriState perspective);

        /**
        Returns the extrusion preset camera for a ThreeDFormat object. Read-only MsoPresetCamera. Since Excel 2007.

        [MSDN documentation for ThreeDFormat.PresetCamera](http://msdn.microsoft.com/en-us/library/bb215917).
        */
        MsoPresetCamera GetPresetCamera();

        /**
        Returns the direction that the extrusion's sweep path takes away from the extruded shape (the front face of the extrusion). Read-only MsoPresetExtrusionDirection.

        [MSDN documentation for ThreeDFormat.PresetExtrusionDirection](http://msdn.microsoft.com/en-us/library/bb208980).
        */
        MsoPresetExtrusionDirection GetPresetExtrusionDirection();

        /**
        Returns the extrusion preset lighting for a ThreeDFormat object. Read-only MsoLightRigType. Since Excel 2007.

        [MSDN documentation for ThreeDFormat.PresetLighting](http://msdn.microsoft.com/en-us/library/bb215923).
        */
        MsoLightRigType GetPresetLighting();

        /**
        Sets the extrusion preset lighting for a ThreeDFormat object. Read-only MsoLightRigType.

        [MSDN documentation for ThreeDFormat.PresetLighting](http://msdn.microsoft.com/en-us/library/bb215923).
        */
        void SetPresetLighting(MsoLightRigType presetLighting);

        /**
        Returns the position of the light source relative to the extrusion. Read/write MsoPresetLightingDirection.

        [MSDN documentation for ThreeDFormat.PresetLightingDirection](http://msdn.microsoft.com/en-us/library/bb208982).
        */
        MsoPresetLightingDirection GetPresetLightingDirection();

        /**
        Sets the position of the light source relative to the extrusion. Read/write MsoPresetLightingDirection.

        [MSDN documentation for ThreeDFormat.PresetLightingDirection](http://msdn.microsoft.com/en-us/library/bb208982).
        */
        void SetPresetLightingDirection(MsoPresetLightingDirection presetLightingDirection);

        /**
        Returns the intensity of the extrusion lighting. Read/write MsoPresetLightingSoftness.

        [MSDN documentation for ThreeDFormat.PresetLightingSoftness](http://msdn.microsoft.com/en-us/library/bb208984).
        */
        MsoPresetLightingSoftness GetPresetLightingSoftness();

        /**
        Sets the intensity of the extrusion lighting. Read/write MsoPresetLightingSoftness.

        [MSDN documentation for ThreeDFormat.PresetLightingSoftness](http://msdn.microsoft.com/en-us/library/bb208984).
        */
        void SetPresetLightingSoftness(MsoPresetLightingSoftness presetLightingSoftness);

        /**
        Returns the extrusion surface material. Read/write MsoPresetMaterial.

        [MSDN documentation for ThreeDFormat.PresetMaterial](http://msdn.microsoft.com/en-us/library/bb208986).
        */
        MsoPresetMaterial GetPresetMaterial();

        /**
        Sets the extrusion surface material. Read/write MsoPresetMaterial.

        [MSDN documentation for ThreeDFormat.PresetMaterial](http://msdn.microsoft.com/en-us/library/bb208986).
        */
        void SetPresetMaterial(MsoPresetMaterial presetMaterial);

        /**
        Returns the preset extrusion format. Each preset extrusion format contains a set of preset values for the various properties of the extrusion. Read-only MsoPresetThreeDFormat.

        [MSDN documentation for ThreeDFormat.PresetThreeDFormat](http://msdn.microsoft.com/en-us/library/bb208992).
        */
        MsoPresetThreeDFormat GetPresetThreeDFormat();

        /**
        Returns the project text state for the specified ThreeDFormat object. Read/write MsoTriState. Since Excel 2007.

        [MSDN documentation for ThreeDFormat.ProjectText](http://msdn.microsoft.com/en-us/library/bb215924).
        */
        MsoTriState GetProjectText();

        /**
        Sets the project text state for the specified ThreeDFormat object. Read/write MsoTriState.

        [MSDN documentation for ThreeDFormat.ProjectText](http://msdn.microsoft.com/en-us/library/bb215924).
        */
        void SetProjectText(MsoTriState projectText);

        /**
        Returns the rotation of the extruded shape around the x-axis in degrees. Can be a value from – 90 through 90. A positive value indicates upward rotation; a negative value indicates downward rotation. Read/write Single.

        [MSDN documentation for ThreeDFormat.RotationX](http://msdn.microsoft.com/en-us/library/bb221533).
        */
        double GetRotationX();

        /**
        Sets the rotation of the extruded shape around the x-axis in degrees. Can be a value from – 90 through 90. A positive value indicates upward rotation; a negative value indicates downward rotation. Read/write Single.

        [MSDN documentation for ThreeDFormat.RotationX](http://msdn.microsoft.com/en-us/library/bb221533).
        */
        void SetRotationX(double rotationX);

        /**
        Returns the rotation of the extruded shape around the y-axis in degrees. Can be a value from – 90 through 90. A positive value indicates rotation to the left; a negative value indicates rotation to the right. Read/write Single.

        [MSDN documentation for ThreeDFormat.RotationY](http://msdn.microsoft.com/en-us/library/bb221538).
        */
        double GetRotationY();

        /**
        Sets the rotation of the extruded shape around the y-axis in degrees. Can be a value from – 90 through 90. A positive value indicates rotation to the left; a negative value indicates rotation to the right. Read/write Single.

        [MSDN documentation for ThreeDFormat.RotationY](http://msdn.microsoft.com/en-us/library/bb221538).
        */
        void SetRotationY(double rotationY);

        /**
        Returns the rotation of the extruded shape around the z-axis in degrees. Read/write Single. Since Excel 2007.

        [MSDN documentation for ThreeDFormat.RotationZ](http://msdn.microsoft.com/en-us/library/bb215927).
        */
        double GetRotationZ();

        /**
        Sets the rotation of the extruded shape around the z-axis in degrees. Read/write Single.

        [MSDN documentation for ThreeDFormat.RotationZ](http://msdn.microsoft.com/en-us/library/bb215927).
        */
        void SetRotationZ(double rotationZ);

        /**
        Returns a MsoTriState value that determines whether the object is visible.

        [MSDN documentation for ThreeDFormat.Visible](http://msdn.microsoft.com/en-us/library/bb215159).
        */
        MsoTriState GetVisible();

        /**
        Sets a MsoTriState value that determines whether the object is visible.

        [MSDN documentation for ThreeDFormat.Visible](http://msdn.microsoft.com/en-us/library/bb215159).
        */
        void SetVisible(MsoTriState visible);

        /**
        Returns the Z order of the specified ThreeDFormat object. Read/write Single. Since Excel 2007.

        [MSDN documentation for ThreeDFormat.Z](http://msdn.microsoft.com/en-us/library/bb215929).
        */
        double GetZ();

        /**
        Returns the Z order of the specified ThreeDFormat object. Read/write Single.

        [MSDN documentation for ThreeDFormat.Z](http://msdn.microsoft.com/en-us/library/bb215929).
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
