/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Modified by:
// RCS-ID:      $Id: $
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////


#ifndef _WXAUTOEXCEL_CHARTFILLFORMAT_H
#define _WXAUTOEXCEL_CHARTFILLFORMAT_H

#include "wx/wxAutoExcel_defs.h"

#if WXAUTOEXCEL_USE_CHARTS

#include "wx/wxAutoExcelObject.h"
#include "wx/wxAutoExcel_enums.h"

namespace wxAutoExcel {

    /**
    Represents Microsoft Excel ChartFillFormat object.
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelChartFillFormat : public wxExcelObject
    {
    public:
        // ***** METHODS *****

        /**
        Sets the specified fill to a one-color gradient.

        [MSDN documentation for ChartFillFormat.OneColorGradient](http://msdn.microsoft.com/en-us/library/bb148226).
        */
        void OneColorGradient(MsoGradientStyle style, long variant, double degree);

        /**
        Sets the specified fill to a pattern.

        [MSDN documentation for ChartFillFormat.Patterned](http://msdn.microsoft.com/en-us/library/bb148231).
        */
        void Patterned(MsoPatternType pattern);

        /**
        Sets the specified fill to a preset gradient.

        [MSDN documentation for ChartFillFormat.PresetGradient](http://msdn.microsoft.com/en-us/library/bb148235).
        */
        void PresetGradient(MsoGradientStyle style, long variant, MsoPresetGradientType presetGradientType);

        /**
        Sets the specified fill format to a preset texture.

        [MSDN documentation for ChartFillFormat.PresetTextured](http://msdn.microsoft.com/en-us/library/bb148241).
        */
        void PresetTextured(MsoPresetTexture presetTexture);

        /**
        Sets the specified fill to a uniform color. Use this method to convert a gradient, textured, patterned, or background fill back to a solid fill.

        [MSDN documentation for ChartFillFormat.Solid](http://msdn.microsoft.com/en-us/library/bb213879).
        */
        void Solid();

        /**
        Sets the specified fill to a two-color gradient.

        [MSDN documentation for ChartFillFormat.TwoColorGradient](http://msdn.microsoft.com/en-us/library/bb213883).
        */
        void TwoColorGradient(MsoGradientStyle style, long variant);

        /**
        Fills the specified shape with an image.

        [MSDN documentation for ChartFillFormat.UserPicture](http://msdn.microsoft.com/en-us/library/bb213894).
        */
        void UserPicture(const wxString& pictureFile = wxEmptyString, XlChartPictureType* pictureFormat = NULL, 
                         double* pictureStackUnit = NULL, XlChartPicturePlacement* picturePlacement = NULL);

        /**
        Fills the specified shape with small tiles of an image. If you want to fill the shape with one large image, use the UserPicture method.

        [MSDN documentation for ChartFillFormat.UserTextured](http://msdn.microsoft.com/en-us/library/bb213900).
        */
        void UserTextured(const wxString& textureFile);

        // ***** PROPERTIES *****

        /**
        Returns a ChartColorFormat object that represents the specified fill background color.

        [MSDN documentation for ChartFillFormat.BackColor](http://msdn.microsoft.com/en-us/library/bb179430).
        */
        wxExcelChartColorFormat GetBackColor();

        /**
        Returns a ChartColorFormat object that represents the specified foreground fill or solid color.

        [MSDN documentation for ChartFillFormat.ForeColor](http://msdn.microsoft.com/en-us/library/bb179433).
        */
        wxExcelChartColorFormat GetForeColor();

        /**
        Returns the gradient color type for the specified fill. Read-only MsoGradientColorType.

        [MSDN documentation for ChartFillFormat.GradientColorType](http://msdn.microsoft.com/en-us/library/bb179436).
        */
        MsoGradientColorType GetGradientColorType();

        /**
        Returns the gradient degree of the specified one-color shaded fill as a floating-point value from 0.0 (dark) through 1.0 (light). Read-only Single.

        [MSDN documentation for ChartFillFormat.GradientDegree](http://msdn.microsoft.com/en-us/library/bb179437).
        */
        double GetGradientDegree();

        /**
        Returns the gradient style for the specified fill. Read-only MsoGradientStyle.

        [MSDN documentation for ChartFillFormat.GradientStyle](http://msdn.microsoft.com/en-us/library/bb179440).
        */
        MsoGradientStyle GetGradientStyle();

        /**
        Returns the shade variant for the specified fill as an integer value from 1 through 4. The values for this property correspond to the gradient variants (numbered from left to right and from top to bottom) on the Gradient tab in the Fill Effects dialog box. Read-only Long

        [MSDN documentation for ChartFillFormat.GradientVariant](http://msdn.microsoft.com/en-us/library/bb179442).
        */
        long GetGradientVariant();

        /**
        Returns an MsoPatternType value that represents the fill pattern.

        [MSDN documentation for ChartFillFormat.Pattern](http://msdn.microsoft.com/en-us/library/bb179446).
        */
        MsoPatternType GetPattern();

        /**
        Sets an MsoPatternType value that represents the fill pattern.

        [MSDN documentation for ChartFillFormat.Pattern](http://msdn.microsoft.com/en-us/library/bb179446).
        */
        void SetPattern(MsoPatternType pattern);

        /**
        Returns the preset gradient type for the specified fill. Read-only MsoPresetGradientType.

        [MSDN documentation for ChartFillFormat.PresetGradientType](http://msdn.microsoft.com/en-us/library/bb179447).
        */
        MsoPresetGradientType GetPresetGradientType();

        /**
        Returns the preset texture for the specified fill. Read-only MsoPresetTexture.

        [MSDN documentation for ChartFillFormat.PresetTexture](http://msdn.microsoft.com/en-us/library/bb179449).
        */
        MsoPresetTexture GetPresetTexture();

        /**
        Returns the name of the custom texture file for the specified fill.

        [MSDN documentation for ChartFillFormat.TextureName](http://msdn.microsoft.com/en-us/library/bb148856).
        */
        wxString GetTextureName();

        /**
        Returns the texture type for the specified fill. Read-only MsoTextureType.

        [MSDN documentation for ChartFillFormat.TextureType](http://msdn.microsoft.com/en-us/library/bb148861).
        */
        MsoTextureType GetTextureType();

        /**
        Returns a MsoFillType value that represents the the fill type.

        [MSDN documentation for ChartFillFormat.Type](http://msdn.microsoft.com/en-us/library/bb148867).
        */
        MsoFillType GetType();

        /**
        Returns a MsoTriState value that determines whether the object is visible.

        [MSDN documentation for ChartFillFormat.Visible](http://msdn.microsoft.com/en-us/library/bb148870).
        */
        MsoTriState GetVisible();

        /**
        Sets a MsoTriState value that determines whether the object is visible.

        [MSDN documentation for ChartFillFormat.Visible](http://msdn.microsoft.com/en-us/library/bb148870).
        */
        void SetVisible(MsoTriState visible);

        /**
        Returns "ChartFillFormat".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("ChartFillFormat"); }
    };


} // namespace wxAutoExcel

#endif // #if WXAUTOEXCEL_USE_CHARTS

#endif // _WXAUTOEXCEL_CHARTFILLFORMAT_H
