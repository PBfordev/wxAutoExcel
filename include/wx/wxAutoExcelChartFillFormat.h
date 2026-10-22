/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// License:     MIT license
/////////////////////////////////////////////////////////////////////////////


#ifndef _WXAUTOEXCEL_CHARTFILLFORMAT_H
#define _WXAUTOEXCEL_CHARTFILLFORMAT_H

#include "wx/wxAutoExcel_defs.h"

#if WXAUTOEXCEL_USE_CHARTS

#include "wx/wxAutoExcel_object.h"
#include "wx/wxAutoExcel_enums.h"

namespace wxAutoExcel {

    /**
    @brief Represents Microsoft Excel ChartFillFormat object.
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelChartFillFormat : public wxExcelObject
    {
    public:
        // ***** METHODS *****

        /**
        Sets the specified fill to a one-color gradient.

        [MSDN documentation for ChartFillFormat.OneColorGradient](https://learn.microsoft.com/en-us/previous-versions/office/developer/office-2007/bb148226%28v=office.12%29).
        */
        void OneColorGradient(MsoGradientStyle style, long variant, double degree);

        /**
        Sets the specified fill to a pattern.

        [MSDN documentation for ChartFillFormat.Patterned](https://learn.microsoft.com/en-us/previous-versions/office/developer/office-2007/bb148231%28v=office.12%29).
        */
        void Patterned(MsoPatternType pattern);

        /**
        Sets the specified fill to a preset gradient.

        [MSDN documentation for ChartFillFormat.PresetGradient](https://learn.microsoft.com/en-us/previous-versions/office/developer/office-2007/bb148235%28v=office.12%29).
        */
        void PresetGradient(MsoGradientStyle style, long variant, MsoPresetGradientType presetGradientType);

        /**
        Sets the specified fill format to a preset texture.

        [MSDN documentation for ChartFillFormat.PresetTextured](https://learn.microsoft.com/en-us/previous-versions/office/developer/office-2007/bb148241%28v=office.12%29).
        */
        void PresetTextured(MsoPresetTexture presetTexture);

        /**
        Sets the specified fill to a uniform color. Use this method to convert a gradient, textured, patterned, or background fill back to a solid fill.

        [MSDN documentation for ChartFillFormat.Solid](https://learn.microsoft.com/en-us/previous-versions/office/developer/office-2007/bb213879%28v=office.12%29).
        */
        void Solid();

        /**
        Sets the specified fill to a two-color gradient.

        [MSDN documentation for ChartFillFormat.TwoColorGradient](https://learn.microsoft.com/en-us/previous-versions/office/developer/office-2007/bb213883%28v=office.12%29).
        */
        void TwoColorGradient(MsoGradientStyle style, long variant);

        /**
        Fills the specified shape with an image.

        [MSDN documentation for ChartFillFormat.UserPicture](https://learn.microsoft.com/en-us/previous-versions/office/developer/office-2007/bb213894%28v=office.12%29).
        */
        void UserPicture(const wxString& pictureFile = wxEmptyString, XlChartPictureType* pictureFormat = NULL,
                         double* pictureStackUnit = NULL, XlChartPicturePlacement* picturePlacement = NULL);

        /**
        Fills the specified shape with small tiles of an image. If you want to fill the shape with one large image, use the UserPicture method.

        [MSDN documentation for ChartFillFormat.UserTextured](https://learn.microsoft.com/en-us/previous-versions/office/developer/office-2007/bb213900%28v=office.12%29).
        */
        void UserTextured(const wxString& textureFile);

        // ***** PROPERTIES *****

        /**
        Returns a ChartColorFormat object that represents the specified fill background color.

        [MSDN documentation for ChartFillFormat.BackColor](https://learn.microsoft.com/en-us/previous-versions/office/developer/office-2007/bb179430%28v=office.12%29).
        */
        wxExcelChartColorFormat GetBackColor();

        /**
        Returns a ChartColorFormat object that represents the specified foreground fill or solid color.

        [MSDN documentation for ChartFillFormat.ForeColor](https://learn.microsoft.com/en-us/previous-versions/office/developer/office-2007/bb179433%28v=office.12%29).
        */
        wxExcelChartColorFormat GetForeColor();

        /**
        Returns the gradient color type for the specified fill. Read-only MsoGradientColorType.

        [MSDN documentation for ChartFillFormat.GradientColorType](https://learn.microsoft.com/en-us/previous-versions/office/developer/office-2007/bb179436%28v=office.12%29).
        */
        MsoGradientColorType GetGradientColorType();

        /**
        Returns the gradient degree of the specified one-color shaded fill as a floating-point value from 0.0 (dark) through 1.0 (light). Read-only Single.

        [MSDN documentation for ChartFillFormat.GradientDegree](https://learn.microsoft.com/en-us/previous-versions/office/developer/office-2007/bb179437%28v=office.12%29).
        */
        double GetGradientDegree();

        /**
        Returns the gradient style for the specified fill. Read-only MsoGradientStyle.

        [MSDN documentation for ChartFillFormat.GradientStyle](https://learn.microsoft.com/en-us/previous-versions/office/developer/office-2007/bb179440%28v=office.12%29).
        */
        MsoGradientStyle GetGradientStyle();

        /**
        Returns the shade variant for the specified fill as an integer value from 1 through 4. The values for this property correspond to the gradient variants (numbered from left to right and from top to bottom) on the Gradient tab in the Fill Effects dialog box. Read-only Long

        [MSDN documentation for ChartFillFormat.GradientVariant](https://learn.microsoft.com/en-us/previous-versions/office/developer/office-2007/bb179442%28v=office.12%29).
        */
        long GetGradientVariant();

        /**
        Returns an MsoPatternType value that represents the fill pattern.

        [MSDN documentation for ChartFillFormat.Pattern](https://learn.microsoft.com/en-us/previous-versions/office/developer/office-2007/bb179446%28v=office.12%29).
        */
        MsoPatternType GetPattern();

        /**
        Returns the preset gradient type for the specified fill. Read-only MsoPresetGradientType.

        [MSDN documentation for ChartFillFormat.PresetGradientType](https://learn.microsoft.com/en-us/previous-versions/office/developer/office-2007/bb179447%28v=office.12%29).
        */
        MsoPresetGradientType GetPresetGradientType();

        /**
        Returns the preset texture for the specified fill. Read-only MsoPresetTexture.

        [MSDN documentation for ChartFillFormat.PresetTexture](https://learn.microsoft.com/en-us/previous-versions/office/developer/office-2007/bb179449%28v=office.12%29).
        */
        MsoPresetTexture GetPresetTexture();

        /**
        Returns the name of the custom texture file for the specified fill.

        [MSDN documentation for ChartFillFormat.TextureName](https://learn.microsoft.com/en-us/previous-versions/office/developer/office-2007/bb148856%28v=office.12%29).
        */
        wxString GetTextureName();

        /**
        Returns the texture type for the specified fill. Read-only MsoTextureType.

        [MSDN documentation for ChartFillFormat.TextureType](https://learn.microsoft.com/en-us/previous-versions/office/developer/office-2007/bb148861%28v=office.12%29).
        */
        MsoTextureType GetTextureType();

        /**
        Returns a MsoFillType value that represents the the fill type.

        [MSDN documentation for ChartFillFormat.Type](https://learn.microsoft.com/en-us/previous-versions/office/developer/office-2007/bb148867%28v=office.12%29).
        */
        MsoFillType GetType();

        /**
        Returns a MsoTriState value that determines whether the object is visible.

        [MSDN documentation for ChartFillFormat.Visible](https://learn.microsoft.com/en-us/previous-versions/office/developer/office-2007/bb148870%28v=office.12%29).
        */
        MsoTriState GetVisible();

        /**
        Sets a MsoTriState value that determines whether the object is visible.

        [MSDN documentation for ChartFillFormat.Visible](https://learn.microsoft.com/en-us/previous-versions/office/developer/office-2007/bb148870%28v=office.12%29).
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
