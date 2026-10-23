/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// License:     MIT license
/////////////////////////////////////////////////////////////////////////////


#ifndef _WXAUTOEXCEL_FILLFORMAT_H
#define _WXAUTOEXCEL_FILLFORMAT_H

#include "wx/wxAutoExcel_defs.h"

#if WXAUTOEXCEL_USE_SHAPES || WXAUTOEXCEL_USE_CHARTS

#include "wx/wxAutoExcel_object.h"
#include "wx/wxAutoExcel_enums.h"

namespace wxAutoExcel {

    /**
    @brief Represents Microsoft Excel FillFormat object.
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelFillFormat : public wxExcelObject
    {
    public:
        // ***** METHODS *****

        /**
        Sets the specified fill to a one-color gradient.

        [MSDN documentation for FillFormat.OneColorGradient](https://learn.microsoft.com/en-us/office/vba/api/excel.fillformat.onecolorgradient).
        */
        void OneColorGradient(MsoGradientStyle style, long variant, double degree);

        /**
        Sets the specified fill to a pattern.

        [MSDN documentation for FillFormat.Patterned](https://learn.microsoft.com/en-us/office/vba/api/excel.fillformat.patterned).
        */
        void Patterned(MsoPatternType pattern);

        /**
        Sets the specified fill to a preset gradient.

        [MSDN documentation for FillFormat.PresetGradient](https://learn.microsoft.com/en-us/office/vba/api/excel.fillformat.presetgradient).
        */
        void PresetGradient(MsoGradientStyle style, long variant, MsoPresetGradientType presetGradientType);

        /**
        Sets the specified fill format to a preset texture.

        [MSDN documentation for FillFormat.PresetTextured](https://learn.microsoft.com/en-us/office/vba/api/excel.fillformat.presettextured).
        */
        void PresetTextured(MsoPresetTexture presetTexture);

        /**
        Sets the specified fill to a uniform color. Use this method to convert a gradient, textured, patterned, or background fill back to a solid fill.

        [MSDN documentation for FillFormat.Solid](https://learn.microsoft.com/en-us/office/vba/api/excel.fillformat.solid).
        */
        void Solid();

        /**
        Sets the specified fill to a two-color gradient.

        [MSDN documentation for FillFormat.TwoColorGradient](https://learn.microsoft.com/en-us/office/vba/api/excel.fillformat.twocolorgradient).
        */
        void TwoColorGradient(MsoGradientStyle style, long variant);

        /**
        Fills the specified shape with an image.

        [MSDN documentation for FillFormat.UserPicture](https://learn.microsoft.com/en-us/office/vba/api/excel.fillformat.userpicture).
        */
        void UserPicture(const wxString& pictureFile);

        /**
        Fills the specified shape with small tiles of an image. If you want to fill the shape with one large image, use the UserPicture method.

        [MSDN documentation for FillFormat.UserTextured](https://learn.microsoft.com/en-us/office/vba/api/excel.fillformat.usertextured).
        */
        void UserTextured(const wxString& pictureFile);

        // ***** PROPERTIES *****


        /**
        Returns a ColorFormat Represents the specified fill background color.

        [MSDN documentation for FillFormat.BackColor](https://learn.microsoft.com/en-us/office/vba/api/excel.fillformat.backcolor).
        */
        wxExcelColorFormat GetBackColor();

        /**
        Returns a ColorFormat Represents the specified foreground fill or solid color.

        [MSDN documentation for FillFormat.ForeColor](https://learn.microsoft.com/en-us/office/vba/api/excel.fillformat.forecolor).
        */
        wxExcelColorFormat GetForeColor();

        /**
        Returns the gradient color type for the specified fill. Read-only MsoGradientColorType.

        [MSDN documentation for FillFormat.GradientColorType](https://learn.microsoft.com/en-us/office/vba/api/excel.fillformat.gradientcolortype).
        */
        MsoGradientColorType GetGradientColorType();

        /**
        Returns the gradient degree of the specified one-color shaded fill as a floating-point value from 0.0 (dark) through 1.0 (light). Read-only Single.

        [MSDN documentation for FillFormat.GradientDegree](https://learn.microsoft.com/en-us/office/vba/api/excel.fillformat.gradientdegree).
        */
        double GetGradientDegree();

        /**
        Since Excel 2007.

        [MSDN documentation for FillFormat.GradientStops](https://learn.microsoft.com/en-us/office/vba/api/excel.fillformat.gradientstops).
        */
        // wxExcelGradientStops GetGradientStops();

        /**
        Returns the gradient style for the specified fill. Read-only MsoGradientStyle.

        [MSDN documentation for FillFormat.GradientStyle](https://learn.microsoft.com/en-us/office/vba/api/excel.fillformat.gradientstyle).
        */
        MsoGradientStyle GetGradientStyle();

        /**
        Returns the shade variant for the specified fill as an integer value from 1 through 4. The values for this property correspond to the gradient variants (numbered from left to right and from top to bottom) on the Gradient tab in the Fill Effects dialog box. Read-only Long

        [MSDN documentation for FillFormat.GradientVariant](https://learn.microsoft.com/en-us/office/vba/api/excel.fillformat.gradientvariant).
        */
        long GetGradientVariant();

        /**
        Returns an MsoPatternType value that represents the fill pattern.

        [MSDN documentation for FillFormat.Pattern](https://learn.microsoft.com/en-us/office/vba/api/excel.fillformat.pattern).
        */
        MsoPatternType GetPattern();

        /**
        Returns the preset gradient type for the specified fill. Read-only MsoPresetGradientType.

        [MSDN documentation for FillFormat.PresetGradientType](https://learn.microsoft.com/en-us/office/vba/api/excel.fillformat.presetgradienttype).
        */
        MsoPresetGradientType GetPresetGradientType();

        /**
        Returns the preset texture for the specified fill. Read-only MsoPresetTexture.

        [MSDN documentation for FillFormat.PresetTexture](https://learn.microsoft.com/en-us/office/vba/api/excel.fillformat.presettexture).
        */
        MsoPresetTexture GetPresetTexture();

        /**
        Returns if the fill style should rotate with the object.  Since Excel 2007.

        [MSDN documentation for FillFormat.RotateWithObject](https://learn.microsoft.com/en-us/office/vba/api/excel.fillformat.rotatewithobject).
        */
        MsoTriState GetRotateWithObject();

        /**
        Sets if the fill style should rotate with the object.  Since Excel 2007.

        [MSDN documentation for FillFormat.RotateWithObject](https://learn.microsoft.com/en-us/office/vba/api/excel.fillformat.rotatewithobject).
        */
        void SetRotateWithObject(MsoTriState rotateWithObject);

        /**
        Returns the texture alignment. Since Excel 2007.

        [MSDN documentation for FillFormat.TextureAlignment](https://learn.microsoft.com/en-us/office/vba/api/excel.fillformat.texturealignment).
        */
        MsoTextureAlignment GetTextureAlignment();

        /**
        Sets the alignment (the origin of the coordinate grid) for the tiling of the texture fill.

        [MSDN documentation for FillFormat.TextureAlignment](https://learn.microsoft.com/en-us/office/vba/api/excel.fillformat.texturealignment).
        */
        void SetTextureAlignment(MsoTextureAlignment textureAlignment);

        /**
        Returns the value for horizontally scaling the text.  Since Excel 2007.

        [MSDN documentation for FillFormat.TextureHorizontalScale](https://learn.microsoft.com/en-us/office/vba/api/excel.fillformat.texturehorizontalscale).
        */
        double GetTextureHorizontalScale();

        /**
        Sets the value for horizontally scaling the text.  Since Excel 2007.

        [MSDN documentation for FillFormat.TextureHorizontalScale](https://learn.microsoft.com/en-us/office/vba/api/excel.fillformat.texturehorizontalscale).
        */
        void SetTextureHorizontalScale(double textureHorizontalScale);

        /**
        Returns the name of the custom texture file for the specified fill.

        [MSDN documentation for FillFormat.TextureName](https://learn.microsoft.com/en-us/office/vba/api/excel.fillformat.texturename).
        */
        wxString GetTextureName();

        /**
        Returns the offset X value. Since Excel 2007.

        [MSDN documentation for FillFormat.TextureOffsetX](https://learn.microsoft.com/en-us/office/vba/api/excel.fillformat.textureoffsetx).
        */
        double GetTextureOffsetX();

        /**
        Returns the offset X value. Since Excel 2007.

        [MSDN documentation for FillFormat.TextureOffsetX](https://learn.microsoft.com/en-us/office/vba/api/excel.fillformat.textureoffsetx).
        */
        void SetTextureOffsetX(double textureOffsetX);

        /**
        Returns the offset Y value. Since Excel 2007.

        [MSDN documentation for FillFormat.TextureOffsetY](https://learn.microsoft.com/en-us/office/vba/api/excel.fillformat.textureoffsety).
        */
        double GetTextureOffsetY();

        /**
        Returns the offset Y value. Since Excel 2007.

        [MSDN documentation for FillFormat.TextureOffsetY](https://learn.microsoft.com/en-us/office/vba/api/excel.fillformat.textureoffsety).
        */
        void SetTextureOffsetY(double textureOffsetY);

        /**
        The texture tile.  Since Excel 2007.

        [MSDN documentation for FillFormat.TextureTile](https://learn.microsoft.com/en-us/office/vba/api/excel.fillformat.texturetile).
        */
        MsoTriState GetTextureTile();

        /**
        The texture tile.  Since Excel 2007.

        [MSDN documentation for FillFormat.TextureTile](https://learn.microsoft.com/en-us/office/vba/api/excel.fillformat.texturetile).
        */
        void SetTextureTile(MsoTriState textureTile);

        /**
        Returns the texture type for the specified fill. Read-only MsoTextureType.

        [MSDN documentation for FillFormat.TextureType](https://learn.microsoft.com/en-us/office/vba/api/excel.fillformat.texturetype).
        */
        MsoTextureType GetTextureType();

        /**
        The texture vertical scale. Since Excel 2007.

        [MSDN documentation for FillFormat.TextureVerticalScale](https://learn.microsoft.com/en-us/office/vba/api/excel.fillformat.textureverticalscale).
        */
        double GetTextureVerticalScale();

        /**
        The texture vertical scale. Since Excel 2007.

        [MSDN documentation for FillFormat.TextureVerticalScale](https://learn.microsoft.com/en-us/office/vba/api/excel.fillformat.textureverticalscale).
        */
        void SetTextureVerticalScale(double textureVerticalScale);

        /**
        Returns the degree of transparency of the specified fill as a value from 0.0 (opaque) through 1.0 (clear).

        [MSDN documentation for FillFormat.Transparency](https://learn.microsoft.com/en-us/office/vba/api/excel.fillformat.transparency).
        */
        double GetTransparency();

        /**
        Sets the degree of transparency of the specified fill as a value from 0.0 (opaque) through 1.0 (clear).

        [MSDN documentation for FillFormat.Transparency](https://learn.microsoft.com/en-us/office/vba/api/excel.fillformat.transparency).
        */
        void SetTransparency(double transparency);

        /**
        Returns a MsoFillType value that represents the the fill type.

        [MSDN documentation for FillFormat.Type](https://learn.microsoft.com/en-us/office/vba/api/excel.fillformat.type).
        */
        MsoFillType  GetType();

        /**
        Returns a MsoTriState value that determines whether the object is visible.

        [MSDN documentation for FillFormat.Visible](https://learn.microsoft.com/en-us/office/vba/api/excel.fillformat.visible).
        */
        MsoTriState  GetVisible();

        /**
        Sets a MsoTriState value that determines whether the object is visible.

        [MSDN documentation for FillFormat.Visible](https://learn.microsoft.com/en-us/office/vba/api/excel.fillformat.visible).
        */
        void SetVisible(MsoTriState  visible);


        /**
        Returns "FillFormat".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("FillFormat"); }
    };


} // namespace wxAutoExcel


#endif // #if WXAUTOEXCEL_USE_SHAPES || WXAUTOEXCEL_USE_CHARTS

#endif //_WXAUTOEXCEL_FILLFORMAT_H
