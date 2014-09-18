/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Modified by:
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////


#ifndef _WXAUTOEXCEL_FILLFORMAT_H
#define _WXAUTOEXCEL_FILLFORMAT_H

#include "wx/wxAutoExcel_defs.h"

#if WXAUTOEXCEL_USE_SHAPES || WXAUTOEXCEL_USE_CHARTS

#include "wx/wxAutoExcelObject.h"
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

        [MSDN documentation for FillFormat.OneColorGradient](http://msdn.microsoft.com/en-us/library/bb211758).
        */
        void OneColorGradient(MsoGradientStyle style, long variant, double degree);

        /**
        Sets the specified fill to a pattern.

        [MSDN documentation for FillFormat.Patterned](http://msdn.microsoft.com/en-us/library/bb211761).
        */
        void Patterned(MsoPatternType pattern);

        /**
        Sets the specified fill to a preset gradient.

        [MSDN documentation for FillFormat.PresetGradient](http://msdn.microsoft.com/en-us/library/bb211763).
        */
        void PresetGradient(MsoGradientStyle style, long variant, MsoPresetGradientType presetGradientType);

        /**
        Sets the specified fill format to a preset texture.

        [MSDN documentation for FillFormat.PresetTextured](http://msdn.microsoft.com/en-us/library/bb211766).
        */
        void PresetTextured(MsoPresetTexture presetTexture);

        /**
        Sets the specified fill to a uniform color. Use this method to convert a gradient, textured, patterned, or background fill back to a solid fill.

        [MSDN documentation for FillFormat.Solid](http://msdn.microsoft.com/en-us/library/bb237901).
        */
        void Solid();

        /**
        Sets the specified fill to a two-color gradient.

        [MSDN documentation for FillFormat.TwoColorGradient](http://msdn.microsoft.com/en-us/library/bb237911).
        */
        void TwoColorGradient(MsoGradientStyle style, long variant);

        /**
        Fills the specified shape with an image.

        [MSDN documentation for FillFormat.UserPicture](http://msdn.microsoft.com/en-us/library/bb237920).
        */
        void UserPicture(const wxString& pictureFile);

        /**
        Fills the specified shape with small tiles of an image. If you want to fill the shape with one large image, use the UserPicture method.

        [MSDN documentation for FillFormat.UserTextured](http://msdn.microsoft.com/en-us/library/bb237925).
        */
        void UserTextured(const wxString& pictureFile);

        // ***** PROPERTIES *****


        /**
        Returns a ColorFormat Represents the specified fill background color.

        [MSDN documentation for FillFormat.BackColor](http://msdn.microsoft.com/en-us/library/bb236968).
        */
        wxExcelColorFormat GetBackColor();

        /**
        Returns a ColorFormat Represents the specified foreground fill or solid color.

        [MSDN documentation for FillFormat.ForeColor](http://msdn.microsoft.com/en-us/library/bb236970).
        */
        wxExcelColorFormat GetForeColor();

        /**
        Returns the gradient color type for the specified fill. Read-only MsoGradientColorType.

        [MSDN documentation for FillFormat.GradientColorType](http://msdn.microsoft.com/en-us/library/bb213141).
        */
        MsoGradientColorType GetGradientColorType();

        /**
        Returns the gradient degree of the specified one-color shaded fill as a floating-point value from 0.0 (dark) through 1.0 (light). Read-only Single.

        [MSDN documentation for FillFormat.GradientDegree](http://msdn.microsoft.com/en-us/library/bb213144).
        */
        double GetGradientDegree();

        /**
        Since Excel 2007.

        [MSDN documentation for FillFormat.GradientStops](http://msdn.microsoft.com/en-us/library/bb242651).
        */
        // wxExcelGradientStops GetGradientStops();

        /**
        Returns the gradient style for the specified fill. Read-only MsoGradientStyle.

        [MSDN documentation for FillFormat.GradientStyle](http://msdn.microsoft.com/en-us/library/bb213147).
        */
        MsoGradientStyle GetGradientStyle();

        /**
        Returns the shade variant for the specified fill as an integer value from 1 through 4. The values for this property correspond to the gradient variants (numbered from left to right and from top to bottom) on the Gradient tab in the Fill Effects dialog box. Read-only Long

        [MSDN documentation for FillFormat.GradientVariant](http://msdn.microsoft.com/en-us/library/bb213150).
        */
        long GetGradientVariant();        

        /**
        Returns an MsoPatternType value that represents the fill pattern.

        [MSDN documentation for FillFormat.Pattern](http://msdn.microsoft.com/en-us/library/bb213152).
        */
        MsoPatternType GetPattern();

        /**
        Sets an MsoPatternType value that represents the fill pattern.

        [MSDN documentation for FillFormat.Pattern](http://msdn.microsoft.com/en-us/library/bb213152).
        */
        void SetPattern(MsoPatternType pattern);

        /**
        Returns the preset gradient type for the specified fill. Read-only MsoPresetGradientType.

        [MSDN documentation for FillFormat.PresetGradientType](http://msdn.microsoft.com/en-us/library/bb213154).
        */
        MsoPresetGradientType GetPresetGradientType();

        /**
        Returns the preset texture for the specified fill. Read-only MsoPresetTexture.

        [MSDN documentation for FillFormat.PresetTexture](http://msdn.microsoft.com/en-us/library/bb213158).
        */
        MsoPresetTexture GetPresetTexture();

        /**
        Returns if the fill style should rotate with the object.  Since Excel 2007.

        [MSDN documentation for FillFormat.RotateWithObject](http://msdn.microsoft.com/en-us/library/bb242619).
        */
        MsoTriState GetRotateWithObject();

        /**
        Sets if the fill style should rotate with the object.  Since Excel 2007.

        [MSDN documentation for FillFormat.RotateWithObject](http://msdn.microsoft.com/en-us/library/bb242619).
        */
        void SetRotateWithObject(MsoTriState rotateWithObject);

        /**
        Returns the texture alignment. Since Excel 2007.

        [MSDN documentation for FillFormat.TextureAlignment](http://msdn.microsoft.com/en-us/library/bb242655).
        */
        MsoTextureAlignment GetTextureAlignment();

        /**
        Returns the value for horizontally scaling the text.  Since Excel 2007.

        [MSDN documentation for FillFormat.TextureHorizontalScale](http://msdn.microsoft.com/en-us/library/bb242621).
        */
        double GetTextureHorizontalScale();

        /**
        Sets the value for horizontally scaling the text.  Since Excel 2007.

        [MSDN documentation for FillFormat.TextureHorizontalScale](http://msdn.microsoft.com/en-us/library/bb242621).
        */
        void SetTextureHorizontalScale(double textureHorizontalScale);

        /**
        Returns the name of the custom texture file for the specified fill.

        [MSDN documentation for FillFormat.TextureName](http://msdn.microsoft.com/en-us/library/bb214605).
        */
        wxString GetTextureName();

        /**
        Returns the offset X value. Since Excel 2007.

        [MSDN documentation for FillFormat.TextureOffsetX](http://msdn.microsoft.com/en-us/library/bb242623).
        */
        double GetTextureOffsetX();

        /**
        Returns the offset X value. Since Excel 2007.

        [MSDN documentation for FillFormat.TextureOffsetX](http://msdn.microsoft.com/en-us/library/bb242623).
        */
        void SetTextureOffsetX(double textureOffsetX);

        /**
        Returns the offset Y value. Since Excel 2007.

        [MSDN documentation for FillFormat.TextureOffsetY](http://msdn.microsoft.com/en-us/library/bb242626).
        */
        double GetTextureOffsetY();

        /**
        Returns the offset Y value. Since Excel 2007.

        [MSDN documentation for FillFormat.TextureOffsetY](http://msdn.microsoft.com/en-us/library/bb242626).
        */
        void SetTextureOffsetY(double textureOffsetY);

        /**
        The texture tile.  Since Excel 2007.

        [MSDN documentation for FillFormat.TextureTile](http://msdn.microsoft.com/en-us/library/bb242628).
        */
        MsoTriState GetTextureTile();

        /**
        The texture tile.  Since Excel 2007.

        [MSDN documentation for FillFormat.TextureTile](http://msdn.microsoft.com/en-us/library/bb242628).
        */
        void SetTextureTile(MsoTriState textureTile);

        /**
        Returns the texture type for the specified fill. Read-only MsoTextureType.

        [MSDN documentation for FillFormat.TextureType](http://msdn.microsoft.com/en-us/library/bb214606).
        */
        MsoTextureType GetTextureType();

        /**
        The texture vertical scale. Since Excel 2007.

        [MSDN documentation for FillFormat.TextureVerticalScale](http://msdn.microsoft.com/en-us/library/bb242632).
        */
        double GetTextureVerticalScale();

        /**
        The texture vertical scale. Since Excel 2007.

        [MSDN documentation for FillFormat.TextureVerticalScale](http://msdn.microsoft.com/en-us/library/bb242632).
        */
        void SetTextureVerticalScale(double textureVerticalScale);

        /**
        Returns the degree of transparency of the specified fill as a value from 0.0 (opaque) through 1.0 (clear). 

        [MSDN documentation for FillFormat.Transparency](http://msdn.microsoft.com/en-us/library/bb214608).
        */
        double GetTransparency();

        /**
        Sets the degree of transparency of the specified fill as a value from 0.0 (opaque) through 1.0 (clear).

        [MSDN documentation for FillFormat.Transparency](http://msdn.microsoft.com/en-us/library/bb214608).
        */
        void SetTransparency(double transparency);

        /**
        Returns a MsoFillType value that represents the the fill type.

        [MSDN documentation for FillFormat.Type](http://msdn.microsoft.com/en-us/library/bb214611).
        */
        MsoFillType  GetType();

        /**
        Returns a MsoTriState value that determines whether the object is visible.

        [MSDN documentation for FillFormat.Visible](http://msdn.microsoft.com/en-us/library/bb214613).
        */
        MsoTriState  GetVisible();

        /**
        Sets a MsoTriState value that determines whether the object is visible.

        [MSDN documentation for FillFormat.Visible](http://msdn.microsoft.com/en-us/library/bb214613).
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
