/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// License:     MIT license
/////////////////////////////////////////////////////////////////////////////


#ifndef _WXAUTOEXCEL_BULLETFORMAT2_H
#define _WXAUTOEXCEL_BULLETFORMAT2_H

#include "wx/wxAutoExcel_defs.h"

#if WXAUTOEXCEL_USE_SHAPES || WXAUTOEXCEL_USE_CHARTS

#include "wx/wxAutoExcel_object.h"
#include "wx/wxAutoExcel_enums.h"

namespace wxAutoExcel {

    /**
    @brief Represents Microsoft Excel BulletFormat2 object.
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelBulletFormat2: public wxExcelObject
    {
    public:
        // ***** METHODS *****

        /**
        Sets the graphics file to be used for bullets in a bulleted list.

        [MSDN documentation for BulletFormat2.Picture](https://learn.microsoft.com/en-us/office/vba/api/office.bulletformat2.picture).
        */
        void Picture(const wxString& fileName);

        // ***** PROPERTIES *****

        /**
        Gets or sets the Unicode character value that is used for bullets in the specified text.

        [MSDN documentation for BulletFormat2.Character](https://learn.microsoft.com/en-us/office/vba/api/office.bulletformat2.character).
        */
        long GetCharacter();

        /**
        Gets or sets the Unicode character value that is used for bullets in the specified text.

        [MSDN documentation for BulletFormat2.Character](https://learn.microsoft.com/en-us/office/vba/api/office.bulletformat2.character).
        */
        void SetCharacter(long character);


        /**
        Gets a Font2 Represents character formatting for a BulletFormat2 object.

        [MSDN documentation for BulletFormat2.Font](https://learn.microsoft.com/en-us/office/vba/api/office.bulletformat2.font).
        */
        wxExcelFont2 GetFont();

        /**
        Gets the bullet number of a paragraph.

        [MSDN documentation for BulletFormat2.Number](https://learn.microsoft.com/en-us/office/vba/api/office.bulletformat2.number).
        */
        long GetNumber();

        /**
        Returns the bullet size relative to the size of the first text character in the paragraph.

        [MSDN documentation for BulletFormat2.RelativeSize](https://learn.microsoft.com/en-us/office/vba/api/office.bulletformat2.relativesize).
        */
        double GetRelativeSize();

        /**
        Sets the bullet size relative to the size of the first text character in the paragraph.

        [MSDN documentation for BulletFormat2.RelativeSize](https://learn.microsoft.com/en-us/office/vba/api/office.bulletformat2.relativesize).
        */
        void SetRelativeSize(double relativeSize);

        /**
        Gets or sets the beginning value of a bulleted list.

        [MSDN documentation for BulletFormat2.StartValue](https://learn.microsoft.com/en-us/office/vba/api/office.bulletformat2.startvalue).
        */
        long GetStartValue();

        /**
        Gets or sets the beginning value of a bulleted list.

        [MSDN documentation for BulletFormat2.StartValue](https://learn.microsoft.com/en-us/office/vba/api/office.bulletformat2.startvalue).
        */
        void SetStartValue(long startValue);

        /**
        Returns a constant that represents the style of a bullet.

        [MSDN documentation for BulletFormat2.Style](https://learn.microsoft.com/en-us/office/vba/api/office.bulletformat2.style).
        */
        MsoNumberedBulletStyle GetStyle();

        /**
        Sets a constant that represents the style of a bullet.

        [MSDN documentation for BulletFormat2.Style](https://learn.microsoft.com/en-us/office/vba/api/office.bulletformat2.style).
        */
        void SetStyle(MsoNumberedBulletStyle style);

        /**
        Gets or sets a constant that represents the type of bullet.

        [MSDN documentation for BulletFormat2.Type](https://learn.microsoft.com/en-us/office/vba/api/office.bulletformat2.type).
        */
        MsoBulletType GetType();

        /**
        Gets or sets a constant that represents the type of bullet.

        [MSDN documentation for BulletFormat2.Type](https://learn.microsoft.com/en-us/office/vba/api/office.bulletformat2.type).
        */
        void SetType(MsoBulletType type);

        /**
        Determines whether the specified bullets are set to the color of the first text character in the paragraph.

        [MSDN documentation for BulletFormat2.UseTextColor](https://learn.microsoft.com/en-us/office/vba/api/office.bulletformat2.usetextcolor).
        */
        MsoTriState GetUseTextColor();

        /**
        Determines whether the specified bullets are set to the color of the first text character in the paragraph.

        [MSDN documentation for BulletFormat2.UseTextColor](https://learn.microsoft.com/en-us/office/vba/api/office.bulletformat2.usetextcolor).
        */
        void SetUseTextColor(MsoTriState useTextColor);

        /**
        Determines whether the specified bullets are set to the font of the first text character in the paragraph.

        [MSDN documentation for BulletFormat2.UseTextFont](https://learn.microsoft.com/en-us/office/vba/api/office.bulletformat2.usetextfont).
        */
        MsoTriState GetUseTextFont();

        /**
        Determines whether the specified bullets are set to the font of the first text character in the paragraph.

        [MSDN documentation for BulletFormat2.UseTextFont](https://learn.microsoft.com/en-us/office/vba/api/office.bulletformat2.usetextfont).
        */
        void SetUseTextFont(MsoTriState useTextFont);

        /**
        Gets or sets a value that specifies whether the bullet is visible.

        [MSDN documentation for BulletFormat2.Visible](https://learn.microsoft.com/en-us/office/vba/api/office.bulletformat2.visible).
        */
        MsoTriState GetVisible();

        /**
        Gets or sets a value that specifies whether the bullet is visible.

        [MSDN documentation for BulletFormat2.Visible](https://learn.microsoft.com/en-us/office/vba/api/office.bulletformat2.visible).
        */
        void SetVisible(MsoTriState visible);


        /**
        Returns "BulletFormat2".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("BulletFormat2"); }
    };


} // namespace wxAutoExcel

#endif // #if WXAUTOEXCEL_USE_SHAPES || WXAUTOEXCEL_USE_CHARTS

#endif //_WXAUTOEXCEL_BULLETFORMAT2_H
