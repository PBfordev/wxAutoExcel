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

        [MSDN documentation for BulletFormat2.Picture](http://msdn.microsoft.com/en-us/library/aa434509.aspx).
        */
        void Picture(const wxString& fileName);

        // ***** PROPERTIES *****

        /**
        Gets or sets the Unicode character value that is used for bullets in the specified text.

        [MSDN documentation for BulletFormat2.Character](http://msdn.microsoft.com/en-us/library/aa434522.aspx).
        */
        long GetCharacter();

        /**
        Gets or sets the Unicode character value that is used for bullets in the specified text.

        [MSDN documentation for BulletFormat2.Character](http://msdn.microsoft.com/en-us/library/aa434522.aspx).
        */
        void SetCharacter(long character);


        /**
        Gets a Font2 Represents character formatting for a BulletFormat2 object.

        [MSDN documentation for BulletFormat2.Font](http://msdn.microsoft.com/en-us/library/aa434524.aspx).
        */
        wxExcelFont2 GetFont();

        /**
        Gets the bullet number of a paragraph.

        [MSDN documentation for BulletFormat2.Number](http://msdn.microsoft.com/en-us/library/aa434525.aspx).
        */
        long GetNumber();

        /**
        Returns the bullet size relative to the size of the first text character in the paragraph.

        [MSDN documentation for BulletFormat2.RelativeSize](http://msdn.microsoft.com/en-us/library/aa434527.aspx).
        */
        double GetRelativeSize();

        /**
        Sets the bullet size relative to the size of the first text character in the paragraph.

        [MSDN documentation for BulletFormat2.RelativeSize](http://msdn.microsoft.com/en-us/library/aa434527.aspx).
        */
        void SetRelativeSize(double relativeSize);

        /**
        Gets or sets the beginning value of a bulleted list.

        [MSDN documentation for BulletFormat2.StartValue](http://msdn.microsoft.com/en-us/library/aa434528.aspx).
        */
        long GetStartValue();

        /**
        Gets or sets the beginning value of a bulleted list.

        [MSDN documentation for BulletFormat2.StartValue](http://msdn.microsoft.com/en-us/library/aa434528.aspx).
        */
        void SetStartValue(long startValue);

        /**
        Returns a constant that represents the style of a bullet.

        [MSDN documentation for BulletFormat2.Style](http://msdn.microsoft.com/en-us/library/aa434529.aspx).
        */
        MsoNumberedBulletStyle GetStyle();

        /**
        Sets a constant that represents the style of a bullet.

        [MSDN documentation for BulletFormat2.Style](http://msdn.microsoft.com/en-us/library/aa434529.aspx).
        */
        void SetStyle(MsoNumberedBulletStyle style);

        /**
        Gets or sets a constant that represents the type of bullet.

        [MSDN documentation for BulletFormat2.Type](http://msdn.microsoft.com/en-us/library/aa434530.aspx).
        */
        MsoBulletType GetType();

        /**
        Gets or sets a constant that represents the type of bullet.

        [MSDN documentation for BulletFormat2.Type](http://msdn.microsoft.com/en-us/library/aa434530.aspx).
        */
        void SetType(MsoBulletType type);

        /**
        Determines whether the specified bullets are set to the color of the first text character in the paragraph.

        [MSDN documentation for BulletFormat2.UseTextColor](http://msdn.microsoft.com/en-us/library/aa434531.aspx).
        */
        MsoTriState GetUseTextColor();

        /**
        Determines whether the specified bullets are set to the color of the first text character in the paragraph.

        [MSDN documentation for BulletFormat2.UseTextColor](http://msdn.microsoft.com/en-us/library/aa434531.aspx).
        */
        void SetUseTextColor(MsoTriState useTextColor);

        /**
        Determines whether the specified bullets are set to the font of the first text character in the paragraph.

        [MSDN documentation for BulletFormat2.UseTextFont](http://msdn.microsoft.com/en-us/library/aa434532.aspx).
        */
        MsoTriState GetUseTextFont();

        /**
        Determines whether the specified bullets are set to the font of the first text character in the paragraph.

        [MSDN documentation for BulletFormat2.UseTextFont](http://msdn.microsoft.com/en-us/library/aa434532.aspx).
        */
        void SetUseTextFont(MsoTriState useTextFont);

        /**
        Gets or sets a value that specifies whether the bullet is visible.

        [MSDN documentation for BulletFormat2.Visible](http://msdn.microsoft.com/en-us/library/aa434533.aspx).
        */
        MsoTriState GetVisible();

        /**
        Gets or sets a value that specifies whether the bullet is visible.

        [MSDN documentation for BulletFormat2.Visible](http://msdn.microsoft.com/en-us/library/aa434533.aspx).
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
