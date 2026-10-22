/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// License:     MIT license
/////////////////////////////////////////////////////////////////////////////


#ifndef _WXAUTOEXCEL_INTERIOR_H
#define _WXAUTOEXCEL_INTERIOR_H

#include "wx/wxAutoExcel_defs.h"

#include "wx/wxAutoExcel_object.h"
#include "wx/wxAutoExcel_enums.h"

class WXDLLIMPEXP_FWD_CORE wxColour;

namespace wxAutoExcel {


    /**
    @brief Represents Microsoft Excel Interior.
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelInterior : public wxExcelObject
    {
    public:
        // ***** PROPERTIES *****

        /**
        Returns the primary color.

        [MSDN documentation for Interior.Color](https://learn.microsoft.com/en-us/office/vba/api/excel.interior.color).
        */
        wxColour GetColor();

        /**
        Sets the primary color.

        [MSDN documentation for Interior.Color](https://learn.microsoft.com/en-us/office/vba/api/excel.interior.color).
        */
        void SetColor(const wxColour& color);

        /**
        Returns the color index into palette.

        [MSDN documentation for Interior.ColorIndex](https://learn.microsoft.com/en-us/office/vba/api/excel.interior.colorindex).
        */
        long GetColorIndex();

        /**
        Sets the color index into palette.

        [MSDN documentation for Interior.ColorIndex](https://learn.microsoft.com/en-us/office/vba/api/excel.interior.colorindex).
        */
        void SetColorIndex(long colorIndex);


        /**
        Returns the Gradient property of an Interior object of a selection. Since Excel 2007.
        You can call GetLinearGradient() only if GetPattern() returns xlPatternLinearGradient.
        You can call GetRectangularGradient() only if GetPattern() returns xlPatternRectangularGradient.

        [MSDN documentation for Interior.Gradient](https://learn.microsoft.com/en-us/office/vba/api/excel.interior.gradient).
        */
        wxExcelLinearGradient GetLinearGradient();

        wxExcelRectangularGradient GetRectangularGradient();


        /**
        True if Microsoft Excel inverts the pattern in the item when it corresponds to a negative number.

        [MSDN documentation for Interior.InvertIfNegative](https://learn.microsoft.com/en-us/office/vba/api/excel.interior.invertifnegative).
        */
        bool GetInvertIfNegative();

        /**
        True if Microsoft Excel inverts the pattern in the item when it corresponds to a negative number.

        [MSDN documentation for Interior.InvertIfNegative](https://learn.microsoft.com/en-us/office/vba/api/excel.interior.invertifnegative).
        */
        void SetInvertIfNegative(bool invertIfNegative);

        /**
        An xlPattern constant that represents the interior pattern.

        [MSDN documentation for Interior.Pattern](https://learn.microsoft.com/en-us/office/vba/api/excel.interior.pattern).
        */
        XlPattern GetPattern();

        /**
        An xlPattern constant that represents the interior pattern.

        [MSDN documentation for Interior.Pattern](https://learn.microsoft.com/en-us/office/vba/api/excel.interior.pattern).
        */
        void SetPattern(XlPattern pattern);

        /**
        Returns the color of the interior pattern.

        [MSDN documentation for Interior.PatternColor](https://learn.microsoft.com/en-us/office/vba/api/excel.interior.patterncolor).
        */
        wxColour GetPatternColor();

        /**
        Sets the color of the interior pattern.

        [MSDN documentation for Interior.PatternColor](https://learn.microsoft.com/en-us/office/vba/api/excel.interior.patterncolor).
        */
        void SetPatternColor(const wxColour& patternColor);

        /**
        Returns the color of the interior pattern as an index into the current color palette, or as one of the following XlColorIndex constants: xlColorIndexAutomatic or xlColorIndexNone.

        [MSDN documentation for Interior.PatternColorIndex](https://learn.microsoft.com/en-us/office/vba/api/excel.interior.patterncolorindex).
        */
        long GetPatternColorIndex();

        /**
        Sets the color of the interior pattern as an index into the current color palette, or as one of the following XlColorIndex constants: xlColorIndexAutomatic or xlColorIndexNone.

        [MSDN documentation for Interior.PatternColorIndex](https://learn.microsoft.com/en-us/office/vba/api/excel.interior.patterncolorindex).
        */
        void SetPatternColorIndex(long patternColorIndex);

        /**
        Returns a theme color pattern. Since Excel 2007.

        [MSDN documentation for Interior.PatternThemeColor](https://learn.microsoft.com/en-us/office/vba/api/excel.interior.patternthemecolor).
        */
        XlThemeColor GetPatternThemeColor();

        /**
        Sets a theme color pattern.

        [MSDN documentation for Interior.PatternThemeColor](https://learn.microsoft.com/en-us/office/vba/api/excel.interior.patternthemecolor).
        */
        void SetPatternThemeColor(XlThemeColor patternThemeColor);

        /**
        Returns a tint and shade pattern. Since Excel 2007.

        [MSDN documentation for Interior.PatternTintAndShade](https://learn.microsoft.com/en-us/office/vba/api/excel.interior.patterntintandshade).
        */
        double GetPatternTintAndShade();

        /**
        Sets a tint and shade pattern.

        [MSDN documentation for Interior.PatternTintAndShade](https://learn.microsoft.com/en-us/office/vba/api/excel.interior.patterntintandshade).
        */
        void SetPatternTintAndShade(double patternTintAndShade);

        /**
        Returns the theme color in the applied color scheme. Since Excel 2007.

        [MSDN documentation for Interior.ThemeColor](https://learn.microsoft.com/en-us/office/vba/api/excel.interior.themecolor).
        */
        XlThemeColor GetThemeColor();

        /**
        Sets the theme color in the applied color scheme.

        [MSDN documentation for Interior.ThemeColor](https://learn.microsoft.com/en-us/office/vba/api/excel.interior.themecolor).
        */
        void SetThemeColor(XlThemeColor themeColor);

        /**
        Returns a value that lightens or darkens a color. Since Excel 2007.

        [MSDN documentation for Interior.TintAndShade](https://learn.microsoft.com/en-us/office/vba/api/excel.interior.tintandshade).
        */
        double GetTintAndShade();

        /**
        Sets a value that lightens or darkens a color.

        [MSDN documentation for Interior.TintAndShade](https://learn.microsoft.com/en-us/office/vba/api/excel.interior.tintandshade).
        */
        void SetTintAndShade(double tintAndShade);


        virtual wxString GetAutoExcelObjectName_() const { return wxS("Interior"); }
    };


} // namespace wxAutoExcel

#endif //_WXAUTOEXCEL_INTERIOR_H
