/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Modified by:
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
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

        [MSDN documentation for Interior.Color](http://msdn.microsoft.com/en-us/library/bb148527).
        */
        wxColour GetColor();

        /**
        Sets the primary color.

        [MSDN documentation for Interior.Color](http://msdn.microsoft.com/en-us/library/bb148527).
        */
        void SetColor(const wxColour& color);

        /**
        Returns the color index into palette.

        [MSDN documentation for Interior.ColorIndex](http://msdn.microsoft.com/en-us/library/bb148529).
        */
        long GetColorIndex();

        /**
        Sets the color index into palette.

        [MSDN documentation for Interior.ColorIndex](http://msdn.microsoft.com/en-us/library/bb148529).
        */
        void SetColorIndex(long colorIndex);


        /**
        Returns the Gradient property of an Interior object of a selection. Since Excel 2007.
        You can call GetLinearGradient() only if GetPattern() returns xlPatternLinearGradient.
        You can call GetRectangularGradient() only if GetPattern() returns xlPatternRectangularGradient.

        [MSDN documentation for Interior.Gradient](http://msdn.microsoft.com/en-us/library/bb243170).
        */
        wxExcelLinearGradient GetLinearGradient();

        wxExcelRectangularGradient GetRectangularGradient();


        /**
        True if Microsoft Excel inverts the pattern in the item when it corresponds to a negative number.

        [MSDN documentation for Interior.InvertIfNegative](http://msdn.microsoft.com/en-us/library/bb148530).
        */
        bool GetInvertIfNegative();

        /**
        True if Microsoft Excel inverts the pattern in the item when it corresponds to a negative number.

        [MSDN documentation for Interior.InvertIfNegative](http://msdn.microsoft.com/en-us/library/bb148530).
        */
        void SetInvertIfNegative(bool invertIfNegative);

        /**
        An xlPattern constant that represents the interior pattern.

        [MSDN documentation for Interior.Pattern](http://msdn.microsoft.com/en-us/library/bb148532).
        */
        XlPattern GetPattern();

        /**
        An xlPattern constant that represents the interior pattern.

        [MSDN documentation for Interior.Pattern](http://msdn.microsoft.com/en-us/library/bb148532).
        */
        void SetPattern(XlPattern pattern);

        /**
        Returns the color of the interior pattern.

        [MSDN documentation for Interior.PatternColor](http://msdn.microsoft.com/en-us/library/bb221414).
        */
        wxColour GetPatternColor();

        /**
        Sets the color of the interior pattern.

        [MSDN documentation for Interior.PatternColor](http://msdn.microsoft.com/en-us/library/bb221414).
        */
        void SetPatternColor(const wxColour& patternColor);

        /**
        Returns the color of the interior pattern as an index into the current color palette, or as one of the following XlColorIndex constants: xlColorIndexAutomatic or xlColorIndexNone.

        [MSDN documentation for Interior.PatternColorIndex](http://msdn.microsoft.com/en-us/library/bb221415).
        */
        long GetPatternColorIndex();

        /**
        Sets the color of the interior pattern as an index into the current color palette, or as one of the following XlColorIndex constants: xlColorIndexAutomatic or xlColorIndexNone.

        [MSDN documentation for Interior.PatternColorIndex](http://msdn.microsoft.com/en-us/library/bb221415).
        */
        void SetPatternColorIndex(long patternColorIndex);

        /**
        Returns a theme color pattern. Since Excel 2007.

        [MSDN documentation for Interior.PatternThemeColor](http://msdn.microsoft.com/en-us/library/bb240144).
        */
        XlThemeColor GetPatternThemeColor();

        /**
        Sets a theme color pattern.

        [MSDN documentation for Interior.PatternThemeColor](http://msdn.microsoft.com/en-us/library/bb240144).
        */
        void SetPatternThemeColor(XlThemeColor patternThemeColor);

        /**
        Returns a tint and shade pattern. Since Excel 2007.

        [MSDN documentation for Interior.PatternTintAndShade](http://msdn.microsoft.com/en-us/library/bb240146).
        */
        double GetPatternTintAndShade();

        /**
        Sets a tint and shade pattern.

        [MSDN documentation for Interior.PatternTintAndShade](http://msdn.microsoft.com/en-us/library/bb240146).
        */
        void SetPatternTintAndShade(double patternTintAndShade);

        /**
        Returns the theme color in the applied color scheme. Since Excel 2007.

        [MSDN documentation for Interior.ThemeColor](http://msdn.microsoft.com/en-us/library/bb214634).
        */
        XlThemeColor GetThemeColor();

        /**
        Sets the theme color in the applied color scheme.

        [MSDN documentation for Interior.ThemeColor](http://msdn.microsoft.com/en-us/library/bb214634).
        */
        void SetThemeColor(XlThemeColor themeColor);

        /**
        Returns a value that lightens or darkens a color. Since Excel 2007.

        [MSDN documentation for Interior.TintAndShade](http://msdn.microsoft.com/en-us/library/bb214636).
        */
        double GetTintAndShade();

        /**
        Sets a value that lightens or darkens a color.

        [MSDN documentation for Interior.TintAndShade](http://msdn.microsoft.com/en-us/library/bb214636).
        */
        void SetTintAndShade(double tintAndShade);


        virtual wxString GetAutoExcelObjectName_() const { return wxS("Interior"); }
    };


} // namespace wxAutoExcel

#endif //_WXAUTOEXCEL_INTERIOR_H
