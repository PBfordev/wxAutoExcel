/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Modified by:
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////


#ifndef _WXAUTOEXCEL_FORMATCOLOR_H
#define _WXAUTOEXCEL_FORMATCOLOR_H

#include "wx/wxAutoExcel_defs.h"

#if WXAUTOEXCEL_USE_CONDFORMAT

#include "wx/wxAutoExcel_object.h"
#include "wx/wxAutoExcel_enums.h"

namespace wxAutoExcel {
    /**
    @brief Represents a Microsoft Excel FormatColor object.
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelFormatColor : public wxExcelObject
    {
    public:
        // ***** PROPERTIES *****

        /**
        Returns the fill color associated with a threshold for a data bar or color scale conditional formatting rule. Since Excel 2007.

        [MSDN documentation for FormatColor.Color](http://msdn.microsoft.com/en-us/library/bb224454.aspx).
        */
        wxColour GetColor();

        /**
        Sets the fill color associated with a threshold for a data bar or color scale conditional formatting rule. Since Excel 2007.

        [MSDN documentation for FormatColor.Color](http://msdn.microsoft.com/en-us/library/bb224454.aspx).
        */
        void SetColor(const wxColour& color);

        /**
        Returns one of the constants of the XlColorIndex enumeration specifying if the fill color is expressed as an index value into the current color palette. Since Excel 2007.

        [MSDN documentation for FormatColor.ColorIndex](http://msdn.microsoft.com/en-us/library/bb224459.aspx).
        */
        XlColorIndex GetColorIndex();

        /**
        Sets one of the constants of the XlColorIndex enumeration specifying if the fill color is expressed as an index value into the current color palette. Since Excel 2007.

        [MSDN documentation for FormatColor.ColorIndex](http://msdn.microsoft.com/en-us/library/bb224459.aspx).
        */
        void SetColorIndex(XlColorIndex colorIndex);

        /**
        Returns of sets one of the constants of the XlThemeColor enumeration specifying the theme color used in a threshold of a data bar or color scale conditional format. Since Excel 2007.

        [MSDN documentation for FormatColor.ThemeColor](http://msdn.microsoft.com/en-us/library/bb224461.aspx).
        */
        XlThemeColor GetThemeColor();

        /**
        Returns of sets one of the constants of the XlThemeColor enumeration specifying the theme color used in a threshold of a data bar or color scale conditional format. Since Excel 2007.

        [MSDN documentation for FormatColor.ThemeColor](http://msdn.microsoft.com/en-us/library/bb224461.aspx).
        */
        void SetThemeColor(XlThemeColor themeColor);

        /**
        Returns a Single that lightens or darkens the fill color of a cell for a thresholds of a data bar or color scale conditional formatting rule. Since Excel 2007.

        [MSDN documentation for FormatColor.TintAndShade](http://msdn.microsoft.com/en-us/library/bb224465.aspx).
        */
        double GetTintAndShade();

        /**
        Sets a Single that lightens or darkens the fill color of a cell for a thresholds of a data bar or color scale conditional formatting rule. Since Excel 2007.

        [MSDN documentation for FormatColor.TintAndShade](http://msdn.microsoft.com/en-us/library/bb224465.aspx).
        */
        void SetTintAndShade(double tintAndShade);

        
        /**
        Returns "FormatColor".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("FormatColor"); }    
    };


} // namespace wxAutoExcel

#endif // WXAUTOEXCEL_USE_CONDFORMAT

#endif //_WXAUTOEXCEL_FORMATCOLOR_H
