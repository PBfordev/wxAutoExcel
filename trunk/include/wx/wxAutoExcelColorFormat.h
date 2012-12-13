/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Modified by:
// RCS-ID:      $Id: $
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////


#ifndef _WXAUTOEXCEL_COLORFORMAT_H
#define _WXAUTOEXCEL_COLORFORMAT_H

#include "wx/wxAutoExcel_defs.h"

#if WXAUTOEXCEL_USE_SHAPES || WXAUTOEXCEL_USE_CHARTS

#include "wx/wxAutoExcelObject.h"
#include "wx/wxAutoExcel_enums.h"

class WXDLLIMPEXP_FWD_CORE wxColour;

namespace wxAutoExcel {

    /**
    @brief Represents Microsoft Excel ColorFormat object.
    */

   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelColorFormat : public wxExcelObject
    {
    public:
        // ***** PROPERTIES *****

        /**
        Returns a color that is mapped to the theme color scheme Since Excel 2007.

        [MSDN documentation for ColorFormat.ObjectThemeColor](http://msdn.microsoft.com/en-us/library/bb239965).
        */
        MsoThemeColorIndex GetObjectThemeColor();

        /**
        Sets a color that is mapped to the theme color scheme

        [MSDN documentation for ColorFormat.ObjectThemeColor](http://msdn.microsoft.com/en-us/library/bb239965).
        */
        void SetObjectThemeColor(MsoThemeColorIndex objectThemeColor);


        /**
        Returns a Long value that represents the red-green-blue value of the specified color.

        [MSDN documentation for ColorFormat.RGB](http://msdn.microsoft.com/en-us/library/bb214500).
        */
        wxColour GetRGB();

        /**
        Sets a Long value that represents the red-green-blue value of the specified color.

        [MSDN documentation for ColorFormat.RGB](http://msdn.microsoft.com/en-us/library/bb214500).
        */
        void SetRGB(const wxColour& RGB);

        /**
        Returns an Integer value that represents the color of a Color object, as an index in the current color scheme.

        [MSDN documentation for ColorFormat.SchemeColor](http://msdn.microsoft.com/en-us/library/bb214505).
        */
        long GetSchemeColor();

        /**
        Sets an Integer value that represents the color of a Color object, as an index in the current color scheme.

        [MSDN documentation for ColorFormat.SchemeColor](http://msdn.microsoft.com/en-us/library/bb214505).
        */
        void SetSchemeColor(long schemeColor);

        /**
        Returns a value that lightens or darkens a color.

        [MSDN documentation for ColorFormat.TintAndShade](http://msdn.microsoft.com/en-us/library/bb214508).
        */
        double GetTintAndShade();

        /**
        Sets a value that lightens or darkens a color.

        [MSDN documentation for ColorFormat.TintAndShade](http://msdn.microsoft.com/en-us/library/bb214508).
        */
        void SetTintAndShade(double tintAndShade);

        /**
        Returns a MsoColorType value that represents the color format type.

        [MSDN documentation for ColorFormat.Type](http://msdn.microsoft.com/en-us/library/bb214512).
        */
        MsoColorType GetType();

        /**
        Returns "ColorFormat".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("ColorFormat"); }
    };


} // namespace wxAutoExcel

#endif // #if WXAUTOEXCEL_USE_SHAPES || WXAUTOEXCEL_USE_CHARTS

#endif //_WXAUTOEXCEL_COLORFORMAT_H
