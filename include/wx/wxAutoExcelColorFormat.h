/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// License:     MIT license
/////////////////////////////////////////////////////////////////////////////


#ifndef _WXAUTOEXCEL_COLORFORMAT_H
#define _WXAUTOEXCEL_COLORFORMAT_H

#include "wx/wxAutoExcel_defs.h"

#if WXAUTOEXCEL_USE_SHAPES || WXAUTOEXCEL_USE_CHARTS

#include "wx/wxAutoExcel_object.h"
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
        The luminosity of the specified object. Since Excel 2010.

        [MSDN documentation for ColorFormat.Brightness](https://learn.microsoft.com/en-us/office/vba/api/excel.colorformat.brightness).
        */
        double GetBrightness();

        /**
        Returns a color that is mapped to the theme color scheme. Since Excel 2007.

        [MSDN documentation for ColorFormat.ObjectThemeColor](https://learn.microsoft.com/en-us/office/vba/api/excel.colorformat.objectthemecolor).
        */
        MsoThemeColorIndex GetObjectThemeColor();

        /**
        Sets a color that is mapped to the theme color scheme

        [MSDN documentation for ColorFormat.ObjectThemeColor](https://learn.microsoft.com/en-us/office/vba/api/excel.colorformat.objectthemecolor).
        */
        void SetObjectThemeColor(MsoThemeColorIndex objectThemeColor);


        /**
        Returns a Long value that represents the red-green-blue value of the specified color.

        [MSDN documentation for ColorFormat.RGB](https://learn.microsoft.com/en-us/office/vba/api/excel.colorformat.rgb).
        */
        wxColour GetRGB();

        /**
        Sets a Long value that represents the red-green-blue value of the specified color.

        [MSDN documentation for ColorFormat.RGB](https://learn.microsoft.com/en-us/office/vba/api/excel.colorformat.rgb).
        */
        void SetRGB(const wxColour& RGB);

        /**
        Returns an Integer value that represents the color of a Color object, as an index in the current color scheme.

        [MSDN documentation for ColorFormat.SchemeColor](https://learn.microsoft.com/en-us/office/vba/api/excel.colorformat.schemecolor).
        */
        long GetSchemeColor();

        /**
        Sets an Integer value that represents the color of a Color object, as an index in the current color scheme.

        [MSDN documentation for ColorFormat.SchemeColor](https://learn.microsoft.com/en-us/office/vba/api/excel.colorformat.schemecolor).
        */
        void SetSchemeColor(long schemeColor);

        /**
        Returns a value that lightens or darkens a color.

        [MSDN documentation for ColorFormat.TintAndShade](https://learn.microsoft.com/en-us/office/vba/api/excel.colorformat.tintandshade).
        */
        double GetTintAndShade();

        /**
        Sets a value that lightens or darkens a color.

        [MSDN documentation for ColorFormat.TintAndShade](https://learn.microsoft.com/en-us/office/vba/api/excel.colorformat.tintandshade).
        */
        void SetTintAndShade(double tintAndShade);

        /**
        Returns a MsoColorType value that represents the color format type.

        [MSDN documentation for ColorFormat.Type](https://learn.microsoft.com/en-us/office/vba/api/excel.colorformat.type).
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
