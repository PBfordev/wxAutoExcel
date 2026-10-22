/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// License:     MIT license
/////////////////////////////////////////////////////////////////////////////


#ifndef _WXAUTOEXCEL_TAB_H
#define _WXAUTOEXCEL_TAB_H

#include "wx/wxAutoExcel_defs.h"

#include "wx/wxAutoExcel_object.h"
#include "wx/wxAutoExcel_enums.h"

class WXDLLIMPEXP_FWD_CORE wxColour;

namespace wxAutoExcel {


    /**
    @brief Represents Microsoft Excel Sheet Tab.
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelTab : public wxExcelObject
    {
    public:

        /**
        Returns the tab background color.

        [MSDN documentation for Tab.Color](https://learn.microsoft.com/en-us/office/vba/api/excel.tab.color).
        */
        wxColour GetColor();

        /**
        Sets the tab background color.

        [MSDN documentation for Tab.Color](https://learn.microsoft.com/en-us/office/vba/api/excel.tab.color).
        */
        void SetColor(const wxColour& color);

        /**
        Returns palette color index of the tab.

        [MSDN documentation for Tab.ColorIndex](https://learn.microsoft.com/en-us/office/vba/api/excel.tab.colorindex).
        */
        long GetColorIndex();

        /**
        Sets palette color index of the tab.

        [MSDN documentation for Tab.ColorIndex](https://learn.microsoft.com/en-us/office/vba/api/excel.tab.colorindex).
        */
        void SetColorIndex(long colorIndex);


        /**
        Returns the theme color in the applied color scheme.  Since Excel 2007.

        [MSDN documentation for Tab.ThemeColor](https://learn.microsoft.com/en-us/office/vba/api/excel.tab.themecolor).
        */
        XlThemeFont GetThemeColor();

        /**
        Sets the theme color in the applied color scheme.

        [MSDN documentation for Tab.ThemeColor](https://learn.microsoft.com/en-us/office/vba/api/excel.tab.themecolor).
        */
        void SetThemeColor(XlThemeFont themeColor);

        /**
        A number from -1 (darkest) to 1 (lightest) or 0 for neutral. Since Excel 2007.

        [MSDN documentation for Tab.TintAndShade](https://learn.microsoft.com/en-us/office/vba/api/excel.tab.tintandshade).
        */
        double GetTintAndShade();

        /**
        A number from -1 (darkest) to 1 (lightest) or 0 for neutral.

        [MSDN documentation for Tab.TintAndShade](https://learn.microsoft.com/en-us/office/vba/api/excel.tab.tintandshade).
        */
        void SetTintAndShade(double tintAndShade);

        /**
        Returns "Tab".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("Tab"); }
    };


} // namespace wxAutoExcel

#endif //_WXAUTOEXCEL_TAB_H
