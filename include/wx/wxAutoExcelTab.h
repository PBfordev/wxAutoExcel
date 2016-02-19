/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Modified by:
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
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

        [MSDN documentation for Tab.Color](http://msdn.microsoft.com/en-us/library/bb238208).
        */
        wxColour GetColor();

        /**
        Sets the tab background color.

        [MSDN documentation for Tab.Color](http://msdn.microsoft.com/en-us/library/bb238208).
        */
        void SetColor(const wxColour& color);

        /**
        Returns palette color index of the tab.

        [MSDN documentation for Tab.ColorIndex](http://msdn.microsoft.com/en-us/library/bb238216).
        */
        long GetColorIndex();

        /**
        Sets palette color index of the tab.

        [MSDN documentation for Tab.ColorIndex](http://msdn.microsoft.com/en-us/library/bb238216).
        */
        void SetColorIndex(long colorIndex);
        

        /**
        Returns the theme color in the applied color scheme.  Since Excel 2007.

        [MSDN documentation for Tab.ThemeColor](http://msdn.microsoft.com/en-us/library/bb215144).
        */
        XlThemeFont GetThemeColor();

        /**
        Sets the theme color in the applied color scheme.

        [MSDN documentation for Tab.ThemeColor](http://msdn.microsoft.com/en-us/library/bb215144).
        */
        void SetThemeColor(XlThemeFont themeColor);

        /**
        A number from -1 (darkest) to 1 (lightest) or 0 for neutral. Since Excel 2007.

        [MSDN documentation for Tab.TintAndShade](http://msdn.microsoft.com/en-us/library/bb215148).
        */
        double GetTintAndShade();

        /**
        A number from -1 (darkest) to 1 (lightest) or 0 for neutral.

        [MSDN documentation for Tab.TintAndShade](http://msdn.microsoft.com/en-us/library/bb215148).
        */
        void SetTintAndShade(double tintAndShade);

        /**
        Returns "Tab".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("Tab"); }
    };


} // namespace wxAutoExcel

#endif //_WXAUTOEXCEL_TAB_H
