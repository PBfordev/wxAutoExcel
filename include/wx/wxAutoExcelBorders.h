/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// License:     MIT license
/////////////////////////////////////////////////////////////////////////////


#ifndef _WXAUTOEXCEL_BORDERS_H
#define _WXAUTOEXCEL_BORDERS_H

#include "wx/wxAutoExcel_defs.h"
#include "wx/wxAutoExcel_object.h"
#include "wx/wxAutoExcel_enums.h"

class WXDLLIMPEXP_FWD_CORE wxColour;

namespace wxAutoExcel {

    /**
    @brief Represents Microsoft Excel Border object.
    */

   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelBorder : public wxExcelObject
    {
    public:
        // ***** PROPERTIES *****


        /**
        Returns the primary color.

        [MSDN documentation for Border.Color](https://learn.microsoft.com/en-us/office/vba/api/excel.border.color).
        */
        wxColour GetColor();

        /**
        Sets the primary color of the object.

        [MSDN documentation for Border.Color](https://learn.microsoft.com/en-us/office/vba/api/excel.border.color).
        */
        void SetColor(const wxColour& color);

        /**
        Returns the primary color index.

        [MSDN documentation for Border.ColorIndex](https://learn.microsoft.com/en-us/office/vba/api/excel.border.colorindex).
        */
        long GetColorIndex();

        /**
        Sets the primary color index.

        [MSDN documentation for Border.ColorIndex](https://learn.microsoft.com/en-us/office/vba/api/excel.border.colorindex).
        */
        void SetColorIndex(long colorIndex);

        /**
        Returns the line style for the border. Read/write XlLineStyle, xlGray25, xlGray50, xlGray75, or xlAutomatic.

        [MSDN documentation for Border.LineStyle](https://learn.microsoft.com/en-us/office/vba/api/excel.border.linestyle).
        */
        long GetLineStyle();

        /**
        Sets the line style for the border. Read/write XlLineStyle, xlGray25, xlGray50, xlGray75, or xlAutomatic.

        [MSDN documentation for Border.LineStyle](https://learn.microsoft.com/en-us/office/vba/api/excel.border.linestyle).
        */
        void SetLineStyle(long lineStyle);

        /**
        Returns the theme color in the applied color scheme.  Since Excel 2007.

        [MSDN documentation for Border.ThemeColor](https://learn.microsoft.com/en-us/office/vba/api/excel.border.themecolor).
        */
        XlThemeColor GetThemeColor();

        /**
        Sets the theme color in the applied color scheme.

        [MSDN documentation for Border.ThemeColor](https://learn.microsoft.com/en-us/office/vba/api/excel.border.themecolor).
        */
        void SetThemeColor(XlThemeColor themeColor);

        /**
        A number from -1 (darkest) to 1 (lightest) or 0 for neutral. Since Excel 2007.

        [MSDN documentation for Border.TintAndShade](https://learn.microsoft.com/en-us/office/vba/api/excel.border.tintandshade).
        */
        double GetTintAndShade();

        /**
        A number from -1 (darkest) to 1 (lightest) or 0 for neutral.

        [MSDN documentation for Border.TintAndShade](https://learn.microsoft.com/en-us/office/vba/api/excel.border.tintandshade).
        */
        void SetTintAndShade(double tintAndShade);

        /**
        Returns a XlBorderWeight value that represents the weight of the border.

        [MSDN documentation for Border.Weight](https://learn.microsoft.com/en-us/office/vba/api/excel.border.weight).
        */
        XlBorderWeight GetWeight();

        /**
        Sets a XlBorderWeight value that represents the weight of the border.

        [MSDN documentation for Border.Weight](https://learn.microsoft.com/en-us/office/vba/api/excel.border.weight).
        */
        void SetWeight(XlBorderWeight weight);

        /**
        Returns "Border".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("Border"); }
    };


    /**
    @brief Represents Microsoft Excel Borders collection.
    */

   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelBorders : public wxExcelBorder
    {
    public:
        // ***** PROPERTIES *****

        /**
        Returns a Long value that represents the number of objects in the collection.

        [MSDN documentation for Borders.Count](https://learn.microsoft.com/en-us/office/vba/api/excel.borders.count).
        */
        long GetCount();

        //@{
        /**
        Returns a Border Represents one of the borders of either a range of cells or a style.

        [MSDN documentation for Borders.Item](https://learn.microsoft.com/en-us/office/vba/api/excel.borders.item).
        */
        wxExcelBorder GetItem(XlBordersIndex index);
        wxExcelBorder operator[](XlBordersIndex index);
        //@}

        /**
        Returns "Borders".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("Borders"); }
    };


} // namespace wxAutoExcel

#endif //_WXAUTOEXCEL_BORDER_H
