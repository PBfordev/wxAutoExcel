/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// License:     MIT license
/////////////////////////////////////////////////////////////////////////////


#ifndef _WXAUTOEXCEL_FONT_H
#define _WXAUTOEXCEL_FONT_H

#include "wx/wxAutoExcel_defs.h"

#include "wx/wxAutoExcel_object.h"
#include "wx/wxAutoExcel_enums.h"

class WXDLLIMPEXP_FWD_CORE wxColour;

namespace wxAutoExcel {


    /**
    @brief Represents Microsoft Excel Font.

    [MSDN documentation for Font](https://learn.microsoft.com/en-us/office/vba/api/excel.font%28object%29)).
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelFont : public wxExcelObject
    {
    public:
        // ***** PROPERTIES *****


        /**
        Returns the type of background.

        [MSDN documentation for Font.Background](https://learn.microsoft.com/en-us/office/vba/api/excel.font.background).
        */
        XlBackground GetBackground();

        /**
        Sets the type of background.

        [MSDN documentation for Font.Background](https://learn.microsoft.com/en-us/office/vba/api/excel.font.background).
        */
        void SetBackground(XlBackground background);

        /**
        True if the font is bold.

        [MSDN documentation for Font.Bold](https://learn.microsoft.com/en-us/office/vba/api/excel.font.bold).
        */
        bool GetBold();

        /**
        True if the font is bold.

        [MSDN documentation for Font.Bold](https://learn.microsoft.com/en-us/office/vba/api/excel.font.bold).
        */
        void SetBold(bool bold);

        /**
        Returns the primary color.

        [MSDN documentation for Font.Color](https://learn.microsoft.com/en-us/office/vba/api/excel.font.color).
        */
        wxColour GetColor();

        /**
        Sets the primary color.

        [MSDN documentation for Font.Color](https://learn.microsoft.com/en-us/office/vba/api/excel.font.color).
        */
        void SetColor(const wxColour& color);

        /**
        Returns the color of the font.

        [MSDN documentation for Font.ColorIndex](https://learn.microsoft.com/en-us/office/vba/api/excel.font.colorindex).
        */
        long GetColorIndex();

        /**
        Sets the color of the font.

        [MSDN documentation for Font.ColorIndex](https://learn.microsoft.com/en-us/office/vba/api/excel.font.colorindex).
        */
        void SetColorIndex(long colorIndex);

        /**
        Returns the font style.

        [MSDN documentation for Font.FontStyle](https://learn.microsoft.com/en-us/office/vba/api/excel.font.fontstyle).
        */
        wxString GetFontStyle();

        /**
        Sets the font style.

        [MSDN documentation for Font.FontStyle](https://learn.microsoft.com/en-us/office/vba/api/excel.font.fontstyle).
        */
        void SetFontStyle(const wxString& fontStyle);

        /**
        True if the font style is italic.

        [MSDN documentation for Font.Italic](https://learn.microsoft.com/en-us/office/vba/api/excel.font.italic).
        */
        bool GetItalic();

        /**
        True if the font style is italic.

        [MSDN documentation for Font.Italic](https://learn.microsoft.com/en-us/office/vba/api/excel.font.italic).
        */
        void SetItalic(bool italic);

        /**
        Returns font name.

        [MSDN documentation for Font.Name](https://learn.microsoft.com/en-us/office/vba/api/excel.font.name).
        */
        wxString GetName();

        /**
        Sets font name.

        [MSDN documentation for Font.Name](https://learn.microsoft.com/en-us/office/vba/api/excel.font.name).
        */
        void SetName(const wxString& name);

        /**
        Returns font size in points

        [MSDN documentation for Font.Size](https://learn.microsoft.com/en-us/office/vba/api/excel.font.size).
        */
        double GetSize();

        /**
        Sets font size in points

        [MSDN documentation for Font.Size](https://learn.microsoft.com/en-us/office/vba/api/excel.font.size).
        */
        void SetSize(double size);

        /**
        True if the font is struck through with a horizontal line.

        [MSDN documentation for Font.Strikethrough](https://learn.microsoft.com/en-us/office/vba/api/excel.font.strikethrough).
        */
        bool GetStrikethrough();

        /**
        True if the font is struck through with a horizontal line.

        [MSDN documentation for Font.Strikethrough](https://learn.microsoft.com/en-us/office/vba/api/excel.font.strikethrough).
        */
        void SetStrikethrough(bool strikethrough);

        /**
        True if the font is formatted as subscript. False by default.

        [MSDN documentation for Font.Subscript](https://learn.microsoft.com/en-us/office/vba/api/excel.font.subscript).
        */
        bool GetSubscript();

        /**
        True if the font is formatted as subscript. False by default.

        [MSDN documentation for Font.Subscript](https://learn.microsoft.com/en-us/office/vba/api/excel.font.subscript).
        */
        void SetSubscript(bool subscript);

        /**
        True if the font is formatted as superscript; False by default.

        [MSDN documentation for Font.Superscript](https://learn.microsoft.com/en-us/office/vba/api/excel.font.superscript).
        */
        bool GetSuperscript();

        /**
        True if the font is formatted as superscript; False by default.

        [MSDN documentation for Font.Superscript](https://learn.microsoft.com/en-us/office/vba/api/excel.font.superscript).
        */
        void SetSuperscript(bool superscript);

        /**
        Returns the theme color in the applied color scheme.  Since Excel 2007.

        [MSDN documentation for Font.ThemeColor](https://learn.microsoft.com/en-us/office/vba/api/excel.font.themecolor).
        */
        XlThemeColor GetThemeColor();

        /**
        Sets the theme color in the applied color scheme.

        [MSDN documentation for Font.ThemeColor](https://learn.microsoft.com/en-us/office/vba/api/excel.font.themecolor).
        */
        void SetThemeColor(XlThemeColor themeColor);

        /**
        Returns the theme font in the applied font scheme.

        [MSDN documentation for Font.ThemeFont](https://learn.microsoft.com/en-us/office/vba/api/excel.font.themefont).
        */
        XlThemeFont GetThemeFont();

        /**
        Sets the theme font in the applied font scheme.

        [MSDN documentation for Font.ThemeFont](https://learn.microsoft.com/en-us/office/vba/api/excel.font.themefont).
        */
        void SetThemeFont(XlThemeFont themeFont);

        /**
        A number from -1 (darkest) to 1 (lightest) or 0 for neutral. Since Excel 2007.

        [MSDN documentation for Font.TintAndShade](https://learn.microsoft.com/en-us/office/vba/api/excel.font.tintandshade).
        */
        double GetTintAndShade();

        /**
        A number from -1 (darkest) to 1 (lightest) or 0 for neutral.

        [MSDN documentation for Font.TintAndShade](https://learn.microsoft.com/en-us/office/vba/api/excel.font.tintandshade).
        */
        void SetTintAndShade(double tintAndShade);

        /**
        Returns "Font".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("Font"); }
    };


} // namespace wxAutoExcel

#endif //_WXAUTOEXCEL_FONT_H
