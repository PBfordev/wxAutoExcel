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

    [MSDN documentation for Font](http://msdn.microsoft.com/en-us/library/bb223824%28v=office.12%29.aspx).
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelFont : public wxExcelObject
    {
    public:
        // ***** PROPERTIES *****


        /**
        Returns the type of background.

        [MSDN documentation for Font.Background](http://msdn.microsoft.com/en-us/library/bb220875).
        */
        XlBackground GetBackground();

        /**
        Sets the type of background.

        [MSDN documentation for Font.Background](http://msdn.microsoft.com/en-us/library/bb220875).
        */
        void SetBackground(XlBackground background);

        /**
        True if the font is bold.

        [MSDN documentation for Font.Bold](http://msdn.microsoft.com/en-us/library/bb220891).
        */
        bool GetBold();

        /**
        True if the font is bold.

        [MSDN documentation for Font.Bold](http://msdn.microsoft.com/en-us/library/bb220891).
        */
        void SetBold(bool bold);

        /**
        Returns the primary color.

        [MSDN documentation for Font.Color](http://msdn.microsoft.com/en-us/library/bb213182).
        */
        wxColour GetColor();

        /**
        Sets the primary color.

        [MSDN documentation for Font.Color](http://msdn.microsoft.com/en-us/library/bb213182).
        */
        void SetColor(const wxColour& color);

        /**
        Returns the color of the font.

        [MSDN documentation for Font.ColorIndex](http://msdn.microsoft.com/en-us/library/bb213184).
        */
        long GetColorIndex();

        /**
        Sets the color of the font.

        [MSDN documentation for Font.ColorIndex](http://msdn.microsoft.com/en-us/library/bb213184).
        */
        void SetColorIndex(long colorIndex);

        /**
        Returns the font style.

        [MSDN documentation for Font.FontStyle](http://msdn.microsoft.com/en-us/library/bb208525).
        */
        wxString GetFontStyle();

        /**
        Sets the font style.

        [MSDN documentation for Font.FontStyle](http://msdn.microsoft.com/en-us/library/bb208525).
        */
        void SetFontStyle(const wxString& fontStyle);

        /**
        True if the font style is italic.

        [MSDN documentation for Font.Italic](http://msdn.microsoft.com/en-us/library/bb177802).
        */
        bool GetItalic();

        /**
        True if the font style is italic.

        [MSDN documentation for Font.Italic](http://msdn.microsoft.com/en-us/library/bb177802).
        */
        void SetItalic(bool italic);

        /**
        Returns font name.

        [MSDN documentation for Font.Name](http://msdn.microsoft.com/en-us/library/bb213187).
        */
        wxString GetName();

        /**
        Sets font name.

        [MSDN documentation for Font.Name](http://msdn.microsoft.com/en-us/library/bb213187).
        */
        void SetName(const wxString& name);

        /**
        Returns font size in points

        [MSDN documentation for Font.Size](http://msdn.microsoft.com/en-us/library/bb214617).
        */
        double GetSize();

        /**
        Sets font size in points

        [MSDN documentation for Font.Size](http://msdn.microsoft.com/en-us/library/bb214617).
        */
        void SetSize(double size);

        /**
        True if the font is struck through with a horizontal line.

        [MSDN documentation for Font.Strikethrough](http://msdn.microsoft.com/en-us/library/bb209303).
        */
        bool GetStrikethrough();

        /**
        True if the font is struck through with a horizontal line.

        [MSDN documentation for Font.Strikethrough](http://msdn.microsoft.com/en-us/library/bb209303).
        */
        void SetStrikethrough(bool strikethrough);

        /**
        True if the font is formatted as subscript. False by default.

        [MSDN documentation for Font.Subscript](http://msdn.microsoft.com/en-us/library/bb209310).
        */
        bool GetSubscript();

        /**
        True if the font is formatted as subscript. False by default.

        [MSDN documentation for Font.Subscript](http://msdn.microsoft.com/en-us/library/bb209310).
        */
        void SetSubscript(bool subscript);

        /**
        True if the font is formatted as superscript; False by default.

        [MSDN documentation for Font.Superscript](http://msdn.microsoft.com/en-us/library/bb209324).
        */
        bool GetSuperscript();

        /**
        True if the font is formatted as superscript; False by default.

        [MSDN documentation for Font.Superscript](http://msdn.microsoft.com/en-us/library/bb209324).
        */
        void SetSuperscript(bool superscript);

        /**
        Returns the theme color in the applied color scheme.  Since Excel 2007.

        [MSDN documentation for Font.ThemeColor](http://msdn.microsoft.com/en-us/library/bb214619).
        */
        XlThemeColor GetThemeColor();

        /**
        Sets the theme color in the applied color scheme.

        [MSDN documentation for Font.ThemeColor](http://msdn.microsoft.com/en-us/library/bb214619).
        */
        void SetThemeColor(XlThemeColor themeColor);

        /**
        Returns the theme font in the applied font scheme.

        [MSDN documentation for Font.ThemeFont](http://msdn.microsoft.com/en-us/library/bb215738).
        */
        XlThemeFont GetThemeFont();

        /**
        Sets the theme font in the applied font scheme.

        [MSDN documentation for Font.ThemeFont](http://msdn.microsoft.com/en-us/library/bb215738).
        */
        void SetThemeFont(XlThemeFont themeFont);

        /**
        A number from -1 (darkest) to 1 (lightest) or 0 for neutral. Since Excel 2007.

        [MSDN documentation for Font.TintAndShade](http://msdn.microsoft.com/en-us/library/bb214620).
        */
        double GetTintAndShade();

        /**
        A number from -1 (darkest) to 1 (lightest) or 0 for neutral.

        [MSDN documentation for Font.TintAndShade](http://msdn.microsoft.com/en-us/library/bb214620).
        */
        void SetTintAndShade(double tintAndShade);

        /**
        Returns "Font".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("Font"); }
    };


} // namespace wxAutoExcel

#endif //_WXAUTOEXCEL_FONT_H
