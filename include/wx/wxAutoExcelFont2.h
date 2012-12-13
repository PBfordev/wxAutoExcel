/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Modified by:
// RCS-ID:      $Id: $
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////


#ifndef _WXAUTOEXCEL_FONT2_H
#define _WXAUTOEXCEL_FONT2_H

#include "wx/wxAutoExcel_defs.h"

#if WXAUTOEXCEL_USE_SHAPES || WXAUTOEXCEL_USE_CHARTS

#include "wx/wxAutoExcelObject.h"
#include "wx/wxAutoExcel_enums.h"

namespace wxAutoExcel {

    /**
    @brief Represents Microsoft Excel Font2 object.

    [MSDN documentation for Font](http://msdn.microsoft.com/en-us/library/bb223824%28v=office.12%29.aspx).
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelFont2 : public wxExcelObject
    {
    public:
        /**
        True if the font is formatted as all capital letters.

        [MSDN documentation for Font2.Allcaps](http://msdn.microsoft.com/en-us/library/aa434534.aspx).
        */
        MsoTriState GetAllcaps();

        /**
        True if the font is formatted as all capital letters.

        [MSDN documentation for Font2.Allcaps](http://msdn.microsoft.com/en-us/library/aa434534.aspx).
        */
        void SetAllcaps(MsoTriState allcaps);

        /**
        Gets or sets a value that specifies whether the numbers in a numbered list should be rotated when the text is rotated.

        [MSDN documentation for Font2.AutorotateNumbers](http://msdn.microsoft.com/en-us/library/aa434536.aspx).
        */
        MsoTriState GetAutorotateNumbers();

        /**
        Gets or sets a value that specifies whether the numbers in a numbered list should be rotated when the text is rotated.

        [MSDN documentation for Font2.AutorotateNumbers](http://msdn.microsoft.com/en-us/library/aa434536.aspx).
        */
        void SetAutorotateNumbers(MsoTriState autorotateNumbers);

        /**
        Gets or sets a value specifying the horizontaol offset of the selected font.

        [MSDN documentation for Font2.BaselineOffset](http://msdn.microsoft.com/en-us/library/aa434537.aspx).
        */
        double GetBaselineOffset();

        /**
        Gets or sets a value specifying the horizontaol offset of the selected font.

        [MSDN documentation for Font2.BaselineOffset](http://msdn.microsoft.com/en-us/library/aa434537.aspx).
        */
        void SetBaselineOffset(double baselineOffset);

        /**
        Gets or sets a value specifying whether the font should be bold.

        [MSDN documentation for Font2.Bold](http://msdn.microsoft.com/en-us/library/aa434538.aspx).
        */
        MsoTriState GetBold();

        /**
        Gets or sets a value specifying whether the font should be bold.

        [MSDN documentation for Font2.Bold](http://msdn.microsoft.com/en-us/library/aa434538.aspx).
        */
        void SetBold(MsoTriState bold);

        /**
        Gets or sets a value specifying that the text should be capitalized.

        [MSDN documentation for Font2.Caps](http://msdn.microsoft.com/en-us/library/aa434539.aspx).
        */
        MsoTextCaps GetCaps();

        /**
        Gets or sets a value specifying that the text should be capitalized.

        [MSDN documentation for Font2.Caps](http://msdn.microsoft.com/en-us/library/aa434539.aspx).
        */
        void SetCaps(MsoTextCaps caps);

        /**
        True if the specified font is formatted as double strikethrough text.

        [MSDN documentation for Font2.DoubleStrikeThrough](http://msdn.microsoft.com/en-us/library/aa434541.aspx).
        */
        MsoTriState GetDoubleStrikeThrough();

        /**
        True if the specified font is formatted as double strikethrough text.

        [MSDN documentation for Font2.DoubleStrikeThrough](http://msdn.microsoft.com/en-us/library/aa434541.aspx).
        */
        void SetDoubleStrikeThrough(MsoTriState doubleStrikeThrough);

        /**
        Gets a value indicating whether the font can be embedded in a page.

        [MSDN documentation for Font2.Embeddable](http://msdn.microsoft.com/en-us/library/aa434542.aspx).
        */
        MsoTriState GetEmbeddable();

        /**
        Gets a value specifying whether the font is embedded in a page.

        [MSDN documentation for Font2.Embedded](http://msdn.microsoft.com/en-us/library/aa434543.aspx).
        */
        MsoTriState GetEmbedded();

        /**
        Gets or sets a value specifying whether the text for a selection should be spaced equal distances apart.

        [MSDN documentation for Font2.Equalize](http://msdn.microsoft.com/en-us/library/aa434544.aspx).
        */
        MsoTriState GetEqualize();

        /**
        Gets or sets a value specifying whether the text for a selection should be spaced equal distances apart.

        [MSDN documentation for Font2.Equalize](http://msdn.microsoft.com/en-us/library/aa434544.aspx).
        */
        void SetEqualize(MsoTriState equalize);

        /**
        Gets a value specifying the fill format for a font.

        [MSDN documentation for Font2.Fill](http://msdn.microsoft.com/en-us/library/aa434545.aspx).
        */
        wxExcelFillFormat GetFill();

        /**
        Gets a value indicating whether the font is displayed as a glow effect.

        [MSDN documentation for Font2.Glow](http://msdn.microsoft.com/en-us/library/aa434546.aspx).
        */
        wxExcelGlowFormat GetGlow();

        /**
        Gets a value indicating whether the font is displayed as highlighted.

        [MSDN documentation for Font2.Highlight](http://msdn.microsoft.com/en-us/library/aa434547.aspx).
        */
        wxExcelColorFormat GetHighlight();

        /**
        Gets or sets a value specifying whether the text for a selection is italic.

        [MSDN documentation for Font2.Italic](http://msdn.microsoft.com/en-us/library/aa434548.aspx).
        */
        MsoTriState GetItalic();

        /**
        Gets or sets a value specifying whether the text for a selection is italic.

        [MSDN documentation for Font2.Italic](http://msdn.microsoft.com/en-us/library/aa434548.aspx).
        */
        void SetItalic(MsoTriState italic);

        /**
        Gets or sets a value specifying the amount of spacing between text characters.

        [MSDN documentation for Font2.Kerning](http://msdn.microsoft.com/en-us/library/aa434549.aspx).
        */
        double GetKerning();

        /**
        Gets or sets a value specifying the amount of spacing between text characters.

        [MSDN documentation for Font2.Kerning](http://msdn.microsoft.com/en-us/library/aa434549.aspx).
        */
        void SetKerning(double kerning);

        /**
        Gets a value specifiying the format of a line.

        [MSDN documentation for Font2.Line](http://msdn.microsoft.com/en-us/library/aa434550.aspx).
        */
        wxExcelLineFormat GetLine();

        /**
        Gets or sets a value specifying the font to use for a selection.

        [MSDN documentation for Font2.Name](http://msdn.microsoft.com/en-us/library/aa434551.aspx).
        */
        wxString GetName();

        /**
        Gets or sets a value specifying the font to use for a selection.

        [MSDN documentation for Font2.Name](http://msdn.microsoft.com/en-us/library/aa434551.aspx).
        */
        void SetName(const wxString& name);

        /**
        Gets or sets the font used for Latin text (characters with character codes from 0 (zero) through 127).

        [MSDN documentation for Font2.NameAscii](http://msdn.microsoft.com/en-us/library/aa434552.aspx).
        */
        wxString GetNameAscii();

        /**
        Gets or sets the font used for Latin text (characters with character codes from 0 (zero) through 127).

        [MSDN documentation for Font2.NameAscii](http://msdn.microsoft.com/en-us/library/aa434552.aspx).
        */
        void SetNameAscii(const wxString& nameAscii);

        /**
        Gets or sets the complex script font name. Used for mixed language text..

        [MSDN documentation for Font2.NameComplexScript](http://msdn.microsoft.com/en-us/library/aa434553.aspx).
        */
        wxString GetNameComplexScript();

        /**
        Gets or sets the complex script font name. Used for mixed language text..

        [MSDN documentation for Font2.NameComplexScript](http://msdn.microsoft.com/en-us/library/aa434553.aspx).
        */
        void SetNameComplexScript(const wxString& nameComplexScript);

        /**
        Gets or sets an East Asian font name.

        [MSDN documentation for Font2.NameFarEast](http://msdn.microsoft.com/en-us/library/aa434554.aspx).
        */
        wxString GetNameFarEast();

        /**
        Gets or sets an East Asian font name.

        [MSDN documentation for Font2.NameFarEast](http://msdn.microsoft.com/en-us/library/aa434554.aspx).
        */
        void SetNameFarEast(const wxString& nameFarEast);

        /**
        Gets or sets the font used for characters whose character set numbers are greater than 127.

        [MSDN documentation for Font2.NameOther](http://msdn.microsoft.com/en-us/library/aa434556.aspx).
        */
        wxString GetNameOther();

        /**
        Gets or sets the font used for characters whose character set numbers are greater than 127.

        [MSDN documentation for Font2.NameOther](http://msdn.microsoft.com/en-us/library/aa434556.aspx).
        */
        void SetNameOther(const wxString& nameOther);

        /**
        Gets a value specifying the type of reflection format for the selection of text.

        [MSDN documentation for Font2.Reflection](http://msdn.microsoft.com/en-us/library/aa434558.aspx).
        */
        wxExcelReflectionFormat GetReflection();

        /**
        Gets the value specifying the type of shadow effect for the selection of text.

        [MSDN documentation for Font2.Shadow](http://msdn.microsoft.com/en-us/library/aa434559.aspx).
        */
        wxExcelShadowFormat GetShadow();

        /**
        Gets or sets a value specifying the size of the font.

        [MSDN documentation for Font2.Size](http://msdn.microsoft.com/en-us/library/aa434560.aspx).
        */
        double GetSize();

        /**
        Gets or sets a value specifying the size of the font.

        [MSDN documentation for Font2.Size](http://msdn.microsoft.com/en-us/library/aa434560.aspx).
        */
        void SetSize(double size);

        /**
        Gets or sets a value specifying whether small caps should be used with the slection of text. Small caps are the same height as the lowercase letters in a slection of text.

        [MSDN documentation for Font2.Smallcaps](http://msdn.microsoft.com/en-us/library/aa434561.aspx).
        */
        MsoTriState GetSmallcaps();

        /**
        Gets or sets a value specifying whether small caps should be used with the slection of text. Small caps are the same height as the lowercase letters in a slection of text.

        [MSDN documentation for Font2.Smallcaps](http://msdn.microsoft.com/en-us/library/aa434561.aspx).
        */
        void SetSmallcaps(MsoTriState smallcaps);

        /**
        Gets or sets a value specifying the type of soft edge effect used in a selection of text.

        [MSDN documentation for Font2.SoftEdgeFormat](http://msdn.microsoft.com/en-us/library/aa434562.aspx).
        */
        MsoSoftEdgeType GetSoftEdgeFormat();

        /**
        Gets or sets a value specifying the type of soft edge effect used in a selection of text.

        [MSDN documentation for Font2.SoftEdgeFormat](http://msdn.microsoft.com/en-us/library/aa434562.aspx).
        */
        void SetSoftEdgeFormat(MsoSoftEdgeType softEdgeFormat);

        /**
        Gets or sets a value specifying the spacing between characters in a selection of text.

        [MSDN documentation for Font2.Spacing](http://msdn.microsoft.com/en-us/library/aa434563.aspx).
        */
        double GetSpacing();

        /**
        Gets or sets a value specifying the spacing between characters in a selection of text.

        [MSDN documentation for Font2.Spacing](http://msdn.microsoft.com/en-us/library/aa434563.aspx).
        */
        void SetSpacing(double spacing);

        /**
        Gets or sets a value specifying the strike format used for a selection of text.

        [MSDN documentation for Font2.Strike](http://msdn.microsoft.com/en-us/library/aa434564.aspx).
        */
        MsoTextStrike GetStrike();

        /**
        Gets or sets a value specifying the strike format used for a selection of text.

        [MSDN documentation for Font2.Strike](http://msdn.microsoft.com/en-us/library/aa434564.aspx).
        */
        void SetStrike(MsoTextStrike strike);

        /**
        Gets or sets a value specifying the text should be rendered in a strikethrough appearance.

        [MSDN documentation for Font2.StrikeThrough](http://msdn.microsoft.com/en-us/library/aa434565.aspx).
        */
        MsoTriState GetStrikeThrough();

        /**
        Gets or sets a value specifying the text should be rendered in a strikethrough appearance.

        [MSDN documentation for Font2.StrikeThrough](http://msdn.microsoft.com/en-us/library/aa434565.aspx).
        */
        void SetStrikeThrough(MsoTriState strikeThrough);

        /**
        Gets or sets a value specifying that the selected text should be displayed a subscript.

        [MSDN documentation for Font2.Subscript](http://msdn.microsoft.com/en-us/library/aa434566.aspx).
        */
        MsoTriState GetSubscript();

        /**
        Gets or sets a value specifying that the selected text should be displayed a subscript.

        [MSDN documentation for Font2.Subscript](http://msdn.microsoft.com/en-us/library/aa434566.aspx).
        */
        void SetSubscript(MsoTriState subscript);

        /**
        Gets or sets a value specifying that the selected text should be displayed a subscript.

        [MSDN documentation for Font2.Superscript](http://msdn.microsoft.com/en-us/library/aa434567.aspx).
        */
        MsoTriState GetSuperscript();

        /**
        Gets or sets a value specifying that the selected text should be displayed a subscript.

        [MSDN documentation for Font2.Superscript](http://msdn.microsoft.com/en-us/library/aa434567.aspx).
        */
        void SetSuperscript(MsoTriState superscript);

        /**
        Gets a value specifying the color of the underline for the selected text.

        [MSDN documentation for Font2.UnderlineColor](http://msdn.microsoft.com/en-us/library/aa434568.aspx).
        */
        wxExcelColorFormat GetUnderlineColor();

        /**
        Gets or sets a value specifying the underline style for the selected text.

        [MSDN documentation for Font2.UnderlineStyle](http://msdn.microsoft.com/en-us/library/aa434569.aspx).
        */
        MsoTextUnderlineType GetUnderlineStyle();

        /**
        Gets or sets a value specifying the underline style for the selected text.

        [MSDN documentation for Font2.UnderlineStyle](http://msdn.microsoft.com/en-us/library/aa434569.aspx).
        */
        void SetUnderlineStyle(MsoTextUnderlineType underlineStyle);

        /**
        Gets or sets a value specifying the text effect for the selected text.

        [MSDN documentation for Font2.WordArtformat](http://msdn.microsoft.com/en-us/library/aa434570.aspx).
        */
        MsoPresetTextEffect GetWordArtformat();

        /**
        Gets or sets a value specifying the text effect for the selected text.

        [MSDN documentation for Font2.WordArtformat](http://msdn.microsoft.com/en-us/library/aa434570.aspx).
        */
        void SetWordArtformat(MsoPresetTextEffect wordArtformat);

        /**
        Returns "Font2".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("Font2"); }
    };


} // namespace wxAutoExcel

#endif // #if WXAUTOEXCEL_USE_SHAPES || WXAUTOEXCEL_USE_CHARTS

#endif //_WXAUTOEXCEL_FONT2_H
