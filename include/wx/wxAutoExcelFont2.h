/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// License:     MIT license
/////////////////////////////////////////////////////////////////////////////


#ifndef _WXAUTOEXCEL_FONT2_H
#define _WXAUTOEXCEL_FONT2_H

#include "wx/wxAutoExcel_defs.h"

#if WXAUTOEXCEL_USE_SHAPES || WXAUTOEXCEL_USE_CHARTS

#include "wx/wxAutoExcel_object.h"
#include "wx/wxAutoExcel_enums.h"

namespace wxAutoExcel {

    /**
    @brief Represents Microsoft Excel Font2 object.

    [MSDN documentation for Font2](https://learn.microsoft.com/en-us/office/vba/api/office.font2).
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelFont2 : public wxExcelObject
    {
    public:
        /**
        True if the font is formatted as all capital letters.

        [MSDN documentation for Font2.Allcaps](https://learn.microsoft.com/en-us/office/vba/api/office.font2.allcaps).
        */
        MsoTriState GetAllcaps();

        /**
        True if the font is formatted as all capital letters.

        [MSDN documentation for Font2.Allcaps](https://learn.microsoft.com/en-us/office/vba/api/office.font2.allcaps).
        */
        void SetAllcaps(MsoTriState allcaps);

        /**
        Gets or sets a value that specifies whether the numbers in a numbered list should be rotated when the text is rotated.

        [MSDN documentation for Font2.AutorotateNumbers](https://learn.microsoft.com/en-us/office/vba/api/office.font2.autorotatenumbers).
        */
        MsoTriState GetAutorotateNumbers();

        /**
        Gets or sets a value that specifies whether the numbers in a numbered list should be rotated when the text is rotated.

        [MSDN documentation for Font2.AutorotateNumbers](https://learn.microsoft.com/en-us/office/vba/api/office.font2.autorotatenumbers).
        */
        void SetAutorotateNumbers(MsoTriState autorotateNumbers);

        /**
        Gets or sets a value specifying the horizontaol offset of the selected font.

        [MSDN documentation for Font2.BaselineOffset](https://learn.microsoft.com/en-us/office/vba/api/office.font2.baselineoffset).
        */
        double GetBaselineOffset();

        /**
        Gets or sets a value specifying the horizontaol offset of the selected font.

        [MSDN documentation for Font2.BaselineOffset](https://learn.microsoft.com/en-us/office/vba/api/office.font2.baselineoffset).
        */
        void SetBaselineOffset(double baselineOffset);

        /**
        Gets or sets a value specifying whether the font should be bold.

        [MSDN documentation for Font2.Bold](https://learn.microsoft.com/en-us/office/vba/api/office.font2.bold).
        */
        MsoTriState GetBold();

        /**
        Gets or sets a value specifying whether the font should be bold.

        [MSDN documentation for Font2.Bold](https://learn.microsoft.com/en-us/office/vba/api/office.font2.bold).
        */
        void SetBold(MsoTriState bold);

        /**
        Gets or sets a value specifying that the text should be capitalized.

        [MSDN documentation for Font2.Caps](https://learn.microsoft.com/en-us/office/vba/api/office.font2.caps).
        */
        MsoTextCaps GetCaps();

        /**
        Gets or sets a value specifying that the text should be capitalized.

        [MSDN documentation for Font2.Caps](https://learn.microsoft.com/en-us/office/vba/api/office.font2.caps).
        */
        void SetCaps(MsoTextCaps caps);

        /**
        True if the specified font is formatted as double strikethrough text.

        [MSDN documentation for Font2.DoubleStrikeThrough](https://learn.microsoft.com/en-us/office/vba/api/office.font2.doublestrikethrough).
        */
        MsoTriState GetDoubleStrikeThrough();

        /**
        True if the specified font is formatted as double strikethrough text.

        [MSDN documentation for Font2.DoubleStrikeThrough](https://learn.microsoft.com/en-us/office/vba/api/office.font2.doublestrikethrough).
        */
        void SetDoubleStrikeThrough(MsoTriState doubleStrikeThrough);

        /**
        Gets a value indicating whether the font can be embedded in a page.

        [MSDN documentation for Font2.Embeddable](https://learn.microsoft.com/en-us/office/vba/api/office.font2.embeddable).
        */
        MsoTriState GetEmbeddable();

        /**
        Gets a value specifying whether the font is embedded in a page.

        [MSDN documentation for Font2.Embedded](https://learn.microsoft.com/en-us/office/vba/api/office.font2.embedded).
        */
        MsoTriState GetEmbedded();

        /**
        Gets or sets a value specifying whether the text for a selection should be spaced equal distances apart.

        [MSDN documentation for Font2.Equalize](https://learn.microsoft.com/en-us/office/vba/api/office.font2.equalize).
        */
        MsoTriState GetEqualize();

        /**
        Gets or sets a value specifying whether the text for a selection should be spaced equal distances apart.

        [MSDN documentation for Font2.Equalize](https://learn.microsoft.com/en-us/office/vba/api/office.font2.equalize).
        */
        void SetEqualize(MsoTriState equalize);

        /**
        Gets a value specifying the fill format for a font.

        [MSDN documentation for Font2.Fill](https://learn.microsoft.com/en-us/office/vba/api/office.font2.fill).
        */
        wxExcelFillFormat GetFill();

        /**
        Gets a value indicating whether the font is displayed as a glow effect.

        [MSDN documentation for Font2.Glow](https://learn.microsoft.com/en-us/office/vba/api/office.font2.glow).
        */
        wxExcelGlowFormat GetGlow();

        /**
        Gets a value indicating whether the font is displayed as highlighted.

        [MSDN documentation for Font2.Highlight](https://learn.microsoft.com/en-us/office/vba/api/office.font2.highlight).
        */
        wxExcelColorFormat GetHighlight();

        /**
        Gets or sets a value specifying whether the text for a selection is italic.

        [MSDN documentation for Font2.Italic](https://learn.microsoft.com/en-us/office/vba/api/office.font2.italic).
        */
        MsoTriState GetItalic();

        /**
        Gets or sets a value specifying whether the text for a selection is italic.

        [MSDN documentation for Font2.Italic](https://learn.microsoft.com/en-us/office/vba/api/office.font2.italic).
        */
        void SetItalic(MsoTriState italic);

        /**
        Gets or sets a value specifying the amount of spacing between text characters.

        [MSDN documentation for Font2.Kerning](https://learn.microsoft.com/en-us/office/vba/api/office.font2.kerning).
        */
        double GetKerning();

        /**
        Gets or sets a value specifying the amount of spacing between text characters.

        [MSDN documentation for Font2.Kerning](https://learn.microsoft.com/en-us/office/vba/api/office.font2.kerning).
        */
        void SetKerning(double kerning);

        /**
        Gets a value specifiying the format of a line.

        [MSDN documentation for Font2.Line](https://learn.microsoft.com/en-us/office/vba/api/office.font2.line).
        */
        wxExcelLineFormat GetLine();

        /**
        Gets or sets a value specifying the font to use for a selection.

        [MSDN documentation for Font2.Name](https://learn.microsoft.com/en-us/office/vba/api/office.font2.name).
        */
        wxString GetName();

        /**
        Gets or sets a value specifying the font to use for a selection.

        [MSDN documentation for Font2.Name](https://learn.microsoft.com/en-us/office/vba/api/office.font2.name).
        */
        void SetName(const wxString& name);

        /**
        Gets or sets the font used for Latin text (characters with character codes from 0 (zero) through 127).

        [MSDN documentation for Font2.NameAscii](https://learn.microsoft.com/en-us/office/vba/api/office.font2.nameascii).
        */
        wxString GetNameAscii();

        /**
        Gets or sets the font used for Latin text (characters with character codes from 0 (zero) through 127).

        [MSDN documentation for Font2.NameAscii](https://learn.microsoft.com/en-us/office/vba/api/office.font2.nameascii).
        */
        void SetNameAscii(const wxString& nameAscii);

        /**
        Gets or sets the complex script font name. Used for mixed language text..

        [MSDN documentation for Font2.NameComplexScript](https://learn.microsoft.com/en-us/office/vba/api/office.font2.namecomplexscript).
        */
        wxString GetNameComplexScript();

        /**
        Gets or sets the complex script font name. Used for mixed language text..

        [MSDN documentation for Font2.NameComplexScript](https://learn.microsoft.com/en-us/office/vba/api/office.font2.namecomplexscript).
        */
        void SetNameComplexScript(const wxString& nameComplexScript);

        /**
        Gets or sets an East Asian font name.

        [MSDN documentation for Font2.NameFarEast](https://learn.microsoft.com/en-us/office/vba/api/office.font2.namefareast).
        */
        wxString GetNameFarEast();

        /**
        Gets or sets an East Asian font name.

        [MSDN documentation for Font2.NameFarEast](https://learn.microsoft.com/en-us/office/vba/api/office.font2.namefareast).
        */
        void SetNameFarEast(const wxString& nameFarEast);

        /**
        Gets or sets the font used for characters whose character set numbers are greater than 127.

        [MSDN documentation for Font2.NameOther](https://learn.microsoft.com/en-us/office/vba/api/office.font2.nameother).
        */
        wxString GetNameOther();

        /**
        Gets or sets the font used for characters whose character set numbers are greater than 127.

        [MSDN documentation for Font2.NameOther](https://learn.microsoft.com/en-us/office/vba/api/office.font2.nameother).
        */
        void SetNameOther(const wxString& nameOther);

        /**
        Gets a value specifying the type of reflection format for the selection of text.

        [MSDN documentation for Font2.Reflection](https://learn.microsoft.com/en-us/office/vba/api/office.font2.reflection).
        */
        wxExcelReflectionFormat GetReflection();

        /**
        Gets the value specifying the type of shadow effect for the selection of text.

        [MSDN documentation for Font2.Shadow](https://learn.microsoft.com/en-us/office/vba/api/office.font2.shadow).
        */
        wxExcelShadowFormat GetShadow();

        /**
        Gets or sets a value specifying the size of the font.

        [MSDN documentation for Font2.Size](https://learn.microsoft.com/en-us/office/vba/api/office.font2.size).
        */
        double GetSize();

        /**
        Gets or sets a value specifying the size of the font.

        [MSDN documentation for Font2.Size](https://learn.microsoft.com/en-us/office/vba/api/office.font2.size).
        */
        void SetSize(double size);

        /**
        Gets or sets a value specifying whether small caps should be used with the slection of text. Small caps are the same height as the lowercase letters in a slection of text.

        [MSDN documentation for Font2.Smallcaps](https://learn.microsoft.com/en-us/office/vba/api/office.font2.smallcaps).
        */
        MsoTriState GetSmallcaps();

        /**
        Gets or sets a value specifying whether small caps should be used with the slection of text. Small caps are the same height as the lowercase letters in a slection of text.

        [MSDN documentation for Font2.Smallcaps](https://learn.microsoft.com/en-us/office/vba/api/office.font2.smallcaps).
        */
        void SetSmallcaps(MsoTriState smallcaps);

        /**
        Gets or sets a value specifying the type of soft edge effect used in a selection of text.

        [MSDN documentation for Font2.SoftEdgeFormat](https://learn.microsoft.com/en-us/office/vba/api/office.font2.softedgeformat).
        */
        MsoSoftEdgeType GetSoftEdgeFormat();

        /**
        Gets or sets a value specifying the type of soft edge effect used in a selection of text.

        [MSDN documentation for Font2.SoftEdgeFormat](https://learn.microsoft.com/en-us/office/vba/api/office.font2.softedgeformat).
        */
        void SetSoftEdgeFormat(MsoSoftEdgeType softEdgeFormat);

        /**
        Gets or sets a value specifying the spacing between characters in a selection of text.

        [MSDN documentation for Font2.Spacing](https://learn.microsoft.com/en-us/office/vba/api/office.font2.spacing).
        */
        double GetSpacing();

        /**
        Gets or sets a value specifying the spacing between characters in a selection of text.

        [MSDN documentation for Font2.Spacing](https://learn.microsoft.com/en-us/office/vba/api/office.font2.spacing).
        */
        void SetSpacing(double spacing);

        /**
        Gets or sets a value specifying the strike format used for a selection of text.

        [MSDN documentation for Font2.Strike](https://learn.microsoft.com/en-us/office/vba/api/office.font2.strike).
        */
        MsoTextStrike GetStrike();

        /**
        Gets or sets a value specifying the strike format used for a selection of text.

        [MSDN documentation for Font2.Strike](https://learn.microsoft.com/en-us/office/vba/api/office.font2.strike).
        */
        void SetStrike(MsoTextStrike strike);

        /**
        Gets or sets a value specifying the text should be rendered in a strikethrough appearance.

        [MSDN documentation for Font2.StrikeThrough](https://learn.microsoft.com/en-us/office/vba/api/office.font2.strikethrough).
        */
        MsoTriState GetStrikeThrough();

        /**
        Gets or sets a value specifying the text should be rendered in a strikethrough appearance.

        [MSDN documentation for Font2.StrikeThrough](https://learn.microsoft.com/en-us/office/vba/api/office.font2.strikethrough).
        */
        void SetStrikeThrough(MsoTriState strikeThrough);

        /**
        Gets or sets a value specifying that the selected text should be displayed a subscript.

        [MSDN documentation for Font2.Subscript](https://learn.microsoft.com/en-us/office/vba/api/office.font2.subscript).
        */
        MsoTriState GetSubscript();

        /**
        Gets or sets a value specifying that the selected text should be displayed a subscript.

        [MSDN documentation for Font2.Subscript](https://learn.microsoft.com/en-us/office/vba/api/office.font2.subscript).
        */
        void SetSubscript(MsoTriState subscript);

        /**
        Gets or sets a value specifying that the selected text should be displayed a subscript.

        [MSDN documentation for Font2.Superscript](https://learn.microsoft.com/en-us/office/vba/api/office.font2.superscript).
        */
        MsoTriState GetSuperscript();

        /**
        Gets or sets a value specifying that the selected text should be displayed a subscript.

        [MSDN documentation for Font2.Superscript](https://learn.microsoft.com/en-us/office/vba/api/office.font2.superscript).
        */
        void SetSuperscript(MsoTriState superscript);

        /**
        Gets a value specifying the color of the underline for the selected text.

        [MSDN documentation for Font2.UnderlineColor](https://learn.microsoft.com/en-us/office/vba/api/office.font2.underlinecolor).
        */
        wxExcelColorFormat GetUnderlineColor();

        /**
        Gets or sets a value specifying the underline style for the selected text.

        [MSDN documentation for Font2.UnderlineStyle](https://learn.microsoft.com/en-us/office/vba/api/office.font2.underlinestyle).
        */
        MsoTextUnderlineType GetUnderlineStyle();

        /**
        Gets or sets a value specifying the underline style for the selected text.

        [MSDN documentation for Font2.UnderlineStyle](https://learn.microsoft.com/en-us/office/vba/api/office.font2.underlinestyle).
        */
        void SetUnderlineStyle(MsoTextUnderlineType underlineStyle);

        /**
        Gets or sets a value specifying the text effect for the selected text.

        [MSDN documentation for Font2.WordArtformat](https://learn.microsoft.com/en-us/office/vba/api/office.font2.wordartformat).
        */
        MsoPresetTextEffect GetWordArtformat();

        /**
        Gets or sets a value specifying the text effect for the selected text.

        [MSDN documentation for Font2.WordArtformat](https://learn.microsoft.com/en-us/office/vba/api/office.font2.wordartformat).
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
