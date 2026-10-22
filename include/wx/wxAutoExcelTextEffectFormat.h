/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// License:     MIT license
/////////////////////////////////////////////////////////////////////////////


#ifndef _WXAUTOEXCEL_TEXTEFFECTFORMAT_H
#define _WXAUTOEXCEL_TEXTEFFECTFORMAT_H

#include "wx/wxAutoExcel_defs.h"

#if WXAUTOEXCEL_USE_SHAPES

#include "wx/wxAutoExcel_object.h"


namespace wxAutoExcel {

    /**
    @brief Represents Microsoft Excel TextEffectFormat object.
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelTextEffectFormat : public wxExcelObject
    {
    public:
        // ***** METHODS *****

        /**
        Switches the text flow in the specified WordArt from horizontal to vertical, or vice versa.

        [MSDN documentation for TextEffectFormat.ToggleVerticalText](https://learn.microsoft.com/en-us/office/vba/api/excel.texteffectformat.toggleverticaltext).
        */
        void ToggleVerticalText();

        // ***** PROPERTIES *****

        /**
        Returns an MsoTextEffectAlignment value that represents the alignment for WordArt.

        [MSDN documentation for TextEffectFormat.Alignment](https://learn.microsoft.com/en-us/office/vba/api/excel.texteffectformat.alignment).
        */
        MsoTextEffectAlignment  GetAlignment();

        /**
        True if the font in the specified WordArt is bold. Read/write MsoTriState.

        [MSDN documentation for TextEffectFormat.FontBold](https://learn.microsoft.com/en-us/office/vba/api/excel.texteffectformat.fontbold).
        */
        MsoTriState GetFontBold();

        /**
        True if the font in the specified WordArt is bold. Read/write MsoTriState.

        [MSDN documentation for TextEffectFormat.FontBold](https://learn.microsoft.com/en-us/office/vba/api/excel.texteffectformat.fontbold).
        */
        void SetFontBold(MsoTriState fontBold);

        /**
        Returns msoTrue if the font in the specified WordArt is italic. Read/write MsoTriState.

        [MSDN documentation for TextEffectFormat.FontItalic](https://learn.microsoft.com/en-us/office/vba/api/excel.texteffectformat.fontitalic).
        */
        MsoTriState GetFontItalic();

        /**
        Returns msoTrue if the font in the specified WordArt is italic. Read/write MsoTriState.

        [MSDN documentation for TextEffectFormat.FontItalic](https://learn.microsoft.com/en-us/office/vba/api/excel.texteffectformat.fontitalic).
        */
        void SetFontItalic(MsoTriState fontItalic);

        /**
        Returns the name of the font in the specified WordArt.

        [MSDN documentation for TextEffectFormat.FontName](https://learn.microsoft.com/en-us/office/vba/api/excel.texteffectformat.fontname).
        */
        wxString GetFontName();

        /**
        Sets the name of the font in the specified WordArt.

        [MSDN documentation for TextEffectFormat.FontName](https://learn.microsoft.com/en-us/office/vba/api/excel.texteffectformat.fontname).
        */
        void SetFontName(const wxString& fontName);

        /**
        Returns the font size for the specified WordArt, in points. Read/write Single.

        [MSDN documentation for TextEffectFormat.FontSize](https://learn.microsoft.com/en-us/office/vba/api/excel.texteffectformat.fontsize).
        */
        double GetFontSize();

        /**
        Sets the font size for the specified WordArt, in points. Read/write Single.

        [MSDN documentation for TextEffectFormat.FontSize](https://learn.microsoft.com/en-us/office/vba/api/excel.texteffectformat.fontsize).
        */
        void SetFontSize(double fontSize);

        /**
        True if character pairs in the specified WordArt are kerned. Read/write MsoTriState.

        [MSDN documentation for TextEffectFormat.KernedPairs](https://learn.microsoft.com/en-us/office/vba/api/excel.texteffectformat.kernedpairs).
        */
        MsoTriState GetKernedPairs();

        /**
        True if character pairs in the specified WordArt are kerned. Read/write MsoTriState.

        [MSDN documentation for TextEffectFormat.KernedPairs](https://learn.microsoft.com/en-us/office/vba/api/excel.texteffectformat.kernedpairs).
        */
        void SetKernedPairs(MsoTriState kernedPairs);

        /**
        True if all characters (both uppercase and lowercase) in the specified WordArt are the same height. Read/write MsoTriState.

        [MSDN documentation for TextEffectFormat.NormalizedHeight](https://learn.microsoft.com/en-us/office/vba/api/excel.texteffectformat.normalizedheight).
        */
        MsoTriState GetNormalizedHeight();

        /**
        True if all characters (both uppercase and lowercase) in the specified WordArt are the same height. Read/write MsoTriState.

        [MSDN documentation for TextEffectFormat.NormalizedHeight](https://learn.microsoft.com/en-us/office/vba/api/excel.texteffectformat.normalizedheight).
        */
        void SetNormalizedHeight(MsoTriState normalizedHeight);

        /**
        Returns the shape of the specified WordArt. Read/write MsoPresetTextEffectShape.

        [MSDN documentation for TextEffectFormat.PresetShape](https://learn.microsoft.com/en-us/office/vba/api/excel.texteffectformat.presetshape).
        */
        MsoPresetTextEffectShape GetPresetShape();

        /**
        Sets the shape of the specified WordArt. Read/write MsoPresetTextEffectShape.

        [MSDN documentation for TextEffectFormat.PresetShape](https://learn.microsoft.com/en-us/office/vba/api/excel.texteffectformat.presetshape).
        */
        void SetPresetShape(MsoPresetTextEffectShape presetShape);

        /**
        Returns the style of the specified WordArt. Read/write MsoPresetTextEffect.

        [MSDN documentation for TextEffectFormat.PresetTextEffect](https://learn.microsoft.com/en-us/office/vba/api/excel.texteffectformat.presettexteffect).
        */
        MsoPresetTextEffect GetPresetTextEffect();

        /**
        Sets the style of the specified WordArt. Read/write MsoPresetTextEffect.

        [MSDN documentation for TextEffectFormat.PresetTextEffect](https://learn.microsoft.com/en-us/office/vba/api/excel.texteffectformat.presettexteffect).
        */
        void SetPresetTextEffect(MsoPresetTextEffect presetTextEffect);

        /**
        True if characters in the specified WordArt are rotated 90 degrees relative to the WordArt's bounding shape. False if characters in the specified WordArt retain their original orientation relative to the bounding shape. Read/write MsoTriState.

        [MSDN documentation for TextEffectFormat.RotatedChars](https://learn.microsoft.com/en-us/office/vba/api/excel.texteffectformat.rotatedchars).
        */
        MsoTriState GetRotatedChars();

        /**
        True if characters in the specified WordArt are rotated 90 degrees relative to the WordArt's bounding shape. False if characters in the specified WordArt retain their original orientation relative to the bounding shape. Read/write MsoTriState.

        [MSDN documentation for TextEffectFormat.RotatedChars](https://learn.microsoft.com/en-us/office/vba/api/excel.texteffectformat.rotatedchars).
        */
        void SetRotatedChars(MsoTriState rotatedChars);

        /**
        Returns the text for the specified object.

        [MSDN documentation for TextEffectFormat.Text](https://learn.microsoft.com/en-us/office/vba/api/excel.texteffectformat.text).
        */
        wxString GetText();

        /**
        Sets the text for the specified object.

        [MSDN documentation for TextEffectFormat.Text](https://learn.microsoft.com/en-us/office/vba/api/excel.texteffectformat.text).
        */
        void SetText(const wxString& text);

        /**
        Returns the ratio of the horizontal space allotted to each character in the specified WordArt to the width of the character. Can be a value from 0 (zero) through 5. (Large values for this property specify ample space between characters; values less than 1 can produce character overlap.) Read/write Single.

        [MSDN documentation for TextEffectFormat.Tracking](https://learn.microsoft.com/en-us/office/vba/api/excel.texteffectformat.tracking).
        */
        double GetTracking();

        /**
        Sets the ratio of the horizontal space allotted to each character in the specified WordArt to the width of the character. Can be a value from 0 (zero) through 5. (Large values for this property specify ample space between characters; values less than 1 can produce character overlap.) Read/write Single.

        [MSDN documentation for TextEffectFormat.Tracking](https://learn.microsoft.com/en-us/office/vba/api/excel.texteffectformat.tracking).
        */
        void SetTracking(double tracking);


        /**
        Returns "TextEffectFormat".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("TextEffectFormat"); }
    };


} // namespace wxAutoExcel

#endif // #if WXAUTOEXCEL_USE_SHAPES

#endif //_WXAUTOEXCEL_TEXTEFFECTFORMAT_H
