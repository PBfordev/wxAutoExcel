/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
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

        [MSDN documentation for TextEffectFormat.ToggleVerticalText](http://msdn.microsoft.com/en-us/library/bb178163).
        */
        void ToggleVerticalText();

        // ***** PROPERTIES *****

        /**
        Returns an MsoTextEffectAlignment value that represents the alignment for WordArt.

        [MSDN documentation for TextEffectFormat.Alignment](http://msdn.microsoft.com/en-us/library/bb238220).
        */
        MsoTextEffectAlignment  GetAlignment();

        /**
        True if the font in the specified WordArt is bold. Read/write MsoTriState.

        [MSDN documentation for TextEffectFormat.FontBold](http://msdn.microsoft.com/en-us/library/bb208520).
        */
        MsoTriState GetFontBold();

        /**
        True if the font in the specified WordArt is bold. Read/write MsoTriState.

        [MSDN documentation for TextEffectFormat.FontBold](http://msdn.microsoft.com/en-us/library/bb208520).
        */
        void SetFontBold(MsoTriState fontBold);

        /**
        Returns msoTrue if the font in the specified WordArt is italic. Read/write MsoTriState.

        [MSDN documentation for TextEffectFormat.FontItalic](http://msdn.microsoft.com/en-us/library/bb208521).
        */
        MsoTriState GetFontItalic();

        /**
        Returns msoTrue if the font in the specified WordArt is italic. Read/write MsoTriState.

        [MSDN documentation for TextEffectFormat.FontItalic](http://msdn.microsoft.com/en-us/library/bb208521).
        */
        void SetFontItalic(MsoTriState fontItalic);

        /**
        Returns the name of the font in the specified WordArt.

        [MSDN documentation for TextEffectFormat.FontName](http://msdn.microsoft.com/en-us/library/bb208522).
        */
        wxString GetFontName();

        /**
        Sets the name of the font in the specified WordArt.

        [MSDN documentation for TextEffectFormat.FontName](http://msdn.microsoft.com/en-us/library/bb208522).
        */
        void SetFontName(const wxString& fontName);

        /**
        Returns the font size for the specified WordArt, in points. Read/write Single.

        [MSDN documentation for TextEffectFormat.FontSize](http://msdn.microsoft.com/en-us/library/bb208524).
        */
        double GetFontSize();

        /**
        Sets the font size for the specified WordArt, in points. Read/write Single.

        [MSDN documentation for TextEffectFormat.FontSize](http://msdn.microsoft.com/en-us/library/bb208524).
        */
        void SetFontSize(double fontSize);

        /**
        True if character pairs in the specified WordArt are kerned. Read/write MsoTriState.

        [MSDN documentation for TextEffectFormat.KernedPairs](http://msdn.microsoft.com/en-us/library/bb177813).
        */
        MsoTriState GetKernedPairs();

        /**
        True if character pairs in the specified WordArt are kerned. Read/write MsoTriState.

        [MSDN documentation for TextEffectFormat.KernedPairs](http://msdn.microsoft.com/en-us/library/bb177813).
        */
        void SetKernedPairs(MsoTriState kernedPairs);

        /**
        True if all characters (both uppercase and lowercase) in the specified WordArt are the same height. Read/write MsoTriState.

        [MSDN documentation for TextEffectFormat.NormalizedHeight](http://msdn.microsoft.com/en-us/library/bb208828).
        */
        MsoTriState GetNormalizedHeight();

        /**
        True if all characters (both uppercase and lowercase) in the specified WordArt are the same height. Read/write MsoTriState.

        [MSDN documentation for TextEffectFormat.NormalizedHeight](http://msdn.microsoft.com/en-us/library/bb208828).
        */
        void SetNormalizedHeight(MsoTriState normalizedHeight);

        /**
        Returns the shape of the specified WordArt. Read/write MsoPresetTextEffectShape.

        [MSDN documentation for TextEffectFormat.PresetShape](http://msdn.microsoft.com/en-us/library/bb208988).
        */
        MsoPresetTextEffectShape GetPresetShape();

        /**
        Sets the shape of the specified WordArt. Read/write MsoPresetTextEffectShape.

        [MSDN documentation for TextEffectFormat.PresetShape](http://msdn.microsoft.com/en-us/library/bb208988).
        */
        void SetPresetShape(MsoPresetTextEffectShape presetShape);

        /**
        Returns the style of the specified WordArt. Read/write MsoPresetTextEffect.

        [MSDN documentation for TextEffectFormat.PresetTextEffect](http://msdn.microsoft.com/en-us/library/bb208990).
        */
        MsoPresetTextEffect GetPresetTextEffect();

        /**
        Sets the style of the specified WordArt. Read/write MsoPresetTextEffect.

        [MSDN documentation for TextEffectFormat.PresetTextEffect](http://msdn.microsoft.com/en-us/library/bb208990).
        */
        void SetPresetTextEffect(MsoPresetTextEffect presetTextEffect);

        /**
        True if characters in the specified WordArt are rotated 90 degrees relative to the WordArt's bounding shape. False if characters in the specified WordArt retain their original orientation relative to the bounding shape. Read/write MsoTriState.

        [MSDN documentation for TextEffectFormat.RotatedChars](http://msdn.microsoft.com/en-us/library/bb221530).
        */
        MsoTriState GetRotatedChars();

        /**
        True if characters in the specified WordArt are rotated 90 degrees relative to the WordArt's bounding shape. False if characters in the specified WordArt retain their original orientation relative to the bounding shape. Read/write MsoTriState.

        [MSDN documentation for TextEffectFormat.RotatedChars](http://msdn.microsoft.com/en-us/library/bb221530).
        */
        void SetRotatedChars(MsoTriState rotatedChars);

        /**
        Returns the text for the specified object.

        [MSDN documentation for TextEffectFormat.Text](http://msdn.microsoft.com/en-us/library/bb215151).
        */
        wxString GetText();

        /**
        Sets the text for the specified object.

        [MSDN documentation for TextEffectFormat.Text](http://msdn.microsoft.com/en-us/library/bb215151).
        */
        void SetText(const wxString& text);

        /**
        Returns the ratio of the horizontal space allotted to each character in the specified WordArt to the width of the character. Can be a value from 0 (zero) through 5. (Large values for this property specify ample space between characters; values less than 1 can produce character overlap.) Read/write Single.

        [MSDN documentation for TextEffectFormat.Tracking](http://msdn.microsoft.com/en-us/library/bb221872).
        */
        double GetTracking();

        /**
        Sets the ratio of the horizontal space allotted to each character in the specified WordArt to the width of the character. Can be a value from 0 (zero) through 5. (Large values for this property specify ample space between characters; values less than 1 can produce character overlap.) Read/write Single.

        [MSDN documentation for TextEffectFormat.Tracking](http://msdn.microsoft.com/en-us/library/bb221872).
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
