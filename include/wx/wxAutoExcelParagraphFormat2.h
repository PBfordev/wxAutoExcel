/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Modified by:
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////


#ifndef _WXAUTOEXCEL_PARAGRAPHFORMAT2_H
#define _WXAUTOEXCEL_PARAGRAPHFORMAT2_H

#include "wx/wxAutoExcel_defs.h"

#if WXAUTOEXCEL_USE_SHAPES || WXAUTOEXCEL_USE_CHARTS

#include "wx/wxAutoExcel_object.h"
#include "wx/wxAutoExcel_enums.h"

namespace wxAutoExcel {

    /**
    @brief Represents Microsoft Excel ParagraphFormat2 object.
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelParagraphFormat2 : public wxExcelObject
    {
    public:
        // ***** PROPERTIES *****

        /**
        Gets or sets a value specifying the alignment of the paragraph.

        [MSDN documentation for ParagraphFormat2.Alignment](http://msdn.microsoft.com/en-us/library/aa434571.aspx).
        */
        MsoParagraphAlignment GetAlignment();


        /**
        Gets or sets a constant that represents the vertical position of fonts in a paragraph.

        [MSDN documentation for ParagraphFormat2.BaselineAlignment](http://msdn.microsoft.com/en-us/library/aa434573.aspx).
        */
        MsoBaselineAlignment GetBaselineAlignment();

        /**
        Gets or sets a constant that represents the vertical position of fonts in a paragraph.

        [MSDN documentation for ParagraphFormat2.BaselineAlignment](http://msdn.microsoft.com/en-us/library/aa434573.aspx).
        */
        void SetBaselineAlignment(MsoBaselineAlignment baselineAlignment);

        /**
        Gets a BulletFormat2 object for the paragraph.

        [MSDN documentation for ParagraphFormat2.Bullet](http://msdn.microsoft.com/en-us/library/aa434574.aspx).
        */
        wxExcelBulletFormat2 GetBullet();

        /**
        Gets or sets the East Asian line break control level for the specified paragraph.

        [MSDN documentation for ParagraphFormat2.FarEastLineBreakLevel](http://msdn.microsoft.com/en-us/library/aa434576.aspx).
        */
        MsoTriState GetFarEastLineBreakLevel();

        /**
        Gets or sets the East Asian line break control level for the specified paragraph.

        [MSDN documentation for ParagraphFormat2.FarEastLineBreakLevel](http://msdn.microsoft.com/en-us/library/aa434576.aspx).
        */
        void SetFarEastLineBreakLevel(MsoTriState farEastLineBreakLevel);

        /**
        Gets or sets the value (in points) for a first line or hanging indent.

        [MSDN documentation for ParagraphFormat2.FirstLineIndent](http://msdn.microsoft.com/en-us/library/aa434577.aspx).
        */
        double GetFirstLineIndent();

        /**
        Gets or sets the value (in points) for a first line or hanging indent.

        [MSDN documentation for ParagraphFormat2.FirstLineIndent](http://msdn.microsoft.com/en-us/library/aa434577.aspx).
        */
        void SetFirstLineIndent(double firstLineIndent);

        /**
        Determines whether hanging punctuation is enabled for the specified paragraphs.

        [MSDN documentation for ParagraphFormat2.HangingPunctuation](http://msdn.microsoft.com/en-us/library/aa434578.aspx).
        */
        MsoTriState GetHangingPunctuation();

        /**
        Determines whether hanging punctuation is enabled for the specified paragraphs.

        [MSDN documentation for ParagraphFormat2.HangingPunctuation](http://msdn.microsoft.com/en-us/library/aa434578.aspx).
        */
        void SetHangingPunctuation(MsoTriState hangingPunctuation);

        /**
        Gets or sets a value representing the indent level assigned to text in the selected paragraph.

        [MSDN documentation for ParagraphFormat2.IndentLevel](http://msdn.microsoft.com/en-us/library/aa434579.aspx).
        */
        long GetIndentLevel();

        /**
        Gets or sets a value representing the indent level assigned to text in the selected paragraph.

        [MSDN documentation for ParagraphFormat2.IndentLevel](http://msdn.microsoft.com/en-us/library/aa434579.aspx).
        */
        void SetIndentLevel(long indentLevel);

        /**
        Gets or sets a value that represents the left indent value (in points) for the specified paragraphs.

        [MSDN documentation for ParagraphFormat2.LeftIndent](http://msdn.microsoft.com/en-us/library/aa434580.aspx).
        */
        double GetLeftIndent();

        /**
        Gets or sets a value that represents the left indent value (in points) for the specified paragraphs.

        [MSDN documentation for ParagraphFormat2.LeftIndent](http://msdn.microsoft.com/en-us/library/aa434580.aspx).
        */
        void SetLeftIndent(double leftIndent);

        /**
        Determines whether line spacing after the last line in each paragraph is set to a specific number of points or lines.

        [MSDN documentation for ParagraphFormat2.LineRuleAfter](http://msdn.microsoft.com/en-us/library/aa434581.aspx).
        */
        MsoTriState GetLineRuleAfter();

        /**
        Determines whether line spacing after the last line in each paragraph is set to a specific number of points or lines.

        [MSDN documentation for ParagraphFormat2.LineRuleAfter](http://msdn.microsoft.com/en-us/library/aa434581.aspx).
        */
        void SetLineRuleAfter(MsoTriState lineRuleAfter);

        /**
        Determines whether line spacing before the first line in each paragraph is set to a specific number of points or lines.

        [MSDN documentation for ParagraphFormat2.LineRuleBefore](http://msdn.microsoft.com/en-us/library/aa434582.aspx).
        */
        MsoTriState GetLineRuleBefore();

        /**
        Determines whether line spacing before the first line in each paragraph is set to a specific number of points or lines.

        [MSDN documentation for ParagraphFormat2.LineRuleBefore](http://msdn.microsoft.com/en-us/library/aa434582.aspx).
        */
        void SetLineRuleBefore(MsoTriState lineRuleBefore);

        /**
        Determines whether line spacing between base lines is set to a specific number of points or lines.

        [MSDN documentation for ParagraphFormat2.LineRuleWithin](http://msdn.microsoft.com/en-us/library/aa434583.aspx).
        */
        MsoTriState GetLineRuleWithin();

        /**
        Determines whether line spacing between base lines is set to a specific number of points or lines.

        [MSDN documentation for ParagraphFormat2.LineRuleWithin](http://msdn.microsoft.com/en-us/library/aa434583.aspx).
        */
        void SetLineRuleWithin(MsoTriState lineRuleWithin);

        /**
        Gets or sets the right indent (in points) for the specified paragraphs.

        [MSDN documentation for ParagraphFormat2.RightIndent](http://msdn.microsoft.com/en-us/library/aa434585.aspx).
        */
        double GetRightIndent();

        /**
        Gets or sets the right indent (in points) for the specified paragraphs.

        [MSDN documentation for ParagraphFormat2.RightIndent](http://msdn.microsoft.com/en-us/library/aa434585.aspx).
        */
        void SetRightIndent(double rightIndent);

        /**
        Gets or sets the amount of spacing (in points) after the specified paragraph.

        [MSDN documentation for ParagraphFormat2.SpaceAfter](http://msdn.microsoft.com/en-us/library/aa434586.aspx).
        */
        double GetSpaceAfter();

        /**
        Gets or sets the amount of spacing (in points) after the specified paragraph.

        [MSDN documentation for ParagraphFormat2.SpaceAfter](http://msdn.microsoft.com/en-us/library/aa434586.aspx).
        */
        void SetSpaceAfter(double spaceAfter);

        /**
        Gets or sets the spacing (in points) before the specified paragraphs.

        [MSDN documentation for ParagraphFormat2.SpaceBefore](http://msdn.microsoft.com/en-us/library/aa434587.aspx).
        */
        double GetSpaceBefore();

        /**
        Gets or sets the spacing (in points) before the specified paragraphs.

        [MSDN documentation for ParagraphFormat2.SpaceBefore](http://msdn.microsoft.com/en-us/library/aa434587.aspx).
        */
        void SetSpaceBefore(double spaceBefore);

        /**
        Gets or sets the amount of space between base lines in the specified paragraph, in points or lines.

        [MSDN documentation for ParagraphFormat2.SpaceWithin](http://msdn.microsoft.com/en-us/library/aa434588.aspx).
        */
        double GetSpaceWithin();

        /**
        Gets or sets the amount of space between base lines in the specified paragraph, in points or lines.

        [MSDN documentation for ParagraphFormat2.SpaceWithin](http://msdn.microsoft.com/en-us/library/aa434588.aspx).
        */
        void SetSpaceWithin(double spaceWithin);

        /**
        Gets a TabStops2 collection that represents all the custom tab stops for the specified paragraphs.

        [MSDN documentation for ParagraphFormat2.TabStops](http://msdn.microsoft.com/en-us/library/aa434589.aspx).
        */
        wxExcelTabStops2 GetTabStops();

        /**
        Gets or sets the text direction for the specified paragraph.

        [MSDN documentation for ParagraphFormat2.TextDirection](http://msdn.microsoft.com/en-us/library/aa434590.aspx).
        */
        MsoTextDirection GetTextDirection();

        /**
        Gets or sets the text direction for the specified paragraph.

        [MSDN documentation for ParagraphFormat2.TextDirection](http://msdn.microsoft.com/en-us/library/aa434590.aspx).
        */
        void SetTextDirection(MsoTextDirection textDirection);

        /**
        Determines whether the application wraps the Latin text in the middle of a word in the specified paragraphs.

        [MSDN documentation for ParagraphFormat2.WordWrap](http://msdn.microsoft.com/en-us/library/aa434591.aspx).
        */
        MsoTriState GetWordWrap();

        /**
        Determines whether the application wraps the Latin text in the middle of a word in the specified paragraphs.

        [MSDN documentation for ParagraphFormat2.WordWrap](http://msdn.microsoft.com/en-us/library/aa434591.aspx).
        */
        void SetWordWrap(MsoTriState wordWrap);

        /**
        Returns "ParagraphFormat2".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("ParagraphFormat2"); }
    };


} // namespace wxAutoExcel


#endif // #if WXAUTOEXCEL_USE_SHAPES || WXAUTOEXCEL_USE_CHARTS

#endif //_WXAUTOEXCEL_PARAGRAPHFORMAT2_H
