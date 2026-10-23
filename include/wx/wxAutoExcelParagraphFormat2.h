/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// License:     MIT license
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

        [MSDN documentation for ParagraphFormat2.Alignment](https://learn.microsoft.com/en-us/office/vba/api/office.paragraphformat2.alignment).
        */
        MsoParagraphAlignment GetAlignment();

        /**
        Sets a value specifying the alignment of the paragraph.

        [MSDN documentation for ParagraphFormat2.Alignment](https://learn.microsoft.com/en-us/office/vba/api/office.paragraphformat2.alignment).
        */
        void SetAlignment(MsoParagraphAlignment alignment);


        /**
        Gets or sets a constant that represents the vertical position of fonts in a paragraph.

        [MSDN documentation for ParagraphFormat2.BaselineAlignment](https://learn.microsoft.com/en-us/office/vba/api/office.paragraphformat2.baselinealignment).
        */
        MsoBaselineAlignment GetBaselineAlignment();

        /**
        Gets or sets a constant that represents the vertical position of fonts in a paragraph.

        [MSDN documentation for ParagraphFormat2.BaselineAlignment](https://learn.microsoft.com/en-us/office/vba/api/office.paragraphformat2.baselinealignment).
        */
        void SetBaselineAlignment(MsoBaselineAlignment baselineAlignment);

        /**
        Gets a BulletFormat2 object for the paragraph.

        [MSDN documentation for ParagraphFormat2.Bullet](https://learn.microsoft.com/en-us/office/vba/api/office.paragraphformat2.bullet).
        */
        wxExcelBulletFormat2 GetBullet();

        /**
        Gets or sets the East Asian line break control level for the specified paragraph.

        [MSDN documentation for ParagraphFormat2.FarEastLineBreakLevel](https://learn.microsoft.com/en-us/office/vba/api/office.paragraphformat2.fareastlinebreaklevel).
        */
        MsoTriState GetFarEastLineBreakLevel();

        /**
        Gets or sets the East Asian line break control level for the specified paragraph.

        [MSDN documentation for ParagraphFormat2.FarEastLineBreakLevel](https://learn.microsoft.com/en-us/office/vba/api/office.paragraphformat2.fareastlinebreaklevel).
        */
        void SetFarEastLineBreakLevel(MsoTriState farEastLineBreakLevel);

        /**
        Gets or sets the value (in points) for a first line or hanging indent.

        [MSDN documentation for ParagraphFormat2.FirstLineIndent](https://learn.microsoft.com/en-us/office/vba/api/office.paragraphformat2.firstlineindent).
        */
        double GetFirstLineIndent();

        /**
        Gets or sets the value (in points) for a first line or hanging indent.

        [MSDN documentation for ParagraphFormat2.FirstLineIndent](https://learn.microsoft.com/en-us/office/vba/api/office.paragraphformat2.firstlineindent).
        */
        void SetFirstLineIndent(double firstLineIndent);

        /**
        Determines whether hanging punctuation is enabled for the specified paragraphs.

        [MSDN documentation for ParagraphFormat2.HangingPunctuation](https://learn.microsoft.com/en-us/office/vba/api/office.paragraphformat2.hangingpunctuation).
        */
        MsoTriState GetHangingPunctuation();

        /**
        Determines whether hanging punctuation is enabled for the specified paragraphs.

        [MSDN documentation for ParagraphFormat2.HangingPunctuation](https://learn.microsoft.com/en-us/office/vba/api/office.paragraphformat2.hangingpunctuation).
        */
        void SetHangingPunctuation(MsoTriState hangingPunctuation);

        /**
        Gets or sets a value representing the indent level assigned to text in the selected paragraph.

        [MSDN documentation for ParagraphFormat2.IndentLevel](https://learn.microsoft.com/en-us/office/vba/api/office.paragraphformat2.indentlevel).
        */
        long GetIndentLevel();

        /**
        Gets or sets a value representing the indent level assigned to text in the selected paragraph.

        [MSDN documentation for ParagraphFormat2.IndentLevel](https://learn.microsoft.com/en-us/office/vba/api/office.paragraphformat2.indentlevel).
        */
        void SetIndentLevel(long indentLevel);

        /**
        Gets or sets a value that represents the left indent value (in points) for the specified paragraphs.

        [MSDN documentation for ParagraphFormat2.LeftIndent](https://learn.microsoft.com/en-us/office/vba/api/office.paragraphformat2.leftindent).
        */
        double GetLeftIndent();

        /**
        Gets or sets a value that represents the left indent value (in points) for the specified paragraphs.

        [MSDN documentation for ParagraphFormat2.LeftIndent](https://learn.microsoft.com/en-us/office/vba/api/office.paragraphformat2.leftindent).
        */
        void SetLeftIndent(double leftIndent);

        /**
        Determines whether line spacing after the last line in each paragraph is set to a specific number of points or lines.

        [MSDN documentation for ParagraphFormat2.LineRuleAfter](https://learn.microsoft.com/en-us/office/vba/api/office.paragraphformat2.lineruleafter).
        */
        MsoTriState GetLineRuleAfter();

        /**
        Determines whether line spacing after the last line in each paragraph is set to a specific number of points or lines.

        [MSDN documentation for ParagraphFormat2.LineRuleAfter](https://learn.microsoft.com/en-us/office/vba/api/office.paragraphformat2.lineruleafter).
        */
        void SetLineRuleAfter(MsoTriState lineRuleAfter);

        /**
        Determines whether line spacing before the first line in each paragraph is set to a specific number of points or lines.

        [MSDN documentation for ParagraphFormat2.LineRuleBefore](https://learn.microsoft.com/en-us/office/vba/api/office.paragraphformat2.linerulebefore).
        */
        MsoTriState GetLineRuleBefore();

        /**
        Determines whether line spacing before the first line in each paragraph is set to a specific number of points or lines.

        [MSDN documentation for ParagraphFormat2.LineRuleBefore](https://learn.microsoft.com/en-us/office/vba/api/office.paragraphformat2.linerulebefore).
        */
        void SetLineRuleBefore(MsoTriState lineRuleBefore);

        /**
        Determines whether line spacing between base lines is set to a specific number of points or lines.

        [MSDN documentation for ParagraphFormat2.LineRuleWithin](https://learn.microsoft.com/en-us/office/vba/api/office.paragraphformat2.linerulewithin).
        */
        MsoTriState GetLineRuleWithin();

        /**
        Determines whether line spacing between base lines is set to a specific number of points or lines.

        [MSDN documentation for ParagraphFormat2.LineRuleWithin](https://learn.microsoft.com/en-us/office/vba/api/office.paragraphformat2.linerulewithin).
        */
        void SetLineRuleWithin(MsoTriState lineRuleWithin);

        /**
        Gets or sets the right indent (in points) for the specified paragraphs.

        [MSDN documentation for ParagraphFormat2.RightIndent](https://learn.microsoft.com/en-us/office/vba/api/office.paragraphformat2.rightindent).
        */
        double GetRightIndent();

        /**
        Gets or sets the right indent (in points) for the specified paragraphs.

        [MSDN documentation for ParagraphFormat2.RightIndent](https://learn.microsoft.com/en-us/office/vba/api/office.paragraphformat2.rightindent).
        */
        void SetRightIndent(double rightIndent);

        /**
        Gets or sets the amount of spacing (in points) after the specified paragraph.

        [MSDN documentation for ParagraphFormat2.SpaceAfter](https://learn.microsoft.com/en-us/office/vba/api/office.paragraphformat2.spaceafter).
        */
        double GetSpaceAfter();

        /**
        Gets or sets the amount of spacing (in points) after the specified paragraph.

        [MSDN documentation for ParagraphFormat2.SpaceAfter](https://learn.microsoft.com/en-us/office/vba/api/office.paragraphformat2.spaceafter).
        */
        void SetSpaceAfter(double spaceAfter);

        /**
        Gets or sets the spacing (in points) before the specified paragraphs.

        [MSDN documentation for ParagraphFormat2.SpaceBefore](https://learn.microsoft.com/en-us/office/vba/api/office.paragraphformat2.spacebefore).
        */
        double GetSpaceBefore();

        /**
        Gets or sets the spacing (in points) before the specified paragraphs.

        [MSDN documentation for ParagraphFormat2.SpaceBefore](https://learn.microsoft.com/en-us/office/vba/api/office.paragraphformat2.spacebefore).
        */
        void SetSpaceBefore(double spaceBefore);

        /**
        Gets or sets the amount of space between base lines in the specified paragraph, in points or lines.

        [MSDN documentation for ParagraphFormat2.SpaceWithin](https://learn.microsoft.com/en-us/office/vba/api/office.paragraphformat2.spacewithin).
        */
        double GetSpaceWithin();

        /**
        Gets or sets the amount of space between base lines in the specified paragraph, in points or lines.

        [MSDN documentation for ParagraphFormat2.SpaceWithin](https://learn.microsoft.com/en-us/office/vba/api/office.paragraphformat2.spacewithin).
        */
        void SetSpaceWithin(double spaceWithin);

        /**
        Gets a TabStops2 collection that represents all the custom tab stops for the specified paragraphs.

        [MSDN documentation for ParagraphFormat2.TabStops](https://learn.microsoft.com/en-us/office/vba/api/office.paragraphformat2.tabstops).
        */
        wxExcelTabStops2 GetTabStops();

        /**
        Gets or sets the text direction for the specified paragraph.

        [MSDN documentation for ParagraphFormat2.TextDirection](https://learn.microsoft.com/en-us/office/vba/api/office.paragraphformat2.textdirection).
        */
        MsoTextDirection GetTextDirection();

        /**
        Gets or sets the text direction for the specified paragraph.

        [MSDN documentation for ParagraphFormat2.TextDirection](https://learn.microsoft.com/en-us/office/vba/api/office.paragraphformat2.textdirection).
        */
        void SetTextDirection(MsoTextDirection textDirection);

        /**
        Determines whether the application wraps the Latin text in the middle of a word in the specified paragraphs.

        [MSDN documentation for ParagraphFormat2.WordWrap](https://learn.microsoft.com/en-us/office/vba/api/office.paragraphformat2.wordwrap).
        */
        MsoTriState GetWordWrap();

        /**
        Determines whether the application wraps the Latin text in the middle of a word in the specified paragraphs.

        [MSDN documentation for ParagraphFormat2.WordWrap](https://learn.microsoft.com/en-us/office/vba/api/office.paragraphformat2.wordwrap).
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
