/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// License:     MIT license
/////////////////////////////////////////////////////////////////////////////


#ifndef _WXAUTOEXCEL_TEXTRANGE2_H
#define _WXAUTOEXCEL_TEXTRANGE2_H

#include "wx/wxAutoExcel_defs.h"

#if WXAUTOEXCEL_USE_SHAPES || WXAUTOEXCEL_USE_CHARTS

#include "wx/wxAutoExcel_object.h"
#include "wx/wxAutoExcel_enums.h"

namespace wxAutoExcel {

    /**
    @brief Represents Microsoft Excel TextRange2 object.
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelTextRange2 : public wxExcelObject
    {
    public:
        // ***** METHODS *****

        /**
        Adds period (.) punctuation to the right side of the text contained in TextRange2 object for left-to-right languages and on the left side for right-to-left languages.

        [MSDN documentation for TextRange2.AddPeriods](https://learn.microsoft.com/en-us/office/vba/api/office.textrange2.addperiods).
        */
        void AddPeriods();

        /**
        Changes the case of a TextRange2 object to one of the values in the MsoTextChangeCase enumeration.

        [MSDN documentation for TextRange2.ChangeCase](https://learn.microsoft.com/en-us/office/vba/api/office.textrange2.changecase).
        */
        void ChangeCase(MsoTextChangeCase type);

        /**
        Copies a TextRange2 object.

        [MSDN documentation for TextRange2.Copy](https://learn.microsoft.com/en-us/office/vba/api/office.textrange2.copy).
        */
        void Copy();

        /**
        Removes a portion or all of the text from a range of text.

        [MSDN documentation for TextRange2.Cut]().
        */
        void Cut();

        /**
        Deletes a TextRange2 object.

        [MSDN documentation for TextRange2.Delete](https://learn.microsoft.com/en-us/office/vba/api/office.textrange2.delete).
        */
        void Delete();

        /**
        Searches a TextRange2 object for a subset of text.

        [MSDN documentation for TextRange2.Find](https://learn.microsoft.com/en-us/office/vba/api/office.textrange2.find).
        */
        wxExcelTextRange2 Find(const wxString& findWhat, long* after = NULL, MsoTriState* matchCase = NULL, MsoTriState* wholeWords = NULL);

        /**
        Inserts text to the right of the existing text in the TextRange2 object.

        [MSDN documentation for TextRange2.InsertAfter](https://learn.microsoft.com/en-us/office/vba/api/office.textrange2.insertafter).
        */
        wxExcelTextRange2 InsertAfter(const wxString& newText);

        /**
        Inserts text to the left of the existing text in the TextRange2 object.

        [MSDN documentation for TextRange2.InsertBefore](https://learn.microsoft.com/en-us/office/vba/api/office.textrange2.insertbefore).
        */
        wxExcelTextRange2 InsertBefore(const wxString& newText);

        /**
        Inserts a field into the body of a data label in a chart.

        [MSDN documentation for TextRange2.InsertChartField](https://learn.microsoft.com/en-us/office/vba/api/office.textrange2.insertchartfield).
        */
        wxExcelTextRange2 InsertChartField(MsoChartFieldType chartFieldType,
                                           const wxString& formula = wxEmptyString,
                                           long* position = NULL);

        /**
        Inserts a symbol from the specified font set into the range of text represented by the TextRange2 object.

        [MSDN documentation for TextRange2.InsertSymbol](https://learn.microsoft.com/en-us/office/vba/api/office.textrange2.insertsymbol).
        */
        wxExcelTextRange2 InsertSymbol(const wxString& fontName, long charNumber, MsoTriState unicode);

        //@{
        /**
        Gets the range of text specified by the index number from the TextRange2 object.

        [MSDN documentation for TextRange2.Item](https://learn.microsoft.com/en-us/office/vba/api/office.textrange2.item).
        */
        wxExcelTextRange2 Item(long index);
        wxExcelTextRange2 operator[](long index);
        //@}

        /**
        Returns a TextRange2 Represents the specified subset of left-to-right text runs. A text run consists of a range of characters that share the same font attributes.

        [MSDN documentation for TextRange2.LtrRun](https://learn.microsoft.com/en-us/office/vba/api/office.textrange2.ltrrun).
        */
        void LtrRun();

        /**
        Pastes the contents of the Clipboard into the TextRange2 object.

        [MSDN documentation for TextRange2.Paste](https://learn.microsoft.com/en-us/office/vba/api/office.textrange2.paste).
        */
        wxExcelTextRange2 Paste();

        /**
        Replaces the text range with the contents of the Clipboard in the format specified. If the paste succeeds, this method returns a TextRange2 object including the text range that was pasted.

        [MSDN documentation for TextRange2.PasteSpecial](https://learn.microsoft.com/en-us/office/vba/api/office.textrange2.pastespecial).
        */
        wxExcelTextRange2 PasteSpecial(MsoClipboardFormat format);

        /**
        Removes all period (.) punctuation from the text in the TextRange2 object.

        [MSDN documentation for TextRange2.RemovePeriods](https://learn.microsoft.com/en-us/office/vba/api/office.textrange2.removeperiods).
        */
        void RemovePeriods();

        /**
        Finds specific text in a text range, replaces the found text with a specified string, and returns a TextRange2 Represents the first occurrence of the found text. Returns Nothing if no match is found.

        [MSDN documentation for TextRange2.Replace](https://learn.microsoft.com/en-us/office/vba/api/office.textrange2.replace).
        */
        wxExcelTextRange2 Replace(const wxString& findWhat, const wxString& replaceWhat,
	                              long* after = NULL, MsoTriState* matchCase = NULL, MsoTriState* wholeWords = NULL);

        /**
        NOT IMPLEMENTED. wxAutomationObject does not support passing arguments by reference.
        
        Gets the coordinates of the vertices of the text bounding box for the specified text range. Read-only.

        [MSDN documentation for TextRange2.RotatedBounds](https://learn.microsoft.com/en-us/office/vba/api/office.textrange2.rotatedbounds).
        */
        void RotatedBounds(double& X1, double& Y1, double& X2, double& Y2,
                           double& X3, double& Y3, double& X4, double& Y4);

        /**
        Sets the direction of text in the text range to read from right to left.

        [MSDN documentation for TextRange2.RtlRun](https://learn.microsoft.com/en-us/office/vba/api/office.textrange2.rtlrun).
        */
        void RtlRun();

        /**
        Selects the TextRange2 object.

        [MSDN documentation for TextRange2.Select](https://learn.microsoft.com/en-us/office/vba/api/office.textrange2.select).
        */
        void Select();

        /**
        Removes the white space on the left and right sides of the text in the TextRange2 object.

        [MSDN documentation for TextRange2.TrimText](https://learn.microsoft.com/en-us/office/vba/api/office.textrange2.trimtext).
        */
        wxExcelTextRange2 TrimText();

        // ***** PROPERTIES *****

        /**
        Gets the height, in points, of the text bounding box for the specified text.

        [MSDN documentation for TextRange2.BoundHeight](https://learn.microsoft.com/en-us/office/vba/api/office.textrange2.boundheight).
        */
        double GetBoundHeight();

        /**
        Gets the left coordinate, in points, of the text bounding box for the specified text.

        [MSDN documentation for TextRange2.BoundLeft](https://learn.microsoft.com/en-us/office/vba/api/office.textrange2.boundleft).
        */
        double GetBoundLeft();

        /**
        Gets the top coordinate, in points, of the text bounding box for the specified text.

        [MSDN documentation for TextRange2.BoundTop](https://learn.microsoft.com/en-us/office/vba/api/office.textrange2.boundtop).
        */
        double GetBoundTop();

        /**
        Gets the width, in points, of the text bounding box for the specified text.

        [MSDN documentation for TextRange2.BoundWidth](https://learn.microsoft.com/en-us/office/vba/api/office.textrange2.boundwidth).
        */
        double GetBoundWidth();

        /**        

        [MSDN documentation for TextRange2.Characters](https://learn.microsoft.com/en-us/office/vba/api/office.textrange2.characters).
        */
        wxExcelTextRange2 GetCharacters(long start = 1, long* length = NULL);

        /**
        Gets a Long indicating the number of items in the TextRange2 collection.

        [MSDN documentation for TextRange2.Count](https://learn.microsoft.com/en-us/office/vba/api/office.textrange2.count).
        */
        long GetCount();

        /**
        Returns a Font Represents character formatting for the TextRange2 object.

        [MSDN documentation for TextRange2.Font](https://learn.microsoft.com/en-us/office/vba/api/office.textrange2.font).
        */
        wxExcelFont2 GetFont();

        /**
        Gets or sets the MsoLanguageID value of the TextRange2 object.

        [MSDN documentation for TextRange2.LanguageID](https://learn.microsoft.com/en-us/office/vba/api/office.textrange2.languageid).
        */
        MsoLanguageID GetLanguageID();

        /**
        Gets or sets the MsoLanguageID value of the TextRange2 object.

        [MSDN documentation for TextRange2.LanguageID](https://learn.microsoft.com/en-us/office/vba/api/office.textrange2.languageid).
        */
        void SetLanguageID(MsoLanguageID languageID);

        /**
        Get a Long that represents the length of a text range.

        [MSDN documentation for TextRange2.Length](https://learn.microsoft.com/en-us/office/vba/api/office.textrange2.length).
        */
        long GetLength();

        /**
        Returns a TextRange2 Represents the specified subset of text lines.

        [MSDN documentation for TextRange2.Lines](https://learn.microsoft.com/en-us/office/vba/api/office.textrange2.lines).
        */
        wxExcelTextRange2 GetLines();

        /**
        Returns the math zones of the text range. With the default values of start and length,
        returns the whole text range; with other values, fails if there is no math zone there.

        [MSDN documentation for TextRange2.MathZones](https://learn.microsoft.com/en-us/office/vba/api/office.textrange2.mathzones).
        */
        wxExcelTextRange2 GetMathZones(long start = 1, long* length = NULL);

        /**
        Returns a ParagraphFormat Represents paragraph formatting for the specified text.

        [MSDN documentation for TextRange2.ParagraphFormat](https://learn.microsoft.com/en-us/office/vba/api/office.textrange2.paragraphformat).
        */
        wxExcelParagraphFormat2 GetParagraphFormat();

        /**
        Gets a TextRange2 Represents the specified subset of text paragraphs.

        [MSDN documentation for TextRange2.Paragraphs](https://learn.microsoft.com/en-us/office/vba/api/office.textrange2.paragraphs).
        */
        wxExcelTextRange2 GetParagraphs();


        /**
        Gets a TextRange2 Represents the specified subset of text runs. A text run consists of a range of characters that share the same font attributes.

        [MSDN documentation for TextRange2.Runs](https://learn.microsoft.com/en-us/office/vba/api/office.textrange2.runs).
        */
        wxExcelTextRange2 GetRuns();

        /**
        Returns a TextRange2 Represents the specified subset of text sentences.

        [MSDN documentation for TextRange2.Sentences](https://learn.microsoft.com/en-us/office/vba/api/office.textrange2.sentences).
        */
        wxExcelTextRange2 GetSentences();

        /**
        Gets a Long value indicating the starting point of the specified text range.

        [MSDN documentation for TextRange2.Start](https://learn.microsoft.com/en-us/office/vba/api/office.textrange2.start).
        */
        long GetStart();

        /**
        Gets or sets a String value that represents the text in a text range.

        [MSDN documentation for TextRange2.Text](https://learn.microsoft.com/en-us/office/vba/api/office.textrange2.text).
        */
        wxString GetText();

        /**
        Gets or sets a String value that represents the text in a text range.

        [MSDN documentation for TextRange2.Text](https://learn.microsoft.com/en-us/office/vba/api/office.textrange2.text).
        */
        void SetText(const wxString& text);

        /**
        Gets a TextRange2 Represents the specified subset of text words.

        [MSDN documentation for TextRange2.Words](https://learn.microsoft.com/en-us/office/vba/api/office.textrange2.words).
        */
        wxExcelTextRange2 GetWords();

        /**
        Returns "TextRange2".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("TextRange2"); }
    };


} // namespace wxAutoExcel

#endif // #if WXAUTOEXCEL_USE_SHAPES || WXAUTOEXCEL_USE_CHARTS

#endif //_WXAUTOEXCEL_TEXTRANGE2_H
