/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
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

        [MSDN documentation for TextRange2.AddPeriods](http://msdn.microsoft.com/en-us/library/aa434141.aspx).
        */
        void AddPeriods();

        /**
        Changes the case of a TextRange2 object to one of the values in the MsoTextChangeCase enumeration.

        [MSDN documentation for TextRange2.ChangeCase](http://msdn.microsoft.com/en-us/library/aa434142.aspx).
        */
        void ChangeCase(MsoTextChangeCase type);

        /**
        Copies a TextRange2 object.

        [MSDN documentation for TextRange2.Copy](http://msdn.microsoft.com/en-us/library/aa434143.aspx).
        */
        void Copy();

        /**
        Removes a portion or all of the text from a range of text.

        [MSDN documentation for TextRange2.Cut]().
        */
        void Cut();

        /**
        Deletes a TextRange2 object.

        [MSDN documentation for TextRange2.Delete](http://msdn.microsoft.com/en-us/library/aa434145.aspx).
        */
        void Delete();

        /**
        Searches a TextRange2 object for a subset of text.

        [MSDN documentation for TextRange2.Find](http://msdn.microsoft.com/en-us/library/aa434146.aspx).
        */
        wxExcelTextRange2 Find(const wxString& findWhat, long* after = NULL, MsoTriState* matchCase = NULL, MsoTriState* wholeWords = NULL);

        /**
        Inserts text to the right of the existing text in the TextRange2 object.

        [MSDN documentation for TextRange2.InsertAfter](http://msdn.microsoft.com/en-us/library/aa434147.aspx).
        */
        wxExcelTextRange2 InsertAfter(const wxString& newText);

        /**
        Inserts text to the left of the existing text in the TextRange2 object.

        [MSDN documentation for TextRange2.InsertBefore](http://msdn.microsoft.com/en-us/library/aa434148.aspx).
        */
        wxExcelTextRange2 InsertBefore(const wxString& newText);

        /**
        Inserts a field into the body of a data label in a chart.

        [MSDN documentation for TextRange2.InsertChartField](http://msdn.microsoft.com/en-us/library/aa434148.aspx).
        */
        wxExcelTextRange2 InsertChartField(MsoChartFieldType chartFieldType,
                                           const wxString& formula = wxEmptyString,
                                           long* position = NULL);

        /**
        Inserts a symbol from the specified font set into the range of text represented by the TextRange2 object.

        [MSDN documentation for TextRange2.InsertSymbol](http://msdn.microsoft.com/en-us/library/aa434149.aspx).
        */
        wxExcelTextRange2 InsertSymbol(const wxString& fontName, long charNumber, MsoTriState unicode);

        //@{
        /**
        Gets the range of text specified by the index number from the TextRange2 object.

        [MSDN documentation for TextRange2.Item](http://msdn.microsoft.com/en-us/library/aa434150.aspx).
        */
        wxExcelTextRange2 Item(long index);
        wxExcelTextRange2 operator[](long index);
        //@}

        /**
        Returns a TextRange2 Represents the specified subset of left-to-right text runs. A text run consists of a range of characters that share the same font attributes.

        [MSDN documentation for TextRange2.LtrRun](http://msdn.microsoft.com/en-us/library/aa434151.aspx).
        */
        void LtrRun();

        /**
        Pastes the contents of the Clipboard into the TextRange2 object.

        [MSDN documentation for TextRange2.Paste](http://msdn.microsoft.com/en-us/library/aa434152.aspx).
        */
        wxExcelTextRange2 Paste();

        /**
        Replaces the text range with the contents of the Clipboard in the format specified. If the paste succeeds, this method returns a TextRange2 object including the text range that was pasted.

        [MSDN documentation for TextRange2.PasteSpecial](http://msdn.microsoft.com/en-us/library/aa434153.aspx).
        */
        wxExcelTextRange2 PasteSpecial(MsoClipboardFormat format);

        /**
        Removes all period (.) punctuation from the text in the TextRange2 object.

        [MSDN documentation for TextRange2.RemovePeriods](http://msdn.microsoft.com/en-us/library/aa434154.aspx).
        */
        void RemovePeriods();

        /**
        Finds specific text in a text range, replaces the found text with a specified string, and returns a TextRange2 Represents the first occurrence of the found text. Returns Nothing if no match is found.

        [MSDN documentation for TextRange2.Replace](http://msdn.microsoft.com/en-us/library/aa434155.aspx).
        */
        wxExcelTextRange2 Replace(const wxString& findWhat, const wxString& replaceWhat,
	                              long* after = NULL, MsoTriState* matchCase = NULL, MsoTriState* wholeWords = NULL);

        /**
        NOT IMPLEMENTED. wxAutomationObject does not support passing arguments by reference.
        
        Gets the coordinates of the vertices of the text bounding box for the specified text range. Read-only.

        [MSDN documentation for TextRange2.RotatedBounds](http://msdn.microsoft.com/en-us/library/aa434156.aspx).
        */
        void RotatedBounds(double& X1, double& Y1, double& X2, double& Y2,
                           double& X3, double& Y3, double& X4, double& Y4);

        /**
        Returns a TextRange2 Represents the specified subset of right-to-left text runs. A text run consists of a range of characters that share the same font attributes.

        [MSDN documentation for TextRange2.RtlRun](http://msdn.microsoft.com/en-us/library/aa434157.aspx).
        */
        wxExcelTextRange2 RtlRun();

        /**
        Selects the TextRange2 object.

        [MSDN documentation for TextRange2.Select](http://msdn.microsoft.com/en-us/library/aa434158.aspx).
        */
        void Select();

        /**
        Removes the white space on the left and right sides of the text in the TextRange2 object.

        [MSDN documentation for TextRange2.TrimText](http://msdn.microsoft.com/en-us/library/aa434159.aspx).
        */
        wxExcelTextRange2 TrimText();

        // ***** PROPERTIES *****

        /**
        Gets the height, in points, of the text bounding box for the specified text.

        [MSDN documentation for TextRange2.BoundHeight](http://msdn.microsoft.com/en-us/library/aa434427.aspx).
        */
        double GetBoundHeight();

        /**
        Gets the left coordinate, in points, of the text bounding box for the specified text.

        [MSDN documentation for TextRange2.BoundLeft](http://msdn.microsoft.com/en-us/library/aa434428.aspx).
        */
        double GetBoundLeft();

        /**
        Gets the top coordinate, in points, of the text bounding box for the specified text.

        [MSDN documentation for TextRange2.BoundTop](http://msdn.microsoft.com/en-us/library/aa434429.aspx).
        */
        double GetBoundTop();

        /**
        Gets the width, in points, of the text bounding box for the specified text.

        [MSDN documentation for TextRange2.BoundWidth](http://msdn.microsoft.com/en-us/library/aa434430.aspx).
        */
        double GetBoundWidth();

        /**        

        [MSDN documentation for TextRange2.Characters](http://msdn.microsoft.com/en-us/library/aa434431.aspx).
        */
        wxExcelTextRange2 GetCharacters(long start = 1, long* length = NULL);

        /**
        Gets a Long indicating the number of items in the TextRange2 collection.

        [MSDN documentation for TextRange2.Count](http://msdn.microsoft.com/en-us/library/aa434432.aspx).
        */
        long GetCount();

        /**
        Returns a Font Represents character formatting for the TextRange2 object.

        [MSDN documentation for TextRange2.Font](http://msdn.microsoft.com/en-us/library/aa434434.aspx).
        */
        wxExcelFont2 GetFont();

        /**
        Gets or sets the MsoLanguageID value of the TextRange2 object.

        [MSDN documentation for TextRange2.LanguageID](http://msdn.microsoft.com/en-us/library/aa434435.aspx).
        */
        MsoLanguageID GetLanguageID();

        /**
        Gets or sets the MsoLanguageID value of the TextRange2 object.

        [MSDN documentation for TextRange2.LanguageID](http://msdn.microsoft.com/en-us/library/aa434435.aspx).
        */
        void SetLanguageID(MsoLanguageID languageID);

        /**
        Get a Long that represents the length of a text range.

        [MSDN documentation for TextRange2.Length](http://msdn.microsoft.com/en-us/library/aa434436.aspx).
        */
        long GetLength();

        /**
        Returns a TextRange2 Represents the specified subset of text lines.

        [MSDN documentation for TextRange2.Lines](http://msdn.microsoft.com/en-us/library/aa434437.aspx).
        */
        wxExcelTextRange2 GetLines();

        /**        

        [MSDN documentation for TextRange2.MathZones](https://docs.microsoft.com/en-us/office/vba/api/office.textrange2.mathzones).
        */
        wxExcelTextRange2 GetMathZones(long start = 1, long* length = NULL);

        /**
        Returns a ParagraphFormat Represents paragraph formatting for the specified text.

        [MSDN documentation for TextRange2.ParagraphFormat](http://msdn.microsoft.com/en-us/library/aa434438.aspx).
        */
        wxExcelParagraphFormat2 GetParagraphFormat();

        /**
        Gets a TextRange2 Represents the specified subset of text paragraphs.

        [MSDN documentation for TextRange2.Paragraphs](http://msdn.microsoft.com/en-us/library/aa434439.aspx).
        */
        wxExcelTextRange2 GetParagraphs();


        /**
        Gets a TextRange2 Represents the specified subset of text runs. A text run consists of a range of characters that share the same font attributes.

        [MSDN documentation for TextRange2.Runs](http://msdn.microsoft.com/en-us/library/aa434451.aspx).
        */
        wxExcelTextRange2 GetRuns();

        /**
        Returns a TextRange2 Represents the specified subset of text sentences.

        [MSDN documentation for TextRange2.Sentences](http://msdn.microsoft.com/en-us/library/aa434452.aspx).
        */
        wxExcelTextRange2 GetSentences();

        /**
        Gets a Long value indicating the starting point of the specified text range.

        [MSDN documentation for TextRange2.Start](http://msdn.microsoft.com/en-us/library/aa434453.aspx).
        */
        long GetStart();

        /**
        Gets or sets a String value that represents the text in a text range.

        [MSDN documentation for TextRange2.Text](http://msdn.microsoft.com/en-us/library/aa434454.aspx).
        */
        wxString GetText();

        /**
        Gets or sets a String value that represents the text in a text range.

        [MSDN documentation for TextRange2.Text](http://msdn.microsoft.com/en-us/library/aa434454.aspx).
        */
        void SetText(const wxString& text);

        /**
        Gets a TextRange2 Represents the specified subset of text words.

        [MSDN documentation for TextRange2.Words](http://msdn.microsoft.com/en-us/library/aa434455.aspx).
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
