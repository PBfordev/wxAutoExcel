/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Modified by:
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////


#include "wx/wxAutoExcel_prec.h"

#include "wx/wxAutoExcelTextRange2.h"

#if WXAUTOEXCEL_USE_SHAPES || WXAUTOEXCEL_USE_CHARTS

#include "wx/wxAutoExcelFont2.h"
#include "wx/wxAutoExcelParagraphFormat2.h"

#include "wx/wxAutoExcel_private.h"

namespace wxAutoExcel {

    // ***** class wxExcelTextRange2 METHODS *****

void wxExcelTextRange2::AddPeriods()
{
    WXAUTOEXCEL_CALL_METHOD0_RET("AddPeriods", "null");
}

void wxExcelTextRange2::ChangeCase(MsoTextChangeCase type)
{
    WXAUTOEXCEL_CALL_METHOD1_RET("ChangeCase", (long)type, "null");
}

void wxExcelTextRange2::Copy()
{
    WXAUTOEXCEL_CALL_METHOD0_RET("Copy", "null");
}

void wxExcelTextRange2::Cut()
{
    WXAUTOEXCEL_CALL_METHOD0_RET("Cut", "null");
}

void wxExcelTextRange2::Delete()
{
    WXAUTOEXCEL_CALL_METHOD0_RET("Delete", "null");
}

wxExcelTextRange2 wxExcelTextRange2::Find(const wxString& findWhat, long* after, MsoTriState* matchCase, MsoTriState* wholeWords)
{
    wxASSERT( !findWhat.empty() );

    WXAUTOEXCEL_OPTIONALCPPSTR_TO_OPTIONALVARIANT_NAME(FindWhat, findWhat);
    WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT_NAME(After, after);
    WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT_NAME(MatchCase, ((long*)matchCase));
    WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT_NAME(WholeWords, ((long*)wholeWords));

    wxExcelTextRange2 range;

    WXAUTOEXCEL_CALL_METHOD4("Find", vFindWhat, vAfter, vMatchCase, vWholeWords, "void*", range);
    VariantToObject(vResult, &range);
    return range;
}

wxExcelTextRange2 wxExcelTextRange2::InsertAfter(const wxString& newText)
{
    wxExcelTextRange2 range;
    WXAUTOEXCEL_CALL_METHOD1_OBJECT("InsertAfter", newText, range);
}

wxExcelTextRange2 wxExcelTextRange2::InsertBefore(const wxString& newText)
{
    wxExcelTextRange2 range;
    WXAUTOEXCEL_CALL_METHOD1_OBJECT("InsertBefore", newText, range);
}

wxExcelTextRange2 wxExcelTextRange2::InsertSymbol(const wxString& fontName, long charNumber, MsoTriState unicode)
{
    wxExcelTextRange2 range;

    WXAUTOEXCEL_CALL_METHOD3("InsertSymbol", fontName, charNumber, (long)unicode, "void*", range);
    VariantToObject(vResult, &range);
    return range;

}

wxExcelTextRange2 wxExcelTextRange2::Item(long index)
{
    wxASSERT( index > 0 );

    wxExcelTextRange2 range;
    WXAUTOEXCEL_CALL_METHOD1_OBJECT("Item", index, range);
}

wxExcelTextRange2 wxExcelTextRange2::operator[](long index)
{
    return Item(index);
}

void wxExcelTextRange2::LtrRun()
{
    WXAUTOEXCEL_CALL_METHOD0_RET("LtrRun", "null");
}

wxExcelTextRange2 wxExcelTextRange2::Paste()
{
    wxExcelTextRange2 range;
    WXAUTOEXCEL_CALL_METHOD0_OBJECT("Paste", range);
}

wxExcelTextRange2 wxExcelTextRange2::PasteSpecial(MsoClipboardFormat format)
{
    wxExcelTextRange2 range;
    WXAUTOEXCEL_CALL_METHOD1_OBJECT("PasteSpecial", (long)format, range);
}

void wxExcelTextRange2::RemovePeriods()
{
    WXAUTOEXCEL_CALL_METHOD0_RET("RemovePeriods", "null");
}

wxExcelTextRange2 wxExcelTextRange2::Replace(const wxString& findWhat, const wxString& replaceWhat,
	                                         long* after, MsoTriState* matchCase, MsoTriState* wholeWords)
{
    wxASSERT( !findWhat.empty() );

    WXAUTOEXCEL_OPTIONALCPPSTR_TO_OPTIONALVARIANT_NAME(FindWhat, findWhat);
    WXAUTOEXCEL_OPTIONALCPPSTR_TO_OPTIONALVARIANT_NAME(ReplaceWhat, replaceWhat);
    WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT_NAME(After, after);
    WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT_NAME(MatchCase, ((long*)matchCase));
    WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT_NAME(WholeWords, ((long*)wholeWords));

    wxExcelTextRange2 range;

    WXAUTOEXCEL_CALL_METHOD5("Replace", vFindWhat, vReplaceWhat, vAfter, vMatchCase, vWholeWords, "void*", range);
    VariantToObject(vResult, &range);
    return range;
}

void wxExcelTextRange2::RotatedBounds(double& X1, double& Y1, double& X2, double& Y2,
                                      double& X3, double& Y3, double& X4, double& Y4)
{
    // wxWidgets Automation interface doesn't support passing parameters by reference
    wxFAIL_MSG(wxS("Not implemented"));

    wxUnusedVar(X1); wxUnusedVar(Y1);
    wxUnusedVar(X2); wxUnusedVar(Y2);
    wxUnusedVar(X3); wxUnusedVar(Y3);
    wxUnusedVar(X4); wxUnusedVar(Y4);
}

void wxExcelTextRange2::RtlRun()
{
    WXAUTOEXCEL_CALL_METHOD0_RET("RtlRun", "null");
}

void wxExcelTextRange2::Select()
{
    WXAUTOEXCEL_CALL_METHOD0_RET("Select", "null");
}

wxExcelTextRange2 wxExcelTextRange2::TrimText()
{
    wxExcelTextRange2 range;
    WXAUTOEXCEL_CALL_METHOD0_OBJECT("TrimText", range);
}


// ***** class wxExcelTextRange2 PROPERTIES *****


double wxExcelTextRange2::GetBoundHeight()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("BoundHeight");
}

double wxExcelTextRange2::GetBoundLeft()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("BoundLeft");
}

double wxExcelTextRange2::GetBoundTop()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("BoundTop");
}

double wxExcelTextRange2::GetBoundWidth()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("BoundWidth");
}

wxExcelTextRange2 wxExcelTextRange2::GetCharacters()
{
    wxExcelTextRange2 characters;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("Characters", characters);
}

long wxExcelTextRange2::GetCount()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("Count");
}

wxExcelFont2 wxExcelTextRange2::GetFont()
{
    wxExcelFont2 font;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("Font", font);
}

MsoLanguageID wxExcelTextRange2::GetLanguageID()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("LanguageID", MsoLanguageID, msoLanguageIDNone);
}

void wxExcelTextRange2::SetLanguageID(MsoLanguageID languageID)
{
    InvokePutProperty(wxS("LanguageID"), (long)languageID);
}

long wxExcelTextRange2::GetLength()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("Length");
}

wxExcelTextRange2 wxExcelTextRange2::GetLines()
{
    wxExcelTextRange2 textRange2;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("Lines", textRange2);
}

wxExcelParagraphFormat2 wxExcelTextRange2::GetParagraphFormat()
{
    wxExcelParagraphFormat2 paragraphFormat2;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("ParagraphFormat", paragraphFormat2);
}

wxExcelTextRange2 wxExcelTextRange2::GetParagraphs()
{
    wxExcelTextRange2 textRange2;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("Paragraphs", textRange2);
}

wxExcelTextRange2 wxExcelTextRange2::GetRuns()
{
    wxExcelTextRange2 textRange2;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("Runs", textRange2);
}

wxExcelTextRange2 wxExcelTextRange2::GetSentences()
{
    wxExcelTextRange2 textRange2;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("Sentences", textRange2);
}

long wxExcelTextRange2::GetStart()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("Start");
}

wxString wxExcelTextRange2::GetText()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("Text");
}

void wxExcelTextRange2::SetText(const wxString& text)
{
    InvokePutProperty(wxS("Text"), text);
}

wxExcelTextRange2 wxExcelTextRange2::GetWords()
{
    wxExcelTextRange2 textRange2;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("Words", textRange2);
}

} // namespace wxAutoExcel

#endif // #if WXAUTOEXCEL_USE_SHAPES || WXAUTOEXCEL_USE_CHARTS
