/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Modified by:
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////


#include "wx/wxAutoExcel_prec.h"

#include "wx/wxAutoExcelParagraphFormat2.h"

#if WXAUTOEXCEL_USE_SHAPES || WXAUTOEXCEL_USE_CHARTS

#include "wx/wxAutoExcelBulletFormat2.h"
#include "wx/wxAutoExcelTabStops2.h"

#include "wx/wxAutoExcel_private.h"

namespace wxAutoExcel {

// ***** class wxExcelParagraphFormat2 PROPERTIES *****

MsoParagraphAlignment wxExcelParagraphFormat2::GetAlignment()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("Alignment", MsoParagraphAlignment, msoAlignLeft);
}

MsoBaselineAlignment wxExcelParagraphFormat2::GetBaselineAlignment()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("BaselineAlignment", MsoBaselineAlignment, msoBaselineAlignBaseline);
}

void wxExcelParagraphFormat2::SetBaselineAlignment(MsoBaselineAlignment baselineAlignment)
{
    InvokePutProperty(wxS("BaselineAlignment"), (long)baselineAlignment);
}

wxExcelBulletFormat2 wxExcelParagraphFormat2::GetBullet()
{
    wxExcelBulletFormat2 bulletFormat2;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("Bullet", bulletFormat2);
}

MsoTriState wxExcelParagraphFormat2::GetFarEastLineBreakLevel()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("FarEastLineBreakLevel", MsoTriState, msoFalse);
}

void wxExcelParagraphFormat2::SetFarEastLineBreakLevel(MsoTriState farEastLineBreakLevel)
{
    InvokePutProperty(wxS("FarEastLineBreakLevel"), (long)farEastLineBreakLevel);
}

double wxExcelParagraphFormat2::GetFirstLineIndent()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("FirstLineIndent");
}

void wxExcelParagraphFormat2::SetFirstLineIndent(double firstLineIndent)
{
    InvokePutProperty(wxS("FirstLineIndent"), firstLineIndent);
}

MsoTriState wxExcelParagraphFormat2::GetHangingPunctuation()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("HangingPunctuation", MsoTriState, msoFalse);
}

void wxExcelParagraphFormat2::SetHangingPunctuation(MsoTriState hangingPunctuation)
{
    InvokePutProperty(wxS("HangingPunctuation"), (long)hangingPunctuation);
}

long wxExcelParagraphFormat2::GetIndentLevel()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("IndentLevel");
}

void wxExcelParagraphFormat2::SetIndentLevel(long indentLevel)
{
    InvokePutProperty(wxS("IndentLevel"), indentLevel);
}

double wxExcelParagraphFormat2::GetLeftIndent()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("LeftIndent");
}

void wxExcelParagraphFormat2::SetLeftIndent(double leftIndent)
{
    InvokePutProperty(wxS("LeftIndent"), leftIndent);
}

MsoTriState wxExcelParagraphFormat2::GetLineRuleAfter()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("LineRuleAfter", MsoTriState, msoFalse);
}

void wxExcelParagraphFormat2::SetLineRuleAfter(MsoTriState lineRuleAfter)
{
    InvokePutProperty(wxS("LineRuleAfter"), (long)lineRuleAfter);
}

MsoTriState wxExcelParagraphFormat2::GetLineRuleBefore()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("LineRuleBefore", MsoTriState, msoFalse);
}

void wxExcelParagraphFormat2::SetLineRuleBefore(MsoTriState lineRuleBefore)
{
    InvokePutProperty(wxS("LineRuleBefore"), (long)lineRuleBefore);
}

MsoTriState wxExcelParagraphFormat2::GetLineRuleWithin()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("LineRuleWithin", MsoTriState, msoFalse);
}

void wxExcelParagraphFormat2::SetLineRuleWithin(MsoTriState lineRuleWithin)
{
    InvokePutProperty(wxS("LineRuleWithin"), (long)lineRuleWithin);
}

double wxExcelParagraphFormat2::GetRightIndent()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("RightIndent");
}

void wxExcelParagraphFormat2::SetRightIndent(double rightIndent)
{
    InvokePutProperty(wxS("RightIndent"), rightIndent);
}

double wxExcelParagraphFormat2::GetSpaceAfter()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("SpaceAfter");
}

void wxExcelParagraphFormat2::SetSpaceAfter(double spaceAfter)
{
    InvokePutProperty(wxS("SpaceAfter"), spaceAfter);
}

double wxExcelParagraphFormat2::GetSpaceBefore()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("SpaceBefore");
}

void wxExcelParagraphFormat2::SetSpaceBefore(double spaceBefore)
{
    InvokePutProperty(wxS("SpaceBefore"), spaceBefore);
}

double wxExcelParagraphFormat2::GetSpaceWithin()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("SpaceWithin");
}

void wxExcelParagraphFormat2::SetSpaceWithin(double spaceWithin)
{
    InvokePutProperty(wxS("SpaceWithin"), spaceWithin);
}

wxExcelTabStops2 wxExcelParagraphFormat2::GetTabStops()
{
    wxExcelTabStops2 tabStops2;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("TabStops", tabStops2);
}

MsoTextDirection wxExcelParagraphFormat2::GetTextDirection()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("TextDirection", MsoTextDirection, msoTextDirectionLeftToRight);
}

void wxExcelParagraphFormat2::SetTextDirection(MsoTextDirection textDirection)
{
    InvokePutProperty(wxS("TextDirection"), (long)textDirection);
}

MsoTriState wxExcelParagraphFormat2::GetWordWrap()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("WordWrap", MsoTriState, msoFalse);
}

void wxExcelParagraphFormat2::SetWordWrap(MsoTriState wordWrap)
{
    InvokePutProperty(wxS("WordWrap"), (long)wordWrap);
}

} // namespace wxAutoExcel

#endif // #if WXAUTOEXCEL_USE_SHAPES || WXAUTOEXCEL_USE_CHARTS
