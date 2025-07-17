/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////


#include "wx/wxAutoExcel_prec.h"

#include "wx/wxAutoExcelTextFrame2.h"

#if WXAUTOEXCEL_USE_SHAPES || WXAUTOEXCEL_USE_CHARTS

#include "wx/wxAutoExcelApplication.h"
#include "wx/wxAutoExcelTextColumn2.h"
#include "wx/wxAutoExcelTextRange2.h"
#include "wx/wxAutoExcelThreeDFormat.h"

#include "wx/wxAutoExcel_private.h"

namespace wxAutoExcel {

// ***** class wxExcelTextFrame2 METHODS *****

void wxExcelTextFrame2::DeleteText()
{
    WXAUTOEXCEL_CALL_METHOD0_RET("DeleteText", "null");
}

// ***** class wxExcelTextFrame2 PROPERTIES *****

MsoAutoSize wxExcelTextFrame2::GetAutoSize()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("AutoSize", MsoAutoSize, msoAutoSizeNone);
}

wxExcelTextColumn2 wxExcelTextFrame2::GetColumn()
{
    wxExcelTextColumn2 column;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("Column", column);
}


MsoTriState wxExcelTextFrame2::GetHasText()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("HasText", MsoTriState, msoFalse);
}

MsoHorizontalAnchor wxExcelTextFrame2::GetHorizontalAnchor()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("HorizontalAnchor", MsoHorizontalAnchor, msoAnchorNone);
}

void wxExcelTextFrame2::SetHorizontalAnchor(MsoHorizontalAnchor horizontalAnchor)
{
    InvokePutProperty(wxS("HorizontalAnchor"), (long)horizontalAnchor);
}

double wxExcelTextFrame2::GetMarginBottom()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("MarginBottom");
}

void wxExcelTextFrame2::SetMarginBottom(double marginBottom)
{
    InvokePutProperty(wxS("MarginBottom"), marginBottom);
}

double wxExcelTextFrame2::GetMarginLeft()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("MarginLeft");
}

void wxExcelTextFrame2::SetMarginLeft(double marginLeft)
{
    InvokePutProperty(wxS("MarginLeft"), marginLeft);
}

double wxExcelTextFrame2::GetMarginRight()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("MarginRight");
}

void wxExcelTextFrame2::SetMarginRight(double marginRight)
{
    InvokePutProperty(wxS("MarginRight"), marginRight);
}

double wxExcelTextFrame2::GetMarginTop()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("MarginTop");
}

void wxExcelTextFrame2::SetMarginTop(double marginTop)
{
    InvokePutProperty(wxS("MarginTop"), marginTop);
}

MsoTriState wxExcelTextFrame2::GetNoTextRotation()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("NoTextRotation", MsoTriState, msoFalse);
}

void wxExcelTextFrame2::SetNoTextRotation(MsoTriState rotation)
{
    InvokePutProperty(wxS("NoTextRotation"), (long)rotation);
}


MsoTextOrientation wxExcelTextFrame2::GetOrientation()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("Orientation", MsoTextOrientation, msoTextOrientationHorizontal);
}

void wxExcelTextFrame2::SetOrientation(MsoTextOrientation orientation)
{
    InvokePutProperty(wxS("Orientation"), (long)orientation);
}


MsoPathFormat wxExcelTextFrame2::GetPathFormat()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("PathFormat", MsoPathFormat, msoPathTypeNone);
}

void wxExcelTextFrame2::SetPathFormat(MsoPathFormat pathFormat)
{
    InvokePutProperty(wxS("PathFormat"), (long)pathFormat);
}


wxExcelTextRange2 wxExcelTextFrame2::GetTextRange()
{
    wxExcelTextRange2 range;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("TextRange", range);
}

wxExcelThreeDFormat wxExcelTextFrame2::GetThreeD()
{
    wxExcelThreeDFormat threeD;
    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("ThreeD", threeD);
}

MsoVerticalAnchor wxExcelTextFrame2::GetVerticalAnchor()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("VerticalAnchor", MsoVerticalAnchor, msoAnchorTop);
}

void wxExcelTextFrame2::SetVerticalAnchor(MsoVerticalAnchor verticalAnchor)
{
    InvokePutProperty(wxS("VerticalAnchor"), verticalAnchor);
}

MsoWarpFormat wxExcelTextFrame2::GetWarpFormat()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("WarpFormat", MsoWarpFormat, msoWarpFormat1);
}

void wxExcelTextFrame2::SetWarpFormat(MsoWarpFormat warpFormat)
{
    InvokePutProperty(wxS("WarpFormat"), (long)warpFormat);
}

MsoPresetTextEffect wxExcelTextFrame2::GetWordArtformat()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("WordArtformat", MsoPresetTextEffect, msoTextEffect1);
}

void wxExcelTextFrame2::SetWordArtformat(MsoPresetTextEffect wordArtformat)
{
    InvokePutProperty(wxS("WordArtformat"), (long)wordArtformat);
}

MsoTriState wxExcelTextFrame2::GetWordWrap()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("WordWrap", MsoTriState, msoFalse);
}

void wxExcelTextFrame2::SetWordWrap(MsoTriState wordWrap)
{
    InvokePutProperty(wxS("WordWrap"), (long)wordWrap);
}

} // namespace wxAutoExcel

#endif // #if WXAUTOEXCEL_USE_SHAPES || WXAUTOEXCEL_USE_CHARTS
