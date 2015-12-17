/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Modified by:
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////


#include "wx/wxAutoExcel_prec.h"

#include "wx/wxAutoExcelLineFormat.h"

#if WXAUTOEXCEL_USE_SHAPES || WXAUTOEXCEL_USE_CHARTS

#include "wx/wxAutoExcelColorFormat.h"

#include "wx/wxAutoExcel_private.h"

namespace wxAutoExcel {


// ***** class wxExcelLineFormat PROPERTIES *****

wxExcelColorFormat wxExcelLineFormat::GetBackColor()
{
    wxExcelColorFormat colorFormat;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("BackColor", colorFormat);
}

MsoArrowheadLength wxExcelLineFormat::GetBeginArrowheadLength()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("BeginArrowheadLength", MsoArrowheadLength, msoArrowheadShort);
}

void wxExcelLineFormat::SetBeginArrowheadLength(MsoArrowheadLength beginArrowheadLength)
{
    InvokePutProperty(wxS("BeginArrowheadLength"), (long)beginArrowheadLength);
}

MsoArrowheadStyle wxExcelLineFormat::GetBeginArrowheadStyle()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("BeginArrowheadStyle", MsoArrowheadStyle, msoArrowheadNone);
}

void wxExcelLineFormat::SetBeginArrowheadStyle(MsoArrowheadStyle beginArrowheadStyle)
{
    InvokePutProperty(wxS("BeginArrowheadStyle"), (long)beginArrowheadStyle);
}

MsoArrowheadWidth wxExcelLineFormat::GetBeginArrowheadWidth()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("BeginArrowheadWidth", MsoArrowheadWidth, msoArrowheadNarrow);
}

void wxExcelLineFormat::SetBeginArrowheadWidth(MsoArrowheadWidth beginArrowheadWidth)
{
    InvokePutProperty(wxS("BeginArrowheadWidth"), (long)beginArrowheadWidth);
}


MsoLineDashStyle wxExcelLineFormat::GetDashStyle()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("DashStyle", MsoLineDashStyle, msoLineSolid);
}

void wxExcelLineFormat::SetDashStyle(MsoLineDashStyle dashStyle)
{
    InvokePutProperty(wxS("DashStyle"), (long)dashStyle);
}

MsoArrowheadLength wxExcelLineFormat::GetEndArrowheadLength()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("EndArrowheadLength", MsoArrowheadLength, msoArrowheadShort);
}

void wxExcelLineFormat::SetEndArrowheadLength(MsoArrowheadLength endArrowheadLength)
{
    InvokePutProperty(wxS("EndArrowheadLength"), (long)endArrowheadLength);
}

MsoArrowheadStyle wxExcelLineFormat::GetEndArrowheadStyle()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("EndArrowheadStyle", MsoArrowheadStyle, msoArrowheadNone);
}

void wxExcelLineFormat::SetEndArrowheadStyle(MsoArrowheadStyle endArrowheadStyle)
{
    InvokePutProperty(wxS("EndArrowheadStyle"), (long)endArrowheadStyle);
}

MsoArrowheadWidth wxExcelLineFormat::GetEndArrowheadWidth()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("EndArrowheadWidth", MsoArrowheadWidth, msoArrowheadNarrow);
}

void wxExcelLineFormat::SetEndArrowheadWidth(MsoArrowheadWidth endArrowheadWidth)
{
    InvokePutProperty(wxS("EndArrowheadWidth"), (long)endArrowheadWidth);
}

wxExcelColorFormat wxExcelLineFormat::GetForeColor()
{
    wxExcelColorFormat colorFormat;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("ForeColor", colorFormat);
}

void wxExcelLineFormat::SetForeColor(const wxExcelColorFormat& foreColor)
{
    wxVariant vColor;
    if ( ObjectToVariant(&foreColor, vColor) )    
        InvokePutProperty(wxS("ForeColor"), vColor);
}

MsoTriState wxExcelLineFormat::GetInsetPen()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("InsetPen", MsoTriState, msoFalse);
}

void wxExcelLineFormat::SetInsetPen(MsoTriState insetPen)
{
    InvokePutProperty(wxS("InsetPen"), (long)insetPen);
}


MsoPatternType  wxExcelLineFormat::GetPattern()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("Pattern", MsoPatternType , msoPattern5Percent);
}

void wxExcelLineFormat::SetPattern(MsoPatternType  pattern)
{
    InvokePutProperty(wxS("Pattern"), (long)pattern);
}

MsoLineStyle  wxExcelLineFormat::GetStyle()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("Style", MsoLineStyle, msoLineSingle);
}

void wxExcelLineFormat::SetStyle(MsoLineStyle  style)
{
    InvokePutProperty(wxS("Style"), (long)style);
}

double wxExcelLineFormat::GetTransparency()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("Transparency");
}

void wxExcelLineFormat::SetTransparency(double transparency)
{
    InvokePutProperty(wxS("Transparency"), transparency);
}

MsoTriState wxExcelLineFormat::GetVisible()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("Visible", MsoTriState, msoFalse);
}

void wxExcelLineFormat::SetVisible(MsoTriState visible)
{
    InvokePutProperty(wxS("Visible"), (long)visible);
}

double wxExcelLineFormat::GetWeight()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("Weight");
}

void wxExcelLineFormat::SetWeight(double weight)
{
    InvokePutProperty(wxS("Weight"), weight);
}

} // namespace wxAutoExcel


#endif // #if WXAUTOEXCEL_USE_SHAPES || WXAUTOEXCEL_USE_CHARTS
