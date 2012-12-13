/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Modified by:
// RCS-ID:      $Id: $
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////


#include "wx/wxAutoExcel_prec.h"

#include "wx/wxAutoExcelPoints.h"

#if WXAUTOEXCEL_USE_CHARTS

#include "wx/wxAutoExcelChartFormat.h"
#include "wx/wxAutoExcelDataLabels.h"

#include "wx/wxAutoExcelPrivate.h"

namespace wxAutoExcel {

// ***** class wxExcelPoint METHODS *****

void wxExcelPoint::ApplyDataLabels(XlDataLabelsType* type, wxXlTribool legendKey,
                                    wxXlTribool autoText, wxXlTribool hasLeaderLines,
                                    wxXlTribool showSeriesName, wxXlTribool showCategoryName,
                                    wxXlTribool showValue, wxXlTribool showPercentage,
                                    wxXlTribool showBubbleSize, const wxString& separator)
{
    wxVariantVector args;

    WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT_NAME_VECTOR(Type, ((long*)type), args);
    WXAUTOEXCEL_OPTIONALCPPTBOOL_TO_OPTIONALVARIANT_NAME_VECTOR(LegendKey, legendKey, args);
    WXAUTOEXCEL_OPTIONALCPPTBOOL_TO_OPTIONALVARIANT_NAME_VECTOR(AutoText, autoText, args);
    WXAUTOEXCEL_OPTIONALCPPTBOOL_TO_OPTIONALVARIANT_NAME_VECTOR(HasLeaderLines, hasLeaderLines, args);
    WXAUTOEXCEL_OPTIONALCPPTBOOL_TO_OPTIONALVARIANT_NAME_VECTOR(ShowSeriesName, showSeriesName, args);
    WXAUTOEXCEL_OPTIONALCPPTBOOL_TO_OPTIONALVARIANT_NAME_VECTOR(ShowCategoryName, showCategoryName, args);
    WXAUTOEXCEL_OPTIONALCPPTBOOL_TO_OPTIONALVARIANT_NAME_VECTOR(ShowValue, showValue, args);
    WXAUTOEXCEL_OPTIONALCPPTBOOL_TO_OPTIONALVARIANT_NAME_VECTOR(ShowPercentage, showPercentage, args);
    WXAUTOEXCEL_OPTIONALCPPTBOOL_TO_OPTIONALVARIANT_NAME_VECTOR(ShowBubbleSize, showBubbleSize, args);
    WXAUTOEXCEL_OPTIONALCPPSTR_TO_OPTIONALVARIANT_NAME_VECTOR(Separator, separator, args);

    WXAUTOEXCEL_CALL_METHODARR_RET("ApplyDataLabels", args, "null");
}

bool wxExcelPoint::ClearFormats()
{
    WXAUTOEXCEL_CALL_METHOD0_BOOL("ClearFormats");
}

bool wxExcelPoint::Copy()
{
    WXAUTOEXCEL_CALL_METHOD0_BOOL("Copy");
}


bool wxExcelPoint::Delete()
{
    WXAUTOEXCEL_CALL_METHOD0_BOOL("Delete");
}

bool wxExcelPoint::Paste()
{
    WXAUTOEXCEL_CALL_METHOD0_BOOL("Paste");
}

bool wxExcelPoint::Select()
{
    WXAUTOEXCEL_CALL_METHOD0_BOOL("Select");
}

// ***** class wxExcelPoint PROPERTIES *****

bool wxExcelPoint::GetApplyPictToEnd()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("ApplyPictToEnd");
}

void wxExcelPoint::SetApplyPictToEnd(bool applyPictToEnd)
{
    InvokePutProperty(wxS("ApplyPictToEnd"), applyPictToEnd);
}

bool wxExcelPoint::GetApplyPictToFront()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("ApplyPictToFront");
}

void wxExcelPoint::SetApplyPictToFront(bool applyPictToFront)
{
    InvokePutProperty(wxS("ApplyPictToFront"), applyPictToFront);
}

bool wxExcelPoint::GetApplyPictToSides()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("ApplyPictToSides");
}

void wxExcelPoint::SetApplyPictToSides(bool applyPictToSides)
{
    InvokePutProperty(wxS("ApplyPictToSides"), applyPictToSides);
}

wxExcelDataLabel wxExcelPoint::GetDataLabel()
{
    wxExcelDataLabel dataLabel;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("DataLabel", dataLabel);
}

long wxExcelPoint::GetExplosion()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("Explosion");
}

void wxExcelPoint::SetExplosion(long explosion)
{
    InvokePutProperty(wxS("Explosion"), explosion);
}

wxExcelChartFormat wxExcelPoint::GetFormat()
{
    wxExcelChartFormat chartFormat;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("Format", chartFormat);
}

bool wxExcelPoint::GetHas3DEffect()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("Has3DEffect");
}

void wxExcelPoint::SetHas3DEffect(bool has3DEffect)
{
    InvokePutProperty(wxS("Has3DEffect"), has3DEffect);
}

bool wxExcelPoint::GetHasDataLabel()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("HasDataLabel");
}

void wxExcelPoint::SetHasDataLabel(bool hasDataLabel)
{
    InvokePutProperty(wxS("HasDataLabel"), hasDataLabel);
}

bool wxExcelPoint::GetInvertIfNegative()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("InvertIfNegative");
}

void wxExcelPoint::SetInvertIfNegative(bool invertIfNegative)
{
    InvokePutProperty(wxS("InvertIfNegative"), invertIfNegative);
}

wxColour wxExcelPoint::GetMarkerBackgroundColor()
{
    WXAUTOEXCEL_PROPERTY_COLOR_GET0("MarkerBackgroundColor");
}

void wxExcelPoint::SetMarkerBackgroundColor(const wxColour& markerBackgroundColor)
{
    InvokePutProperty(wxS("MarkerBackgroundColor"), (long)markerBackgroundColor.GetRGB());
}

long wxExcelPoint::GetMarkerBackgroundColorIndex()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("MarkerBackgroundColorIndex");
}

void wxExcelPoint::SetMarkerBackgroundColorIndex(long markerBackgroundColorIndex)
{
    InvokePutProperty(wxS("MarkerBackgroundColorIndex"), markerBackgroundColorIndex);
}

wxColour wxExcelPoint::GetMarkerForegroundColor()
{
    WXAUTOEXCEL_PROPERTY_COLOR_GET0("MarkerForegroundColor");
}

void wxExcelPoint::SetMarkerForegroundColor(const wxColour& markerForegroundColor)
{
    InvokePutProperty(wxS("MarkerForegroundColor"), (long)markerForegroundColor.GetRGB());
}

long wxExcelPoint::GetMarkerForegroundColorIndex()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("MarkerForegroundColorIndex");
}

void wxExcelPoint::SetMarkerForegroundColorIndex(long markerForegroundColorIndex)
{
    InvokePutProperty(wxS("MarkerForegroundColorIndex"), markerForegroundColorIndex);
}

long wxExcelPoint::GetMarkerSize()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("MarkerSize");
}

void wxExcelPoint::SetMarkerSize(long markerSize)
{
    InvokePutProperty(wxS("MarkerSize"), markerSize);
}

XlMarkerStyle wxExcelPoint::GetMarkerStyle()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("MarkerStyle", XlMarkerStyle, xlMarkerStyleSquare);
}

void wxExcelPoint::SetMarkerStyle(XlMarkerStyle markerStyle)
{
    InvokePutProperty(wxS("MarkerStyle"), (long)markerStyle);
}


XlChartPictureType wxExcelPoint::GetPictureType()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("PictureType", XlChartPictureType, xlStretch);
}

void wxExcelPoint::SetPictureType(XlChartPictureType pictureType)
{
    InvokePutProperty(wxS("PictureType"), (long)pictureType);
}

long wxExcelPoint::GetPictureUnit()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("PictureUnit");
}

void wxExcelPoint::SetPictureUnit(long pictureUnit)
{
    InvokePutProperty(wxS("PictureUnit"), pictureUnit);
}

double wxExcelPoint::GetPictureUnit2()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("PictureUnit2");
}

void wxExcelPoint::SetPictureUnit2(double pictureUnit2)
{
    InvokePutProperty(wxS("PictureUnit2"), pictureUnit2);
}

bool wxExcelPoint::GetSecondaryPlot()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("SecondaryPlot");
}

void wxExcelPoint::SetSecondaryPlot(bool secondaryPlot)
{
    InvokePutProperty(wxS("SecondaryPlot"), secondaryPlot);
}

bool wxExcelPoint::GetShadow()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("Shadow");
}

void wxExcelPoint::SetShadow(bool shadow)
{
    InvokePutProperty(wxS("Shadow"), shadow);
}

// ***** class wxExcelPoints METHODS *****

wxExcelPoint wxExcelPoints::Item(long index)
{
    wxASSERT( index > 0 );
    
    wxExcelPoint point;
    WXAUTOEXCEL_CALL_METHOD1_OBJECT("Item", index, point);
}

wxExcelPoint wxExcelPoints::operator[](long index)
{
    return Item(index);
}

// ***** class wxExcelPoints PROPERTIES *****


long wxExcelPoints::GetCount()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("Count");
}

} // namespace wxAutoExcel


#endif // #if WXAUTOEXCEL_USE_CHARTS
