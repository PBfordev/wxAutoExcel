/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Modified by:
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////


#include "wx/wxAutoExcel_prec.h"

#include "wx/wxAutoExcelSeries.h"

#if WXAUTOEXCEL_USE_CHARTS

#include "wx/wxAutoExcelChartFormat.h"
#include "wx/wxAutoExcelDataLabels.h"
#include "wx/wxAutoExcelErrorBars.h"
#include "wx/wxAutoExcelLeaderLines.h"
#include "wx/wxAutoExcelPoints.h"
#include "wx/wxAutoExcelRange.h"
#include "wx/wxAutoExcelTrendlines.h"

#include "wx/wxAutoExcel_private.h"

namespace wxAutoExcel {

// ***** class wxExcelSeries METHODS *****

void wxExcelSeries::ApplyDataLabels(XlDataLabelsType* type, wxXlTribool legendKey,
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

bool wxExcelSeries::ClearFormats()
{
    WXAUTOEXCEL_CALL_METHOD0_BOOL("ClearFormats");
}

bool wxExcelSeries::Copy()
{
    WXAUTOEXCEL_CALL_METHOD0_BOOL("Copy");
}

wxExcelDataLabels wxExcelSeries::DataLabels()
{
    wxExcelDataLabels labels;
    WXAUTOEXCEL_CALL_METHOD0_OBJECT("DataLabels", labels);
}

bool wxExcelSeries::Delete()
{
    WXAUTOEXCEL_CALL_METHOD0_BOOL("Delete");
}

//@FIXME check if amount and minusValues are scalars or not!
bool wxExcelSeries::ErrorBar(XlErrorBarDirection direction, XlErrorBarInclude include, XlErrorBarType type,
                             double* amount, double* minusValues)
{
    wxVariant vDirection((long)direction, wxS("Direction"));
    wxVariant vInclude((long)include, wxS("Include"));
    wxVariant vType((long)type, wxS("Type"));
    WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT_NAME(Amount, amount);
    WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT_NAME(MinusValues, minusValues);
    
    WXAUTOEXCEL_CALL_METHOD5("ErrorBar", vDirection, vInclude, vType, vAmount, vMinusValues, "bool", false);
    
    //@FIXME check if it returns bool or ErrorBar object!
    return vResult.GetBool();
}

bool wxExcelSeries::Paste()
{
    WXAUTOEXCEL_CALL_METHOD0_BOOL("Paste");
}

wxExcelPoints wxExcelSeries::Points()
{
    wxExcelPoints points;
    WXAUTOEXCEL_CALL_METHOD0_OBJECT("Points", points);
}

wxExcelPoint wxExcelSeries::Points(long index)
{
    wxExcelPoint point;
    WXAUTOEXCEL_CALL_METHOD1_OBJECT("Points", index, point);
}

bool wxExcelSeries::Select()
{
    WXAUTOEXCEL_CALL_METHOD0_BOOL("Select");
}

wxExcelTrendlines wxExcelSeries::Trendlines()
{
    wxExcelTrendlines lines;     
    WXAUTOEXCEL_CALL_METHOD0_OBJECT("Trendlines", lines);
}

// ***** class wxExcelSeries PROPERTIES *****


bool wxExcelSeries::GetApplyPictToEnd()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("ApplyPictToEnd");
}

void wxExcelSeries::SetApplyPictToEnd(bool applyPictToEnd)
{
    InvokePutProperty(wxS("ApplyPictToEnd"), applyPictToEnd);
}

bool wxExcelSeries::GetApplyPictToFront()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("ApplyPictToFront");
}

void wxExcelSeries::SetApplyPictToFront(bool applyPictToFront)
{
    InvokePutProperty(wxS("ApplyPictToFront"), applyPictToFront);
}

bool wxExcelSeries::GetApplyPictToSides()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("ApplyPictToSides");
}

void wxExcelSeries::SetApplyPictToSides(bool applyPictToSides)
{
    InvokePutProperty(wxS("ApplyPictToSides"), applyPictToSides);
}

XlAxisGroup wxExcelSeries::GetAxisGroup()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("AxisGroup", XlAxisGroup, xlPrimary);
}

XlBarShape wxExcelSeries::GetBarShape()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("BarShape", XlBarShape, xlBox);
}

void wxExcelSeries::SetBarShape(XlBarShape barShape)
{
    InvokePutProperty(wxS("BarShape"), (long)barShape);
}

wxString wxExcelSeries::GetBubbleSizes()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("BubbleSizes");
}

void wxExcelSeries::SetBubbleSizes(const wxString& bubbleSizes)
{
    InvokePutProperty(wxS("BubbleSizes"), bubbleSizes);
}

XlChartType wxExcelSeries::GetChartType()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("ChartType", XlChartType, xlArea);
}

void wxExcelSeries::SetChartType(XlChartType chartType)
{
    InvokePutProperty(wxS("ChartType"), (long)chartType);
}

wxExcelErrorBars wxExcelSeries::GetErrorBars()
{
    wxExcelErrorBars errorBars;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("ErrorBars", errorBars);
}

long wxExcelSeries::GetExplosion()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("Explosion");
}

void wxExcelSeries::SetExplosion(long explosion)
{
    InvokePutProperty(wxS("Explosion"), explosion);
}

wxExcelChartFormat wxExcelSeries::GetFormat()
{
    wxExcelChartFormat chartFormat;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("Format", chartFormat);
}

wxString wxExcelSeries::GetFormula()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("Formula");
}

void wxExcelSeries::SetFormula(const wxString& formula)
{
    InvokePutProperty(wxS("Formula"), formula);
}

wxString wxExcelSeries::GetFormulaLocal()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("FormulaLocal");
}

void wxExcelSeries::SetFormulaLocal(const wxString& formulaLocal)
{
    InvokePutProperty(wxS("FormulaLocal"), formulaLocal);
}

wxString wxExcelSeries::GetFormulaR1C1()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("FormulaR1C1");
}

void wxExcelSeries::SetFormulaR1C1(const wxString& formulaR1C1)
{
    InvokePutProperty(wxS("FormulaR1C1"), formulaR1C1);
}

wxString wxExcelSeries::GetFormulaR1C1Local()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("FormulaR1C1Local");
}

void wxExcelSeries::SetFormulaR1C1Local(const wxString& formulaR1C1Local)
{
    InvokePutProperty(wxS("FormulaR1C1Local"), formulaR1C1Local);
}

bool wxExcelSeries::GetHas3DEffect()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("Has3DEffect");
}

void wxExcelSeries::SetHas3DEffect(bool has3DEffect)
{
    InvokePutProperty(wxS("Has3DEffect"), has3DEffect);
}

bool wxExcelSeries::GetHasDataLabels()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("HasDataLabels");
}

void wxExcelSeries::SetHasDataLabels(bool hasDataLabels)
{
    InvokePutProperty(wxS("HasDataLabels"), hasDataLabels);
}

bool wxExcelSeries::GetHasErrorBars()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("HasErrorBars");
}

void wxExcelSeries::SetHasErrorBars(bool hasErrorBars)
{
    InvokePutProperty(wxS("HasErrorBars"), hasErrorBars);
}

bool wxExcelSeries::GetHasLeaderLines()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("HasLeaderLines");
}

void wxExcelSeries::SetHasLeaderLines(bool hasLeaderLines)
{
    InvokePutProperty(wxS("HasLeaderLines"), hasLeaderLines);
}

bool wxExcelSeries::GetInvertIfNegative()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("InvertIfNegative");
}

void wxExcelSeries::SetInvertIfNegative(bool invertIfNegative)
{
    InvokePutProperty(wxS("InvertIfNegative"), invertIfNegative);
}

wxExcelLeaderLines wxExcelSeries::GetLeaderLines()
{
    wxExcelLeaderLines leaderLines;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("LeaderLines", leaderLines);
}

wxColour wxExcelSeries::GetMarkerBackgroundColor()
{
    WXAUTOEXCEL_PROPERTY_COLOR_GET0("MarkerBackgroundColor");
}

void wxExcelSeries::SetMarkerBackgroundColor(const wxColour& markerBackgroundColor)
{
    InvokePutProperty(wxS("MarkerBackgroundColor"), (long)markerBackgroundColor.GetRGB());
}

long wxExcelSeries::GetMarkerBackgroundColorIndex()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("MarkerBackgroundColorIndex");
}

void wxExcelSeries::SetMarkerBackgroundColorIndex(long markerBackgroundColorIndex)
{
    InvokePutProperty(wxS("MarkerBackgroundColorIndex"), markerBackgroundColorIndex);
}

wxColour wxExcelSeries::GetMarkerForegroundColor()
{
    WXAUTOEXCEL_PROPERTY_COLOR_GET0("MarkerForegroundColor");
}

void wxExcelSeries::SetMarkerForegroundColor(const wxColour& markerForegroundColor)
{
    InvokePutProperty(wxS("MarkerForegroundColor"), (long)markerForegroundColor.GetRGB());
}

long wxExcelSeries::GetMarkerForegroundColorIndex()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("MarkerForegroundColorIndex");
}

void wxExcelSeries::SetMarkerForegroundColorIndex(long markerForegroundColorIndex)
{
    InvokePutProperty(wxS("MarkerForegroundColorIndex"), markerForegroundColorIndex);
}

long wxExcelSeries::GetMarkerSize()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("MarkerSize");
}

void wxExcelSeries::SetMarkerSize(long markerSize)
{
    InvokePutProperty(wxS("MarkerSize"), markerSize);
}

XlMarkerStyle wxExcelSeries::GetMarkerStyle()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("MarkerStyle", XlMarkerStyle, xlMarkerStyleSquare);
}

void wxExcelSeries::SetMarkerStyle(XlMarkerStyle markerStyle)
{
    InvokePutProperty(wxS("MarkerStyle"), (long)markerStyle);
}

wxString wxExcelSeries::GetName()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("Name");
}

void wxExcelSeries::SetName(const wxString& name)
{
    InvokePutProperty(wxS("Name"), name);
}

XlParentDataLabelOptions wxExcelSeries::GetParentDataLabelOption()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("ParentDataLabelOption", XlParentDataLabelOptions, xlParentDataLabelOptionsNone);
}

void wxExcelSeries::SetParentDataLabelOption(XlParentDataLabelOptions option)
{
    InvokePutProperty(wxS("ParentDataLabelOption"), (long)option);
}

XlChartPictureType wxExcelSeries::GetPictureType()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("PictureType", XlChartPictureType, xlStretch);
}

void wxExcelSeries::SetPictureType(XlChartPictureType pictureType)
{
    InvokePutProperty(wxS("PictureType"), (long)pictureType);
}

long wxExcelSeries::GetPictureUnit()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("PictureUnit");
}

void wxExcelSeries::SetPictureUnit(long pictureUnit)
{
    InvokePutProperty(wxS("PictureUnit"), pictureUnit);
}

double wxExcelSeries::GetPictureUnit2()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("PictureUnit2");
}

void wxExcelSeries::SetPictureUnit2(double pictureUnit2)
{
    InvokePutProperty(wxS("PictureUnit2"), pictureUnit2);
}

long wxExcelSeries::GetPlotOrder()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("PlotOrder");
}

void wxExcelSeries::SetPlotOrder(long plotOrder)
{
    InvokePutProperty(wxS("PlotOrder"), plotOrder);
}

bool wxExcelSeries::GetQuartileCalculationInclusiveMedian()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("QuartileCalculationInclusiveMedian");
}

void wxExcelSeries::SetQuartileCalculationInclusiveMedian(bool value)
{
    InvokePutProperty(wxS("QuartileCalculationInclusiveMedian"), value);
}

bool wxExcelSeries::GetShadow()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("Shadow");
}

void wxExcelSeries::SetShadow(bool shadow)
{
    InvokePutProperty(wxS("Shadow"), shadow);
}

bool wxExcelSeries::GetSmooth()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("Smooth");
}

void wxExcelSeries::SetSmooth(bool smooth)
{
    InvokePutProperty(wxS("Smooth"), smooth);
}

long wxExcelSeries::GetType()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("Type");
}

void wxExcelSeries::SetType(long type)
{
    InvokePutProperty(wxS("Type"), type);
}

bool wxExcelSeries::GetValues(wxExcelRange& rangeValues, wxVariant& variantValues)
{
    wxVariant vResult;    

    variantValues.Clear();
    if ( InvokeGetProperty(wxS("Values"), vResult) )
    {
        if ( vResult.GetType() == wxS("void*") )
        {
            return VariantToObject(vResult, &rangeValues);
        } 
        variantValues = vResult;
        return true;        
    }
    return false;
}

void wxExcelSeries::SetValues(wxExcelRange values)
{
    wxVariant vValues;

    if ( ObjectToVariant(&values, vValues) )
    {
        InvokePutProperty(wxS("Values"), vValues);
    }
}

void wxExcelSeries::SetValues(const wxVariant& values)
{
    InvokePutProperty(wxS("Values"), values);
}

bool wxExcelSeries::GetXValues(wxExcelRange& rangeValues, wxVariant& variantValues)
{
    wxVariant vResult;

    variantValues.Clear();
    if ( InvokeGetProperty(wxS("XValues"), vResult) )
    {
        if ( vResult.GetType() == wxS("void*") )
        {
            return VariantToObject(vResult, &rangeValues);
        } 
        variantValues = vResult;
        return true;        
    }
    return false;
}

void wxExcelSeries::SetXValues(wxExcelRange values)
{
    wxVariant vValues;

    if ( ObjectToVariant(&values, vValues) )
    {
        InvokePutProperty(wxS("XValues"), vValues);
    }
}

void wxExcelSeries::SetXValues(const wxVariant& values)
{
    InvokePutProperty(wxS("XValues"), values);
}

} // namespace wxAutoExcel


#endif // #if WXAUTOEXCEL_USE_CHARTS
