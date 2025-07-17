/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////


#include "wx/wxAutoExcel_prec.h"

#include "wx/wxAutoExcelAxes.h"

#if WXAUTOEXCEL_USE_CHARTS

#include <wx/arrstr.h>

#include "wx/wxAutoExcelAxisTitle.h"
#include "wx/wxAutoExcelBorders.h"
#include "wx/wxAutoExcelChartFormat.h"
#include "wx/wxAutoExcelDisplayUnitLabel.h"
#include "wx/wxAutoExcelGridlines.h"
#include "wx/wxAutoExcelRange.h"
#include "wx/wxAutoExcelTickLabels.h"

#include "wx/wxAutoExcel_private.h"

namespace wxAutoExcel {

// ***** class wxExcelAxis METHODS *****

bool wxExcelAxis::Delete()
{
    WXAUTOEXCEL_CALL_METHOD0_BOOL("Delete");
}

bool wxExcelAxis::Select()
{
    WXAUTOEXCEL_CALL_METHOD0_BOOL("Select");
}

// ***** class wxExcelAxis PROPERTIES *****


bool wxExcelAxis::GetAxisBetweenCategories()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("AxisBetweenCategories");
}

void wxExcelAxis::SetAxisBetweenCategories(bool axisBetweenCategories)
{
    InvokePutProperty(wxS("AxisBetweenCategories"), axisBetweenCategories);
}

XlAxisGroup wxExcelAxis::GetAxisGroup()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("AxisGroup", XlAxisGroup, xlPrimary);
}

wxExcelAxisTitle wxExcelAxis::GetAxisTitle()
{
    wxExcelAxisTitle axisTitle;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("AxisTitle", axisTitle);
}

XlTimeUnit wxExcelAxis::GetBaseUnit()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("BaseUnit", XlTimeUnit,  xlDays);
}

void wxExcelAxis::SetBaseUnit(XlTimeUnit baseUnit)
{
    InvokePutProperty(wxS("BaseUnit"), (long)baseUnit);
}

bool wxExcelAxis::GetBaseUnitIsAuto()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("BaseUnitIsAuto");
}

void wxExcelAxis::SetBaseUnitIsAuto(bool baseUnitIsAuto)
{
    InvokePutProperty(wxS("BaseUnitIsAuto"), baseUnitIsAuto);
}

wxExcelBorder wxExcelAxis::GetBorder()
{
    wxExcelBorder border;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("Border", border);
}

wxArrayString wxExcelAxis::GetCategoryNames()
{
    wxArrayString strings;
    wxVariant vResult;

    if ( InvokeGetProperty(wxS("CategoryNames"), vResult) )
    {
        const wxString type = vResult.GetType();

        if ( type == wxS("arrstring") )
        {
            return vResult.GetArrayString();
        }
        if ( type == wxS("list") )
        {
            strings.reserve(vResult.GetCount());
            for ( size_t i = 0; i < vResult.GetCount(); i++)
            {
                strings.push_back(vResult[i].GetString());
            }
        } else if ( type == wxS("void*" ) )
        {
            //@FIXME maybe can also return a Range - test it?
            wxExcelRange range;
            VariantToObject(vResult, &range); // make sure we don't leave a reference hanging
            wxFAIL;
        }
    }

    return strings;
}

void wxExcelAxis::SetCategoryNames(wxExcelRange categoryNames)
{
    wxVariant vCategoryNames;

    if ( ObjectToVariant(&categoryNames, vCategoryNames) )
    {
        InvokePutProperty(wxS("CategoryNames"), vCategoryNames);
    }
}

void wxExcelAxis::SetCategoryNames(const wxArrayString& categoryNames)
{
    wxVariant vCategoryNames(categoryNames);

    InvokePutProperty(wxS("CategoryNames"), vCategoryNames);
}


XlCategoryType wxExcelAxis::GetCategoryType()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("CategoryType", XlCategoryType, xlAutomaticScale);
}

void wxExcelAxis::SetCategoryType(XlCategoryType categoryType)
{
    InvokePutProperty(wxS("CategoryType"), (long)categoryType);
}

long wxExcelAxis::GetCrosses()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("Crosses");
}

void wxExcelAxis::SetCrosses(long crosses)
{
    InvokePutProperty(wxS("Crosses"), crosses);
}

double wxExcelAxis::GetCrossesAt()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("CrossesAt");
}

void wxExcelAxis::SetCrossesAt(double crossesAt)
{
    InvokePutProperty(wxS("CrossesAt"), crossesAt);
}

long wxExcelAxis::GetDisplayUnit()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("DisplayUnit");
}

void wxExcelAxis::SetDisplayUnit(long displayUnit)
{
    InvokePutProperty(wxS("DisplayUnit"), displayUnit);
}

double wxExcelAxis::GetDisplayUnitCustom()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("DisplayUnitCustom");
}

void wxExcelAxis::SetDisplayUnitCustom(double displayUnitCustom)
{
    InvokePutProperty(wxS("DisplayUnitCustom"), displayUnitCustom);
}

wxExcelDisplayUnitLabel wxExcelAxis::GetDisplayUnitLabel()
{
    wxExcelDisplayUnitLabel displayUnitLabel;
    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("DisplayUnitLabel", displayUnitLabel);
}

wxExcelChartFormat wxExcelAxis::GetFormat()
{
    wxExcelChartFormat chartFormat;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("Format", chartFormat);
}

bool wxExcelAxis::GetHasDisplayUnitLabel()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("HasDisplayUnitLabel");
}

void wxExcelAxis::SetHasDisplayUnitLabel(bool hasDisplayUnitLabel)
{
    InvokePutProperty(wxS("HasDisplayUnitLabel"), hasDisplayUnitLabel);
}

bool wxExcelAxis::GetHasMajorGridlines()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("HasMajorGridlines");
}

void wxExcelAxis::SetHasMajorGridlines(bool hasMajorGridlines)
{
    InvokePutProperty(wxS("HasMajorGridlines"), hasMajorGridlines);
}

bool wxExcelAxis::GetHasMinorGridlines()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("HasMinorGridlines");
}

void wxExcelAxis::SetHasMinorGridlines(bool hasMinorGridlines)
{
    InvokePutProperty(wxS("HasMinorGridlines"), hasMinorGridlines);
}

bool wxExcelAxis::GetHasTitle()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("HasTitle");
}

void wxExcelAxis::SetHasTitle(bool hasTitle)
{
    InvokePutProperty(wxS("HasTitle"), hasTitle);
}

double wxExcelAxis::GetHeight()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("Height");
}

double wxExcelAxis::GetLeft()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("Left");
}

double wxExcelAxis::GetLogBase()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("LogBase");
}

void wxExcelAxis::SetLogBase(double logBase)
{
    InvokePutProperty(wxS("LogBase"), logBase);
}

wxExcelGridlines wxExcelAxis::GetMajorGridlines()
{
    wxExcelGridlines gridlines;
    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("MajorGridlines", gridlines);
}

XlTickMark wxExcelAxis::GetMajorTickMark()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("MajorTickMark", XlTickMark, xlTickMarkNone);
}

void wxExcelAxis::SetMajorTickMark(XlTickMark majorTickMark)
{
    InvokePutProperty(wxS("MajorTickMark"), (long)majorTickMark);
}

double wxExcelAxis::GetMajorUnit()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("MajorUnit");
}

void wxExcelAxis::SetMajorUnit(double majorUnit)
{
    InvokePutProperty(wxS("MajorUnit"), majorUnit);
}

bool wxExcelAxis::GetMajorUnitIsAuto()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("MajorUnitIsAuto");
}

void wxExcelAxis::SetMajorUnitIsAuto(bool majorUnitIsAuto)
{
    InvokePutProperty(wxS("MajorUnitIsAuto"), majorUnitIsAuto);
}

XlTimeUnit wxExcelAxis::GetMajorUnitScale()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("MajorUnitScale", XlTimeUnit, xlDays);
}

void wxExcelAxis::SetMajorUnitScale(XlTimeUnit majorUnitScale)
{
    InvokePutProperty(wxS("MajorUnitScale"), (long)majorUnitScale);
}

double wxExcelAxis::GetMaximumScale()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("MaximumScale");
}

void wxExcelAxis::SetMaximumScale(double maximumScale)
{
    InvokePutProperty(wxS("MaximumScale"), maximumScale);
}

bool wxExcelAxis::GetMaximumScaleIsAuto()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("MaximumScaleIsAuto");
}

void wxExcelAxis::SetMaximumScaleIsAuto(bool maximumScaleIsAuto)
{
    InvokePutProperty(wxS("MaximumScaleIsAuto"), maximumScaleIsAuto);
}

double wxExcelAxis::GetMinimumScale()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("MinimumScale");
}

void wxExcelAxis::SetMinimumScale(double minimumScale)
{
    InvokePutProperty(wxS("MinimumScale"), minimumScale);
}

bool wxExcelAxis::GetMinimumScaleIsAuto()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("MinimumScaleIsAuto");
}

void wxExcelAxis::SetMinimumScaleIsAuto(bool minimumScaleIsAuto)
{
    InvokePutProperty(wxS("MinimumScaleIsAuto"), minimumScaleIsAuto);
}

wxExcelGridlines wxExcelAxis::GetMinorGridlines()
{
    wxExcelGridlines gridlines;
    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("MinorGridlines", gridlines);
}

XlTickMark wxExcelAxis::GetMinorTickMark()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("MinorTickMark", XlTickMark, xlTickMarkNone);
}

void wxExcelAxis::SetMinorTickMark(XlTickMark minorTickMark)
{
    InvokePutProperty(wxS("MinorTickMark"), (long)minorTickMark);
}

double wxExcelAxis::GetMinorUnit()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("MinorUnit");
}

void wxExcelAxis::SetMinorUnit(double minorUnit)
{
    InvokePutProperty(wxS("MinorUnit"), minorUnit);
}

bool wxExcelAxis::GetMinorUnitIsAuto()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("MinorUnitIsAuto");
}

void wxExcelAxis::SetMinorUnitIsAuto(bool minorUnitIsAuto)
{
    InvokePutProperty(wxS("MinorUnitIsAuto"), minorUnitIsAuto);
}

XlTimeUnit wxExcelAxis::GetMinorUnitScale()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("MinorUnitScale", XlTimeUnit, xlDays);
}

void wxExcelAxis::SetMinorUnitScale(XlTimeUnit minorUnitScale)
{
    InvokePutProperty(wxS("MinorUnitScale"), (long)minorUnitScale);
}

bool wxExcelAxis::GetReversePlotOrder()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("ReversePlotOrder");
}

void wxExcelAxis::SetReversePlotOrder(bool reversePlotOrder)
{
    InvokePutProperty(wxS("ReversePlotOrder"), reversePlotOrder);
}

XlScaleType wxExcelAxis::GetScaleType()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("ScaleType", XlScaleType, xlScaleLinear);
}

void wxExcelAxis::SetScaleType(XlScaleType scaleType)
{
    InvokePutProperty(wxS("ScaleType"), (long)scaleType);
}

XlTickLabelPosition wxExcelAxis::GetTickLabelPosition()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("TickLabelPosition", XlTickLabelPosition, xlTickLabelPositionNone);
}

void wxExcelAxis::SetTickLabelPosition(XlTickLabelPosition tickLabelPosition)
{
    InvokePutProperty(wxS("TickLabelPosition"), (long)tickLabelPosition);
}

wxExcelTickLabels wxExcelAxis::GetTickLabels()
{
    wxExcelTickLabels tickLabels;
    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("TickLabels", tickLabels);
}

long wxExcelAxis::GetTickLabelSpacing()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("TickLabelSpacing");
}

void wxExcelAxis::SetTickLabelSpacing(long tickLabelSpacing)
{
    InvokePutProperty(wxS("TickLabelSpacing"), tickLabelSpacing);
}

bool wxExcelAxis::GetTickLabelSpacingIsAuto()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("TickLabelSpacingIsAuto");
}

void wxExcelAxis::SetTickLabelSpacingIsAuto(bool tickLabelSpacingIsAuto)
{
    InvokePutProperty(wxS("TickLabelSpacingIsAuto"), tickLabelSpacingIsAuto);
}

long wxExcelAxis::GetTickMarkSpacing()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("TickMarkSpacing");
}

void wxExcelAxis::SetTickMarkSpacing(long tickMarkSpacing)
{
    InvokePutProperty(wxS("TickMarkSpacing"), tickMarkSpacing);
}

double wxExcelAxis::GetTop()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("Top");
}

XlAxisType wxExcelAxis::GetType()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("Type", XlAxisType, xlCategory);
}

double wxExcelAxis::GetWidth()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("Width");
}

// ***** class wxExcelAxes METHODS *****

wxExcelAxis wxExcelAxes::Item(long index)
{
    wxASSERT( index > 0 );

    wxExcelAxis axis;
    WXAUTOEXCEL_CALL_METHOD1_OBJECT("Item", index, axis);
}

wxExcelAxis wxExcelAxes::operator[](long index)
{
    return Item(index);
}

// ***** class wxExcelAxes PROPERTIES *****


long wxExcelAxes::GetCount()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("Count");
}



} // namespace wxAutoExcel


#endif // #if WXAUTOEXCEL_USE_CHARTS
