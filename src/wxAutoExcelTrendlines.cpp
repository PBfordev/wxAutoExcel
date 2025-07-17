/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// License:     MIT license
/////////////////////////////////////////////////////////////////////////////


#include "wx/wxAutoExcel_prec.h"

#include "wx/wxAutoExcelTrendLines.h"

#if WXAUTOEXCEL_USE_CHARTS

#include "wx/wxAutoExcelBorders.h"
#include "wx/wxAutoExcelChartFormat.h"
#include "wx/wxAutoExcelDataLabels.h"

#include "wx/wxAutoExcel_private.h"

namespace wxAutoExcel {

// ***** class wxExcelTrendline METHODS *****

bool wxExcelTrendline::ClearFormats()
{
    WXAUTOEXCEL_CALL_METHOD0_BOOL("ClearFormats");
}

bool wxExcelTrendline::Delete()
{
    WXAUTOEXCEL_CALL_METHOD0_BOOL("Delete");
}


bool wxExcelTrendline::Select()
{
    WXAUTOEXCEL_CALL_METHOD0_BOOL("Select");
}


// ***** class wxExcelTrendline PROPERTIES *****


long wxExcelTrendline::GetBackward()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("Backward");
}

void wxExcelTrendline::SetBackward(long backward)
{
    InvokePutProperty(wxS("Backward"), backward);
}

double wxExcelTrendline::GetBackward2()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("Backward2");
}

void wxExcelTrendline::SetBackward2(double backward2)
{
    InvokePutProperty(wxS("Backward2"), backward2);
}

wxExcelBorder wxExcelTrendline::GetBorder()
{
    wxExcelBorder border;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("Border", border);
}

wxExcelDataLabel wxExcelTrendline::GetDataLabel()
{
    wxExcelDataLabel dataLabel;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("DataLabel", dataLabel);
}

bool wxExcelTrendline::GetDisplayEquation()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("DisplayEquation");
}

void wxExcelTrendline::SetDisplayEquation(bool displayEquation)
{
    InvokePutProperty(wxS("DisplayEquation"), displayEquation);
}

bool wxExcelTrendline::GetDisplayRSquared()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("DisplayRSquared");
}

void wxExcelTrendline::SetDisplayRSquared(bool displayRSquared)
{
    InvokePutProperty(wxS("DisplayRSquared"), displayRSquared);
}

wxExcelChartFormat wxExcelTrendline::GetFormat()
{
    wxExcelChartFormat chartFormat;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("Format", chartFormat);
}

long wxExcelTrendline::GetForward()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("Forward");
}

void wxExcelTrendline::SetForward(long forward)
{
    InvokePutProperty(wxS("Forward"), forward);
}

double wxExcelTrendline::GetForward2()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("Forward2");
}

void wxExcelTrendline::SetForward2(double forward2)
{
    InvokePutProperty(wxS("Forward2"), forward2);
}

long wxExcelTrendline::GetIndex()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("Index");
}

double wxExcelTrendline::GetIntercept()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("Intercept");
}

void wxExcelTrendline::SetIntercept(double intercept)
{
    InvokePutProperty(wxS("Intercept"), intercept);
}

bool wxExcelTrendline::GetInterceptIsAuto()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("InterceptIsAuto");
}

void wxExcelTrendline::SetInterceptIsAuto(bool interceptIsAuto)
{
    InvokePutProperty(wxS("InterceptIsAuto"), interceptIsAuto);
}

wxString wxExcelTrendline::GetName()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("Name");
}

void wxExcelTrendline::SetName(const wxString& name)
{
    InvokePutProperty(wxS("Name"), name);
}

bool wxExcelTrendline::GetNameIsAuto()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("NameIsAuto");
}

void wxExcelTrendline::SetNameIsAuto(bool nameIsAuto)
{
    InvokePutProperty(wxS("NameIsAuto"), nameIsAuto);
}

long wxExcelTrendline::GetOrder()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("Order");
}

void wxExcelTrendline::SetOrder(long order)
{
    InvokePutProperty(wxS("Order"), order);
}

long wxExcelTrendline::GetPeriod()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("Period");
}

void wxExcelTrendline::SetPeriod(long period)
{
    InvokePutProperty(wxS("Period"), period);
}

XlTrendlineType wxExcelTrendline::GetType()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("Type", XlTrendlineType, xlPolynomial);
}

void wxExcelTrendline::SetType(XlTrendlineType type)
{
    InvokePutProperty(wxS("Type"), (long)type);
}


// ***** class wxExcelTrendlines METHODS *****

wxExcelTrendline wxExcelTrendlines::Add(XlTrendlineType* type, long* order, long* period,
                            long* forward, long* backward, double* intercept,
                            wxXlTribool displayEquation, wxXlTribool displayRSquared,
                            const wxString& name)
{
    wxVariantVector args;

    WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT_NAME_VECTOR(Type, ((long*)type), args);
    WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT_NAME_VECTOR(Order, order, args);
    WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT_NAME_VECTOR(Period, period, args);
    WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT_NAME_VECTOR(Forward, forward, args);
    WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT_NAME_VECTOR(Backward, backward, args);
    WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT_NAME_VECTOR(Intercept, intercept, args);
    WXAUTOEXCEL_OPTIONALCPPTBOOL_TO_OPTIONALVARIANT_NAME_VECTOR(DisplayEquation, displayEquation, args);
    WXAUTOEXCEL_OPTIONALCPPTBOOL_TO_OPTIONALVARIANT_NAME_VECTOR(DisplayRSquared, displayRSquared, args);
    WXAUTOEXCEL_OPTIONALCPPSTR_TO_OPTIONALVARIANT_NAME_VECTOR(Name, name, args);

    wxExcelTrendline line;

    WXAUTOEXCEL_CALL_METHODARR("Add", args, "void*", line);
    VariantToObject(vResult, &line);
    return line;
}

wxExcelTrendline wxExcelTrendlines::Item(long index)
{
    wxExcelTrendline item;
    WXAUTOEXCEL_CALL_METHOD1_OBJECT("Item", index, item);
}

wxExcelTrendline wxExcelTrendlines::operator[](long index)
{
    return Item(index);
}

// ***** class wxExcelTrendlines PROPERTIES *****


long wxExcelTrendlines::GetCount()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("Count");
}

} // namespace wxAutoExcel

#endif // #if WXAUTOEXCEL_USE_CHARTS
