/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Modified by:
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////


#include "wx/wxAutoExcel_prec.h"

#include "wx/wxAutoExcelChartObjects.h"

#if WXAUTOEXCEL_USE_CHARTS

#include "wx/wxAutoExcelChart.h"
#include "wx/wxAutoExcelRange.h"
#include "wx/wxAutoExcelShapeRange.h"

#include "wx/wxAutoExcel_private.h"

namespace wxAutoExcel {

// ***** class wxExcelChartObject METHODS *****

bool wxExcelChartObject::Activate()
{
    WXAUTOEXCEL_CALL_METHOD0_BOOL("Activate");
}

bool wxExcelChartObject::BringToFront()
{
    WXAUTOEXCEL_CALL_METHOD0_BOOL("BringToFront");
}

bool wxExcelChartObject::Copy()
{
    WXAUTOEXCEL_CALL_METHOD0_BOOL("Copy");
}

bool wxExcelChartObject::CopyPicture(XlPictureAppearance* appearance, XlCopyPictureFormat* format)
{
    WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT_NAME(Appearance, ((long*)appearance));
    WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT_NAME(Format, ((long*)format));

    WXAUTOEXCEL_CALL_METHOD2("CopyPicture", vAppearance, vFormat, "bool", false);
    return vResult.GetBool();
}

bool wxExcelChartObject::Cut()
{
    WXAUTOEXCEL_CALL_METHOD0_BOOL("Cut");
}

bool wxExcelChartObject::Delete()
{
    WXAUTOEXCEL_CALL_METHOD0_BOOL("Delete");
}

wxExcelObject wxExcelChartObject::Duplicate()
{
    wxExcelObject object;
    WXAUTOEXCEL_CALL_METHOD0_OBJECT("Duplicate", object);
}

bool wxExcelChartObject::Select(wxXlTribool replace)
{
    WXAUTOEXCEL_OPTIONALCPPTBOOL_TO_OPTIONALVARIANT(Replace, replace);
    WXAUTOEXCEL_CALL_METHOD1_BOOL("Select", vReplace);
}

bool wxExcelChartObject::SendToBack()
{
    WXAUTOEXCEL_CALL_METHOD0_BOOL("SendToBack");
}

// ***** class wxExcelChartObject PROPERTIES *****


wxExcelRange wxExcelChartObject::GetBottomRightCell()
{
    wxExcelRange range;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("BottomRightCell", range);
}

wxExcelChart wxExcelChartObject::GetChart()
{
    wxExcelChart chart;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("Chart", chart);
}

bool wxExcelChartObject::GetEnabled()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("Enabled");
}

void wxExcelChartObject::SetEnabled(bool enabled)
{
    InvokePutProperty(wxS("Enabled"), enabled);
}

double wxExcelChartObject::GetHeight()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("Height");
}

void wxExcelChartObject::SetHeight(double height)
{
    InvokePutProperty(wxS("Height"), height);
}

long wxExcelChartObject::GetIndex()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("Index");
}

double wxExcelChartObject::GetLeft()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("Left");
}

void wxExcelChartObject::SetLeft(double left)
{
    InvokePutProperty(wxS("Left"), left);
}

bool wxExcelChartObject::GetLocked()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("Locked");
}

void wxExcelChartObject::SetLocked(bool locked)
{
    InvokePutProperty(wxS("Locked"), locked);
}

wxString wxExcelChartObject::GetName()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("Name");
}

XlPlacement wxExcelChartObject::GetPlacement()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("Placement", XlPlacement, xlMoveAndSize);
}

void wxExcelChartObject::SetPlacement(XlPlacement placement)
{
    InvokePutProperty(wxS("Placement"), (long)placement);
}

bool wxExcelChartObject::GetPrintObject()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("PrintObject");
}

void wxExcelChartObject::SetPrintObject(bool printObject)
{
    InvokePutProperty(wxS("PrintObject"), printObject);
}

bool wxExcelChartObject::GetProtectChartObject()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("ProtectChartObject");
}

void wxExcelChartObject::SetProtectChartObject(bool protectChartObject)
{
    InvokePutProperty(wxS("ProtectChartObject"), protectChartObject);
}

bool wxExcelChartObject::GetRoundedCorners()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("RoundedCorners");
}

void wxExcelChartObject::SetRoundedCorners(bool roundedCorners)
{
    InvokePutProperty(wxS("RoundedCorners"), roundedCorners);
}

bool wxExcelChartObject::GetShadow()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("Shadow");
}

void wxExcelChartObject::SetShadow(bool shadow)
{
    InvokePutProperty(wxS("Shadow"), shadow);
}

#if WXAUTOEXCEL_USE_SHAPES

wxExcelShapeRange wxExcelChartObject::GetShapeRange()
{
    wxExcelShapeRange shapeRange;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("ShapeRange", shapeRange);
}

#endif // #if WXAUTOEXCEL_USE_SHAPES

double wxExcelChartObject::GetTop()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("Top");
}

void wxExcelChartObject::SetTop(double top)
{
    InvokePutProperty(wxS("Top"), top);
}

wxExcelRange wxExcelChartObject::GetTopLeftCell()
{
    wxExcelRange range;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("TopLeftCell", range);
}

bool wxExcelChartObject::GetVisible()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("Visible");
}

void wxExcelChartObject::SetVisible(bool visible)
{
    InvokePutProperty(wxS("Visible"), visible);
}

double wxExcelChartObject::GetWidth()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("Width");
}

void wxExcelChartObject::SetWidth(double width)
{
    InvokePutProperty(wxS("Width"), width);
}

long wxExcelChartObject::GetZOrder()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("ZOrder");
}

// ***** class wxExcelChartObjects METHODS *****

 wxExcelChartObject wxExcelChartObjects::Add(double left, double top, double width, double height)
{
    wxExcelChartObject chartObject;

    WXAUTOEXCEL_CALL_METHOD4("Add", left, top, width, height, "void*", chartObject);
    VariantToObject(vResult, &chartObject);
    return chartObject;
}


bool wxExcelChartObjects::Copy()
{
    WXAUTOEXCEL_CALL_METHOD0_BOOL("Copy");
}

bool wxExcelChartObjects::CopyPicture(XlPictureAppearance* appearance, XlCopyPictureFormat* format)
{
    WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT_NAME(Appearance, ((long*)appearance));
    WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT_NAME(Format, ((long*)format));

    WXAUTOEXCEL_CALL_METHOD2("CopyPicture", vAppearance, vFormat, "bool", false);
    return vResult.GetBool();
}

bool wxExcelChartObjects::Cut()
{
    WXAUTOEXCEL_CALL_METHOD0_BOOL("Cut");
}

bool wxExcelChartObjects::Delete()
{
    WXAUTOEXCEL_CALL_METHOD0_BOOL("Delete");
}

wxExcelChartObjects wxExcelChartObjects::Duplicate()
{
    wxExcelChartObjects object;
    WXAUTOEXCEL_CALL_METHOD0_OBJECT("Duplicate", object);
}

wxExcelChartObject wxExcelChartObjects::Item(long index)
{
    wxExcelChartObject object;
    WXAUTOEXCEL_CALL_METHOD1_OBJECT("Item", index, object);
}

wxExcelChartObject wxExcelChartObjects::operator[](long index)
{
    return Item(index);
}

wxExcelChartObject wxExcelChartObjects::Item(const wxString& name)
{
    wxExcelChartObject object;
    WXAUTOEXCEL_CALL_METHOD1_OBJECT("Item", name, object);
}

wxExcelChartObject wxExcelChartObjects::operator[](const wxString& name)
{
    return Item(name);
}

bool wxExcelChartObjects::Select(wxXlTribool replace)
{
    WXAUTOEXCEL_OPTIONALCPPTBOOL_TO_OPTIONALVARIANT(Replace, replace);
    WXAUTOEXCEL_CALL_METHOD1_BOOL("Select", vReplace);
}


// ***** class wxExcelChartObjects PROPERTIES *****


long wxExcelChartObjects::GetCount()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("Count");
}

bool wxExcelChartObjects::GetEnabled()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("Enabled");
}

void wxExcelChartObjects::SetEnabled(bool enabled)
{
    InvokePutProperty(wxS("Enabled"), enabled);
}

double wxExcelChartObjects::GetHeight()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("Height");
}

void wxExcelChartObjects::SetHeight(double height)
{
    InvokePutProperty(wxS("Height"), height);
}

double wxExcelChartObjects::GetLeft()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("Left");
}

void wxExcelChartObjects::SetLeft(double left)
{
    InvokePutProperty(wxS("Left"), left);
}


bool wxExcelChartObjects::GetProtectChartObject()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("ProtectChartObject");
}

void wxExcelChartObjects::SetProtectChartObject(bool protectChartObject)
{
    InvokePutProperty(wxS("ProtectChartObject"), protectChartObject);
}

#if WXAUTOEXCEL_USE_SHAPES

wxExcelShapeRange wxExcelChartObjects::GetShapeRange()
{
    wxExcelShapeRange shapeRange;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("ShapeRange", shapeRange);
}

#endif // #if WXAUTOEXCEL_USE_SHAPES

double wxExcelChartObjects::GetTop()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("Top");
}

void wxExcelChartObjects::SetTop(double top)
{
    InvokePutProperty(wxS("Top"), top);
}

bool wxExcelChartObjects::GetVisible()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("Visible");
}

void wxExcelChartObjects::SetVisible(bool visible)
{
    InvokePutProperty(wxS("Visible"), visible);
}

double wxExcelChartObjects::GetWidth()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("Width");
}

void wxExcelChartObjects::SetWidth(double width)
{
    InvokePutProperty(wxS("Width"), width);
}

} // namespace wxAutoExcel

#endif // #if WXAUTOEXCEL_USE_CHARTS
