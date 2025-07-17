/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////


#include "wx/wxAutoExcel_prec.h"

#include "wx/wxAutoExcelOLEObjects.h"

#include "wx/wxAutoExcelBorders.h"
#include "wx/wxAutoExcelRange.h"
#include "wx/wxAutoExcelShapeRange.h"
#include "wx/wxAutoExcelInterior.h"

#include "wx/wxAutoExcel_private.h"

namespace wxAutoExcel {
// ***** class wxExcelOLEObject METHODS *****

bool wxExcelOLEObject::Activate()
{
    WXAUTOEXCEL_CALL_METHOD0_BOOL("Activate");

}

bool wxExcelOLEObject::BringToFront()
{
    WXAUTOEXCEL_CALL_METHOD0_BOOL("BringToFront");
}

bool wxExcelOLEObject::Copy()
{
    WXAUTOEXCEL_CALL_METHOD0_BOOL("Copy");
}

bool wxExcelOLEObject::CopyPicture(XlPictureAppearance* appearance, XlCopyPictureFormat* format)
{
    WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT_NAME(Appearance, ((long*)appearance));
    WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT_NAME(Format, ((long*)format));

    WXAUTOEXCEL_CALL_METHOD2("CopyPicture", vAppearance, vFormat, "bool", false);
    return vResult.GetBool();
}

bool wxExcelOLEObject::Cut()
{
    WXAUTOEXCEL_CALL_METHOD0_BOOL("Cut");
}

bool wxExcelOLEObject::Delete()
{
    WXAUTOEXCEL_CALL_METHOD0_BOOL("Delete");
}

wxExcelObject wxExcelOLEObject::Duplicate()
{
    wxExcelObject object;
    WXAUTOEXCEL_CALL_METHOD0_OBJECT("Duplicate", object);
}

bool wxExcelOLEObject::Select(wxXlTribool replace)
{
    WXAUTOEXCEL_OPTIONALCPPTBOOL_TO_OPTIONALVARIANT(Replace, replace);

    WXAUTOEXCEL_CALL_METHOD1_BOOL("Select", vReplace);
}

bool wxExcelOLEObject::SendToBack()
{
    WXAUTOEXCEL_CALL_METHOD0_BOOL("SendToBack");
}

bool wxExcelOLEObject::Update()
{
    WXAUTOEXCEL_CALL_METHOD0_BOOL("Update");
}

bool wxExcelOLEObject::Verb(XlOLEVerb* verb)
{
    WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT_NAME(Verb, ((long*)verb));

    WXAUTOEXCEL_CALL_METHOD1_BOOL("Verb", vVerb);
}

// ***** class wxExcelOLEObject PROPERTIES *****

bool wxExcelOLEObject::GetAutoLoad()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("AutoLoad");
}

void wxExcelOLEObject::SetAutoLoad(bool autoLoad)
{
    InvokePutProperty(wxS("AutoLoad"), autoLoad);
}

bool wxExcelOLEObject::GetAutoUpdate()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("AutoUpdate");
}

wxExcelBorder wxExcelOLEObject::GetBorder()
{
    wxExcelBorder border;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("Border", border);
}

wxExcelRange wxExcelOLEObject::GetBottomRightCell()
{
    wxExcelRange range;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("BottomRightCell", range);
}

bool wxExcelOLEObject::GetEnabled()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("Enabled");
}

void wxExcelOLEObject::SetEnabled(bool enabled)
{
    InvokePutProperty(wxS("Enabled"), enabled);
}

double wxExcelOLEObject::GetHeight()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("Height");
}

void wxExcelOLEObject::SetHeight(double height)
{
    InvokePutProperty(wxS("Height"), height);
}

long wxExcelOLEObject::GetIndex()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("Index");
}

wxExcelInterior wxExcelOLEObject::GetInterior()
{
    wxExcelInterior interior;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("Interior", interior);
}

double wxExcelOLEObject::GetLeft()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("Left");
}

void wxExcelOLEObject::SetLeft(double left)
{
    InvokePutProperty(wxS("Left"), left);
}

wxString wxExcelOLEObject::GetLinkedCell()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("LinkedCell");
}

void wxExcelOLEObject::SetLinkedCell(const wxString& linkedCell)
{
    InvokePutProperty(wxS("LinkedCell"), linkedCell);
}

wxString wxExcelOLEObject::GetListFillRange()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("ListFillRange");
}

void wxExcelOLEObject::SetListFillRange(const wxString& listFillRange)
{
    InvokePutProperty(wxS("ListFillRange"), listFillRange);
}

bool wxExcelOLEObject::GetLocked()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("Locked");
}

void wxExcelOLEObject::SetLocked(bool locked)
{
    InvokePutProperty(wxS("Locked"), locked);
}

wxString wxExcelOLEObject::GetName()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("Name");
}

void wxExcelOLEObject::SetName(const wxString& name)
{
    InvokePutProperty(wxS("Name"), name);
}


XlOLEType wxExcelOLEObject::GetOLEType()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("OLEType", XlOLEType, xlOLELink);
}

XlPlacement wxExcelOLEObject::GetPlacement()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("Placement", XlPlacement, xlMoveAndSize);
}

void wxExcelOLEObject::SetPlacement(XlPlacement placement)
{
    InvokePutProperty(wxS("Placement"), (long)placement);
}

bool wxExcelOLEObject::GetPrintObject()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("PrintObject");
}

void wxExcelOLEObject::SetPrintObject(bool printObject)
{
    InvokePutProperty(wxS("PrintObject"), printObject);
}

wxString wxExcelOLEObject::GetprogID()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("progID");
}

bool wxExcelOLEObject::GetShadow()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("Shadow");
}

void wxExcelOLEObject::SetShadow(bool shadow)
{
    InvokePutProperty(wxS("Shadow"), shadow);
}

#if WXAUTOEXCEL_USE_SHAPES

wxExcelShapeRange wxExcelOLEObject::GetShapeRange()
{
    wxExcelShapeRange shapeRange;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("ShapeRange", shapeRange);
}

#endif // #if WXAUTOEXCEL_USE_SHAPES

wxString wxExcelOLEObject::GetSourceName()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("SourceName");
}

void wxExcelOLEObject::SetSourceName(const wxString& sourceName)
{
    InvokePutProperty(wxS("SourceName"), sourceName);
}

double wxExcelOLEObject::GetTop()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("Top");
}

void wxExcelOLEObject::SetTop(double top)
{
    InvokePutProperty(wxS("Top"), top);
}

wxExcelRange wxExcelOLEObject::GetTopLeftCell()
{
    wxExcelRange range;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("TopLeftCell", range);
}

bool wxExcelOLEObject::GetVisible()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("Visible");
}

void wxExcelOLEObject::SetVisible(bool visible)
{
    InvokePutProperty(wxS("Visible"), visible);
}

double wxExcelOLEObject::GetWidth()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("Width");
}

void wxExcelOLEObject::SetWidth(double width)
{
    InvokePutProperty(wxS("Width"), width);
}

long wxExcelOLEObject::GetZOrder()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("ZOrder");
}

// ***** class wxExcelOLEObjects METHODS *****

wxExcelOLEObject wxExcelOLEObjects::Add(const wxString& classType, const wxString& filename,
                            double* height, wxXlTribool link, wxXlTribool displayAsIcon,
                            const wxString& iconFileName, long* iconIndex, const wxString& iconLabel,
                            double* left, double* width)
{
    wxVariantVector args;

    WXAUTOEXCEL_OPTIONALCPPSTR_TO_OPTIONALVARIANT_NAME_VECTOR(ClassType, classType, args);
    WXAUTOEXCEL_OPTIONALCPPSTR_TO_OPTIONALVARIANT_NAME_VECTOR(Filename, filename, args);
    WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT_NAME_VECTOR(Height, height, args);

    WXAUTOEXCEL_OPTIONALCPPTBOOL_TO_OPTIONALVARIANT_NAME_VECTOR(Link, link, args);
    WXAUTOEXCEL_OPTIONALCPPTBOOL_TO_OPTIONALVARIANT_NAME_VECTOR(DisplayAsIcon, displayAsIcon, args);
    WXAUTOEXCEL_OPTIONALCPPSTR_TO_OPTIONALVARIANT_NAME_VECTOR(IconFileName, iconFileName, args);
    WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT_NAME_VECTOR(IconIndex, iconIndex, args);
    WXAUTOEXCEL_OPTIONALCPPSTR_TO_OPTIONALVARIANT_NAME_VECTOR(IconLabel, iconLabel, args);
    WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT_NAME_VECTOR(Left, left, args);
    WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT_NAME_VECTOR(Width, width, args);


    wxExcelOLEObject object;

    WXAUTOEXCEL_CALL_METHODARR("Add", args, "void*", object);
    VariantToObject(vResult, &object);
    return object;
}

wxExcelOLEObject wxExcelOLEObjects::Item(long index)
{
    wxASSERT( index > 0 );

    wxExcelOLEObject object;
    WXAUTOEXCEL_CALL_METHOD1_OBJECT("Item", index, object);
}

wxExcelOLEObject wxExcelOLEObjects::operator[](long index)
{
    return Item(index);
}

wxExcelOLEObject wxExcelOLEObjects::Item(const wxString& name)
{
    wxExcelOLEObject object;
    WXAUTOEXCEL_CALL_METHOD1_OBJECT("Item", name, object);
}

wxExcelOLEObject wxExcelOLEObjects::operator[](const wxString& name)
{
    return Item(name);
}


// ***** class wxExcelOLEObjects PROPERTIES *****

long wxExcelOLEObjects::GetCount()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("Count");
}


} // namespace wxAutoExcel
