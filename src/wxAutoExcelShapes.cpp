/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// License:     MIT license
/////////////////////////////////////////////////////////////////////////////


#include "wx/wxAutoExcel_prec.h"

#include "wx/wxAutoExcelShapes.h"

#if WXAUTOEXCEL_USE_SHAPES

#include <wx/vector.h>
#include <wx/geometry.h>
#include <wx/msw/ole/safearray.h>

#include "wx/wxAutoExcelChart.h"
#include "wx/wxAutoExcelFreeformBuilder.h"
#include "wx/wxAutoExcelShape.h"
#include "wx/wxAutoExcelShapeRange.h"

#include "wx/wxAutoExcel_private.h"

namespace wxAutoExcel {

// ***** class wxExcelShapes METHODS *****

wxExcelShape wxExcelShapes::Add3DModel(const wxString& fileName, 
                                       wxXlTribool linkToFile,
                                       wxXlTribool saveWithDocument, 
                                       double* left, double* top,
                                       double* width, double* height)
{
    wxVariantVector args;
    wxExcelShape shape;

    args.push_back(fileName);

    WXAUTOEXCEL_OPTIONALCPPTBOOL_TO_OPTIONALVARIANT_NAME_VECTOR(LinkToFile, linkToFile, args);
    WXAUTOEXCEL_OPTIONALCPPTBOOL_TO_OPTIONALVARIANT_NAME_VECTOR(SaveWithDocument, saveWithDocument, args);

    WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT_NAME_VECTOR(Left, left, args);
    WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT_NAME_VECTOR(Top, top, args);
    WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT_NAME_VECTOR(Width, width, args);
    WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT_NAME_VECTOR(Height, height, args);

    WXAUTOEXCEL_CALL_METHODARR("Add3DModel", args, "void*", shape);
    VariantToObject(vResult, &shape);
    return shape;
}

wxExcelShape wxExcelShapes::AddCallout(MsoCalloutType type, double left, double top,
                               double width, double height)
{
    wxExcelShape shape;

    WXAUTOEXCEL_CALL_METHOD5("AddCallout", (long)type, left, top, width, height, "void*", shape);
    VariantToObject(vResult, &shape);
    return shape;
}

#if WXAUTOEXCEL_USE_CHARTS

wxExcelShape wxExcelShapes::AddChart(XlChartType type, double left, double top,
                                     double width, double height)
{
    wxExcelShape chart;

    WXAUTOEXCEL_CALL_METHOD5("AddChart", (long)type, left, top, width, height, "void*", chart);
    VariantToObject(vResult, &chart);
    return chart;
}

wxExcelShape wxExcelShapes::AddChart2(long style, XlChartType* type, double* left, double* top, double* width, double* height, wxXlTribool newLayout)
{
    wxExcelShape chart;
    wxVariantVector args;

    args.push_back(wxVariant(style, wxS("Style")));

    WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT_NAME_VECTOR(XlChartType, type, args);
    WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT_NAME_VECTOR(Left, left, args);
    WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT_NAME_VECTOR(Top, top, args);
    WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT_NAME_VECTOR(Width, width, args);
    WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT_NAME_VECTOR(Height, height, args);
    WXAUTOEXCEL_OPTIONALCPPTBOOL_TO_OPTIONALVARIANT_NAME_VECTOR(NewLayout, newLayout, args);

    WXAUTOEXCEL_CALL_METHODARR("AddChart2", args, "void*", chart);
    VariantToObject(vResult, &chart);
    return chart;
}

#endif // #if WXAUTOEXCEL_USE_CHARTS

wxExcelShape wxExcelShapes::AddConnector(MsoConnectorType type, double beginX, double beginY,
                                         double endX, double endY)
{
    wxExcelShape shape;

    WXAUTOEXCEL_CALL_METHOD5("AddConnector", (long)type, beginX, beginY, endX, endY, "void*", shape);
    VariantToObject(vResult, &shape);
    return shape;
}

wxExcelShape wxExcelShapes::AddCurve(const wxVector<wxPoint2DDouble>& points)
{
    wxExcelShape shape;
    wxVariant vPoints, vPoint;

    vPoints.NullList();
    for ( size_t i = 0; i < points.size(); i++ )
    {
        const wxPoint2DDouble& p = points[i];

        vPoint.NullList();
        vPoint.Append(p.m_x);
        vPoint.Append(p.m_y);

        vPoints.Append(vPoint);
    }

    WXAUTOEXCEL_CALL_METHOD1_OBJECT("AddCurve", vPoints, shape);
}

wxExcelShape wxExcelShapes::AddFormControl(XlFormControl type, double left, double top,
                                           double width, double height)
{
    wxExcelShape shape;

    WXAUTOEXCEL_CALL_METHOD5("AddFormControl", (long)type, left, top, width, height, "void*", shape);
    VariantToObject(vResult, &shape);
    return shape;
}

wxExcelShape wxExcelShapes::AddLabel(MsoTextOrientation orientation, double left, double top,
                                     double width, double height)
{
    wxExcelShape shape;

    WXAUTOEXCEL_CALL_METHOD5("AddLabel", (long)orientation, left, top, width, height, "void*", shape);
    VariantToObject(vResult, &shape);
    return shape;
}

wxExcelShape wxExcelShapes::AddLine(double beginX, double beginY, double endX, double endY)
{
    wxExcelShape shape;

    WXAUTOEXCEL_CALL_METHOD4("AddLine", beginX, beginY, endX, endY, "void*", shape);
    VariantToObject(vResult, &shape);
    return shape;
}

wxExcelShape wxExcelShapes::AddOLEObject(const wxString& classType, const wxString& filename,
                                         wxXlTribool link, wxXlTribool displayAsIcon,
                                         const wxString& iconFileName, long* iconIndex, const wxString& iconLabel,
                                         double* left, double* top, double* width, double* height)
{
    wxVariantVector args;

    WXAUTOEXCEL_OPTIONALCPPSTR_TO_OPTIONALVARIANT_NAME_VECTOR(ClassType, classType, args);
    WXAUTOEXCEL_OPTIONALCPPSTR_TO_OPTIONALVARIANT_NAME_VECTOR(Filename, filename, args);
    WXAUTOEXCEL_OPTIONALCPPTBOOL_TO_OPTIONALVARIANT_NAME_VECTOR(Link, link, args);
    WXAUTOEXCEL_OPTIONALCPPTBOOL_TO_OPTIONALVARIANT_NAME_VECTOR(DisplayAsIcon, displayAsIcon, args);
    WXAUTOEXCEL_OPTIONALCPPSTR_TO_OPTIONALVARIANT_NAME_VECTOR(IconFileName, iconFileName, args);
    WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT_NAME_VECTOR(IconIndex, iconIndex, args);
    WXAUTOEXCEL_OPTIONALCPPSTR_TO_OPTIONALVARIANT_NAME_VECTOR(IconLabel, iconLabel, args);
    WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT_NAME_VECTOR(Left, left, args);
    WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT_NAME_VECTOR(Top, top, args);
    WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT_NAME_VECTOR(Width, width, args);
    WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT_NAME_VECTOR(Height, height, args);

    wxExcelShape shape;

    WXAUTOEXCEL_CALL_METHODARR("AddOLEObject", args, "void*", shape);
    VariantToObject(vResult, &shape);
    return shape;
}

wxExcelShape wxExcelShapes::AddPicture(const wxString& fileName, MsoTriState linkToFile,
                                       MsoTriState saveWithDocument, double left, double top,
                                       double width, double height)
{
    wxVariantVector args;

    args.push_back(fileName);
    args.push_back((long)linkToFile);
    args.push_back((long)saveWithDocument);
    args.push_back(left);
    args.push_back(top);
    args.push_back(width);
    args.push_back(height);

    wxExcelShape shape;

    WXAUTOEXCEL_CALL_METHODARR("AddPicture", args, "void*", shape);
    VariantToObject(vResult, &shape);
    return shape;
}

wxExcelShape wxExcelShapes::AddPicture2(const wxString& fileName, MsoTriState linkToFile,
                                       MsoTriState saveWithDocument, double left, double top,
                                       double* width, double* height,
                                       MsoPictureCompress* compress)
{
    wxVariantVector args;
    wxExcelShape shape;

    args.push_back(fileName);
    args.push_back((long)linkToFile);
    args.push_back((long)saveWithDocument);
    args.push_back(left);
    args.push_back(top);
    WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT_NAME_VECTOR(Width, width, args);
    WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT_NAME_VECTOR(Height, height, args);
    WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT_NAME_VECTOR(Compress, (long*)compress, args);
    

    WXAUTOEXCEL_CALL_METHODARR("AddPicture2", args, "void*", shape);
    VariantToObject(vResult, &shape);
    return shape;
}

wxExcelShape wxExcelShapes::AddPolyline(const wxVector<wxPoint2DDouble>& points)
{
    wxExcelShape shape;
    SAFEARRAYBOUND sab[2];
    wxSafeArray<VT_R4> sa;

    sab[0].lLbound = 0;
    sab[0].cElements = points.size();
    sab[1].lLbound = 0;
    sab[1].cElements = 2; // x and y

    if ( !sa.Create(sab, 2) )
    {
        return shape;
    }

    long indices[2];
    for ( size_t i = 0; i < points.size(); i++ )
    {
        const wxPoint2DDouble& p = points[i];
        indices[0] = i;

        indices[1] = 0;
        wxCHECK(sa.SetElement(indices, p.m_x), shape);
        indices[1] = 1;
        wxCHECK(sa.SetElement(indices, p.m_y), shape);
    }

    wxVariant vPoints(new wxVariantDataSafeArray(sa.Detach()));

    WXAUTOEXCEL_CALL_METHOD1_OBJECT("AddPolyline", vPoints, shape);
}

wxExcelShape wxExcelShapes::AddShape(MsoAutoShapeType type, double left, double top,
                                     double width, double height)
{
    wxExcelShape shape;

    WXAUTOEXCEL_CALL_METHOD5("AddShape", (long)type, left, top, width, height, "void*", shape);
    VariantToObject(vResult, &shape);
    return shape;
}

wxExcelShape wxExcelShapes::AddTextbox(MsoTextOrientation orientation, double left, double top,
                                        double width, double height)
{
    wxExcelShape shape;

    WXAUTOEXCEL_CALL_METHOD5("AddTextbox", (long)orientation, left, top, width, height, "void*", shape);
    VariantToObject(vResult, &shape);
    return shape;
}

wxExcelShape wxExcelShapes::AddTextEffect(MsoPresetTextEffect presetTextEffect, const wxString& text,
                                          const wxString& fontName, double fontSize,
                                          MsoTriState fontBold, MsoTriState fontItalic,
                                          double left, double top)
{
    wxVariantVector args;

    args.push_back((long)presetTextEffect);
    args.push_back(text);
    args.push_back(fontName);
    args.push_back(fontSize);
    args.push_back((long)fontBold);
    args.push_back((long)fontItalic);
    args.push_back(left);
    args.push_back(top);

    wxExcelShape shape;

    WXAUTOEXCEL_CALL_METHODARR("AddTextEffect", args, "void*", shape);
    VariantToObject(vResult, &shape);
    return shape;
}

wxExcelFreeformBuilder wxExcelShapes::BuildFreeform(MsoEditingType editingType, double X1, double Y1)
{
    wxExcelFreeformBuilder builder;

    WXAUTOEXCEL_CALL_METHOD3("BuildFreeform", (long)editingType, X1, Y1, "void*", builder);
    VariantToObject(vResult, &builder);
    return builder;
}

wxExcelShape wxExcelShapes::Item(long index)
{
    wxASSERT( index > 0 );

    wxExcelShape shape;
    WXAUTOEXCEL_CALL_METHOD1_OBJECT("Item", index, shape);
}

wxExcelShape wxExcelShapes::operator[](long index)
{
    return Item(index);
}

wxExcelShape wxExcelShapes::Item(const wxString& name)
{
    wxExcelShape shape;

    WXAUTOEXCEL_CALL_METHOD1_OBJECT("Item", name, shape);
}

wxExcelShape wxExcelShapes::operator[](const wxString& name)
{
    return Item(name);
}

void wxExcelShapes::SelectAll()
{
    WXAUTOEXCEL_CALL_METHOD0_RET("SelectAll", "null");
}

// ***** class wxExcelShapes PROPERTIES *****


long wxExcelShapes::GetCount()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("Count");
}


wxExcelShapeRange wxExcelShapes::GetRange(long index)
{
    wxExcelShapeRange shapeRange;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET1("Range", index, shapeRange);
}

wxExcelShapeRange wxExcelShapes::GetRange(const wxString& name)
{
    wxExcelShapeRange shapeRange;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET1("Range", name, shapeRange);
}

wxExcelShapeRange wxExcelShapes::GetRange(const wxVector<long>& indices)
{
    wxExcelShapeRange shapeRange;

    wxCHECK(indices.size() > 0, shapeRange);

    wxVariant vIndices;

    vIndices.NullList();
    for (size_t i = 0; i < indices.size(); i++)
        vIndices.Append(indices[i]);

    WXAUTOEXCEL_PROPERTY_OBJECT_GET1("Range", vIndices, shapeRange);
}

wxExcelShapeRange wxExcelShapes::GetRange(const wxVector<wxString>& names)
{
    wxExcelShapeRange shapeRange;

    wxCHECK(names.size() > 0, shapeRange);

    wxVariant vNames;

    vNames.NullList();
    for (size_t i = 0; i < names.size(); i++)
        vNames.Append(names[i]);

    WXAUTOEXCEL_PROPERTY_OBJECT_GET1("Range", vNames, shapeRange);
}

} // namespace wxAutoExcel

#endif // #if WXAUTOEXCEL_USE_SHAPES
