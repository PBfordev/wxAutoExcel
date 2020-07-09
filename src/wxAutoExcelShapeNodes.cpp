/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Modified by:
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////


#include "wx/wxAutoExcel_prec.h"

#include "wx/wxAutoExcelShapeNodes.h"

#if WXAUTOEXCEL_USE_SHAPES

#include <wx/vector.h>
#include <wx/geometry.h>

#include "wx/wxAutoExcelShape.h"

#include "wx/wxAutoExcel_private.h"

namespace wxAutoExcel {

// ***** class wxExcelShapeNode PROPERTIES *****


MsoEditingType wxExcelShapeNode::GetEditingType()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("EditingType", MsoEditingType, msoEditingAuto);
}


wxVector<wxPoint2DDouble> wxExcelShapeNode::GetPoints()
{
    wxVariant vResult;
    wxVector<wxPoint2DDouble> points;

    if ( InvokeGetProperty(wxS("Points"), vResult) )
    {
        if ( vResult.GetType() == wxS("list") && vResult.GetCount() % 2 == 0 )
        {
            size_t pointCount = vResult.GetCount() / 2;
            for ( size_t i = 0; i < pointCount; i++ )
            {
                wxPoint2DDouble point;
                point.m_x = vResult[i];
                point.m_y = vResult[i+pointCount]; // we get the array with all x coordinates and then all y coordinates
                points.push_back(point);
            }
        }
    }
    return points;
}

MsoSegmentType wxExcelShapeNode::GetSegmentType()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("SegmentType", MsoSegmentType, msoSegmentLine);
}


// ***** class wxExcelShapeNodes METHODS *****

void wxExcelShapeNodes::Delete(long index)
{
    WXAUTOEXCEL_CALL_METHOD1_RET("Delete", index, "null");
}

void wxExcelShapeNodes::Insert(long index, MsoSegmentType segmentType, MsoEditingType editingType,
                               double X1, double Y1, double* X2, double* Y2, double* X3, double* Y3)
{
    wxVariantVector args;

    args.push_back(index);
    args.push_back((long)segmentType);
    args.push_back((long)editingType);
    args.push_back(X1);
    args.push_back(Y1);

    WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT_VECTOR(X2, args);
    WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT_VECTOR(Y2, args);
    WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT_VECTOR(X3, args);
    WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT_VECTOR(Y3, args);

    WXAUTOEXCEL_CALL_METHODARR_RET("Insert", args, "null");
}

wxExcelShapeNode wxExcelShapeNodes::Item(long index)
{
    wxASSERT( index > 0 );

    wxExcelShapeNode node;
    WXAUTOEXCEL_CALL_METHOD1_OBJECT("Item", index, node);
}

wxExcelShapeNode wxExcelShapeNodes::operator[](long index)
{
    return Item(index);
}

void wxExcelShapeNodes::SetEditingType(long index, MsoEditingType editingType)
{
    WXAUTOEXCEL_CALL_METHOD2_RET("SetEditingType", index, (long)editingType, "null");
}

void wxExcelShapeNodes::SetPosition(long index, double X1, double Y1)
{
    WXAUTOEXCEL_CALL_METHOD3_RET("SetPosition", index, X1, Y1, "null");
}

void wxExcelShapeNodes::SetSegmentType(long index, MsoSegmentType segmentType)
{
    WXAUTOEXCEL_CALL_METHOD2_RET("SetSegmentType", index, (long)segmentType, "null");
}

// ***** class wxExcelShapeNodes PROPERTIES *****

long wxExcelShapeNodes::GetCount()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("Count");
}

} // namespace wxAutoExcel

#endif // #include "wx/wxAutoExcelShapeNodes.h"