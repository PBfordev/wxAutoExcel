/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Modified by:
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////


#include "wx/wxAutoExcel_prec.h"

#include "wx/wxAutoExcelGroupShapes.h"

#if WXAUTOEXCEL_USE_SHAPES

#include <wx/vector.h>

#include "wx/wxAutoExcelShape.h"
#include "wx/wxAutoExcelShapeRange.h"

#include "wx/wxAutoExcel_private.h"

namespace wxAutoExcel {

// ***** class wxExcelGroupShapes METHODS *****

wxExcelShape wxExcelGroupShapes::Item(long index)
{
    wxASSERT( index > 0 );

    wxExcelShape shape;
    WXAUTOEXCEL_CALL_METHOD1_OBJECT("Item", index, shape);
}

wxExcelShape wxExcelGroupShapes::operator[](long index)
{
    return Item(index);
}

// ***** class wxExcelGroupShapes PROPERTIES *****


long wxExcelGroupShapes::GetCount()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("Count");
}


wxExcelShapeRange wxExcelGroupShapes::GetRange(long index)
{
    wxExcelShapeRange shapeRange;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET1("Range", index, shapeRange);
}

wxExcelShapeRange wxExcelGroupShapes::GetRange(const wxString& name)
{
    wxExcelShapeRange shapeRange;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET1("Range", name, shapeRange);
}

wxExcelShapeRange wxExcelGroupShapes::GetRange(const wxVector<long>& indices)
{
    wxExcelShapeRange shapeRange;

    wxCHECK(indices.size() > 0, shapeRange);

    wxVariant vIndices;

    vIndices.NullList();
    for (size_t i = 0; i < indices.size(); i++)
        vIndices.Append(indices[i]);

    WXAUTOEXCEL_PROPERTY_OBJECT_GET1("Range", vIndices, shapeRange);
}

} // namespace wxAutoExcel

#endif // #if WXAUTOEXCEL_USE_SHAPES
