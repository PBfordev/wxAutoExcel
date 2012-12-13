/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Modified by:
// RCS-ID:      $Id: $
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////


#include "wx/wxAutoExcel_prec.h"

#include "wx/wxAutoExcelFreeformBuilder.h"

#if WXAUTOEXCEL_USE_SHAPES

#include "wx/wxAutoExcelShape.h"

#include "wx/wxAutoExcelPrivate.h"

namespace wxAutoExcel {

    // ***** class wxExcelFreeformBuilder METHODS *****

void wxExcelFreeformBuilder::AddNodes(MsoSegmentType segmentType, MsoEditingType editingType,
                                      double X1, double Y1, double* X2, double* Y2, double* X3, double* Y3)
{
    wxVariantVector args;
        
    args.push_back((long)segmentType);
    args.push_back((long)editingType);
    args.push_back(X1);
    args.push_back(Y1);
    
    WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT_VECTOR(X2, args);
    WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT_VECTOR(Y2, args);
    WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT_VECTOR(X3, args);
    WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT_VECTOR(Y3, args);

    WXAUTOEXCEL_CALL_METHODARR_RET("AddNodes", args, "null");
}

wxExcelShape wxExcelFreeformBuilder::ConvertToShape()
{        
    wxExcelShape shape;

    WXAUTOEXCEL_CALL_METHOD0_OBJECT("ConvertToShape", shape);    
}

} // namespace wxAutoExcel

#endif // #if WXAUTOEXCEL_USE_SHAPES
