/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Modified by:
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////


#ifndef _WXAUTOEXCEL_FREEFORMBUILDER_H
#define _WXAUTOEXCEL_FREEFORMBUILDER_H

#include "wx/wxAutoExcel_defs.h"

#if WXAUTOEXCEL_USE_SHAPES

#include "wx/wxAutoExcel_object.h"
#include "wx/wxAutoExcel_enums.h"

namespace wxAutoExcel {

    /**
    @brief Represents Microsoft Excel FreeformBuilder object.
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelFreeformBuilder : public wxExcelObject
    {
    public:        
        // ***** METHODS *****

        /**
        Adds a point in the current shape and then draws a line from the current node to last node that was added.

        [MSDN documentation for FreeformBuilder.AddNodes](http://msdn.microsoft.com/en-us/library/bb209588).
        */
        void AddNodes(MsoSegmentType segmentType, MsoEditingType editingType,
                      double X1, double Y1, double* X2 = NULL, double* Y2 = NULL, double* X3 = NULL, double* Y3 = NULL);

        /**
        Creates a shape that has the geometric characteristics of the specified FreeformBuilder object. Returns a Shape Represents the new shape.

        [MSDN documentation for FreeformBuilder.ConvertToShape](http://msdn.microsoft.com/en-us/library/bb223284).
        */
        wxExcelShape ConvertToShape();


        /**
        Returns "FreeformBuilder".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("FreeformBuilder"); }
    };


} // namespace wxAutoExcel

#endif // #if WXAUTOEXCEL_USE_SHAPES

#endif //_WXAUTOEXCEL_FREEFORMBUILDER_H
