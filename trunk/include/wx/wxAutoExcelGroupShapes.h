/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Modified by:
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////


#ifndef _WXAUTOEXCEL_GROUPSHAPES_H
#define _WXAUTOEXCEL_GROUPSHAPES_H

#include "wx/wxAutoExcel_defs.h"

#if WXAUTOEXCEL_USE_SHAPES

#include "wx/wxAutoExcel_object.h"

namespace wxAutoExcel {

    /**
    @brief Represents Microsoft Excel GroupShapes collection.
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelGroupShapes : public wxExcelObject
    {
    public:        
        // ***** METHODS *****

        //@{
        /**
        Returns a single object from a collection.

        [MSDN documentation for GroupShapes.Item](http://msdn.microsoft.com/en-us/library/bb211791).
        */
        wxExcelShape Item(long index);
        wxExcelShape operator[](long index);
        //@}

        // ***** PROPERTIES *****        

        /**
        Returns a Long value that represents the number of objects in the collection.

        [MSDN documentation for GroupShapes.Count](http://msdn.microsoft.com/en-us/library/bb148503).
        */
        long GetCount();

        //@{
        /**
        Returns a ShapeRange Represents a subset of the shapes in a Shapes collection.

        [MSDN documentation for GroupShapes.Range](http://msdn.microsoft.com/en-us/library/bb148505).
        */
        wxExcelShapeRange GetRange(long index);
        wxExcelShapeRange GetRange(const wxString& name);
        wxExcelShapeRange GetRange(const wxVector<long>& indices);

        //@}

        /**
        Returns "GroupShapes".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("GroupShapes"); }
    };


} // namespace wxAutoExcel

#endif // #if WXAUTOEXCEL_USE_SHAPES

#endif //_WXAUTOEXCEL_GROUPSHAPES_H
