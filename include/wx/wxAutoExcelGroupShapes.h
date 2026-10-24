/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// License:     MIT license
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

        [MSDN documentation for GroupShapes.Item](https://learn.microsoft.com/en-us/office/vba/api/excel.groupshapes.item).
        */
        wxExcelShape Item(long index);
        wxExcelShape operator[](long index);
        //@}

        // ***** PROPERTIES *****

        /**
        Returns a Long value that represents the number of objects in the collection.

        [MSDN documentation for GroupShapes.Count](https://learn.microsoft.com/en-us/office/vba/api/excel.groupshapes.count).
        */
        long GetCount();

        //@{
        /**
        Returns a ShapeRange Represents a subset of the shapes in a Shapes collection.

        [MSDN documentation for GroupShapes.Range](https://learn.microsoft.com/en-us/office/vba/api/excel.groupshapes.range).
        */
        wxExcelShapeRange GetRange(long index);
        wxExcelShapeRange GetRange(const wxString& name);
        wxExcelShapeRange GetRange(const wxVector<long>& indices);

        //@}

        /**
        The iterator over the items, see wxExcelCollectionIterator.
        */
        typedef wxExcelCollectionIterator<wxExcelGroupShapes, wxExcelShape> iterator;

        //@{
        /**
        Returns the iterator referring to the first item, or past the last item,
        which allows iterating over the items with a range-based for loop,
        see wxExcelCollectionIterator.
        */
        iterator begin() { return iterator(this, 1); }
        iterator end() { return iterator(this, GetCount() + 1); }
        //@}

        /**
        Returns "GroupShapes".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("GroupShapes"); }
    };


} // namespace wxAutoExcel

#endif // #if WXAUTOEXCEL_USE_SHAPES

#endif //_WXAUTOEXCEL_GROUPSHAPES_H
