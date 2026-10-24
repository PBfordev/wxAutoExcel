/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// License:     MIT license
/////////////////////////////////////////////////////////////////////////////


#ifndef _WXAUTOEXCEL_AREAS_H
#define _WXAUTOEXCEL_AREAS_H

#include "wx/wxAutoExcel_defs.h"
#include "wx/wxAutoExcel_object.h"

namespace wxAutoExcel {

    /**
    @brief Represents Microsoft Excel Areas collection.
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelAreas: public wxExcelObject
    {
    public:

        // ***** PROPERTIES *****

        /**
        Returns the number of ranges in the collection.

        [MSDN documentation for Areas.Count](https://learn.microsoft.com/en-us/office/vba/api/excel.areas.count).
        */
        long GetCount();

        //@{
        /**
        Returns a range from a collection.

        [MSDN documentation for Worksheets.Item](https://learn.microsoft.com/en-us/office/vba/api/excel.worksheets.item).
        */
        wxExcelRange GetItem(long index);
        wxExcelRange operator[](long index);
        //@}

        /**
        The iterator over the items, see wxExcelCollectionIterator.
        */
        typedef wxExcelCollectionIterator<wxExcelAreas, wxExcelRange> iterator;

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
        Returns "Areas".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("Areas"); }
    };

} // namespace wxAutoExcel

#endif //_WXAUTOEXCEL_AREAS_H
