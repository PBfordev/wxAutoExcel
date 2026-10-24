/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// License:     MIT license
/////////////////////////////////////////////////////////////////////////////

#ifndef _WXAUTOEXCEL_ADJUSTMENTS_H
#define _WXAUTOEXCEL_ADJUSTMENTS_H

#include "wx/wxAutoExcel_defs.h"

#if WXAUTOEXCEL_USE_SHAPES

#include "wx/wxAutoExcel_object.h"

namespace wxAutoExcel {

    /**
    @brief Represents Microsoft Excel Adjustments object.
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelAdjustments : public wxExcelObject
    {
    public:

        // ***** PROPERTIES *****

        /**
        Returns the number of objects in the collection.

        [MSDN documentation for Adjustments.Count](https://learn.microsoft.com/en-us/office/vba/api/excel.adjustments.count).
        */
        long GetCount();

        //@{
        /**
        Returns an item from the collection.

        [MSDN documentation for Adjustments.Item](https://learn.microsoft.com/en-us/office/vba/api/excel.adjustments.item).
        */
        double GetItem(long index);

        /**
        Sets the adjustment value specified by the index.

        [MSDN documentation for Adjustments.Item](https://learn.microsoft.com/en-us/office/vba/api/excel.adjustments.item).
        */
        void SetItem(long index, double value);
        double operator[](long index);
        //@}

        /**
        The iterator over the items, see wxExcelCollectionIterator.
        */
        typedef wxExcelCollectionIterator<wxExcelAdjustments, double> iterator;

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
        Returns "Adjustments".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("Adjustments"); }
    };


} // namespace wxAutoExcel

#endif // #if WXAUTOEXCEL_USE_SHAPES

#endif //_WXAUTOEXCEL_ADJUSTMENTS_H
