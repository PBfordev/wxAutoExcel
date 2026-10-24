/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// License:     MIT license
/////////////////////////////////////////////////////////////////////////////


#ifndef _WXAUTOEXCEL_CHARTCATEGORY_H
#define _WXAUTOEXCEL_CHARTCATEGORY_H

#include "wx/wxAutoExcel_defs.h"

#if WXAUTOEXCEL_USE_CHARTS

#include "wx/wxAutoExcel_object.h"
#include "wx/wxAutoExcel_enums.h"

namespace wxAutoExcel {

    /**
    Specifies a chart type category.
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelChartCategory : public wxExcelObject
    {
    public:

        // ***** PROPERTIES *****

        /**
        True if the user filtered out a series.

        [MSDN documentation for ChartCategory.IsFiltered](https://learn.microsoft.com/en-us/office/vba/api/excel.chartcategory.isfiltered).
        */
        bool GetIsFiltered();

        /**
        Returns a String value that represents the name of the object.

        [MSDN documentation for ChartCategory.Name](https://learn.microsoft.com/en-us/office/vba/api/excel.chartcategory.name).
        */
        wxString GetName();

        /**
        Returns "ChartCategory".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("ChartCategory"); }
    };

    /**
    Represents the collection of visible chart categories in the chart.
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelCategoryCollection : public wxExcelObject
    {
    public:
        // ***** METHODS *****

        //@{
        /**
        Returns a single object from a collection.

        [MSDN documentation for CategoryCollection.Item](https://learn.microsoft.com/en-us/office/vba/api/excel.categorycollection.item).
        */
        wxExcelChartCategory Item(long index);
        wxExcelChartCategory operator[](long index);
        //@}

        // ***** PROPERTIES *****

        /**
        Returns a Long value that represents the number of objects in the collection.

        [MSDN documentation for CategoryCollection.Count](https://learn.microsoft.com/en-us/office/vba/api/excel.categorycollection.count).
        */
        long GetCount();

        /**
        The iterator over the items, see wxExcelCollectionIterator.
        */
        typedef wxExcelCollectionIterator<wxExcelCategoryCollection, wxExcelChartCategory> iterator;

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
        Returns "CategoryCollection".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("CategoryCollection"); }
    };


} // namespace wxAutoExcel

#endif // #if WXAUTOEXCEL_USE_CHARTS

#endif // _WXAUTOEXCEL_CHARTCATEGORY_H
