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

        [MSDN documentation for ChartCategory.AutoScaleFont](https://msdn.microsoft.com/vba/excel-vba/articles/chartcategory-isfiltered-property-excel).
        */
        bool GetIsFiltered();

        /**
        Returns a String value that represents the name of the object.

        [MSDN documentation for ChartCategory.Name](https://msdn.microsoft.com/vba/excel-vba/articles/chartcategory-name-property-excel).
        */
        wxString GetName();

        /**
        Sets a String value that represents the name of the object.

        [MSDN documentation for ChartCategory.Name](https://msdn.microsoft.com/vba/excel-vba/articles/chartcategory-name-property-excel).
        */
        void SetName(const wxString& name);

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

        [MSDN documentation for CategoryCollection.Item](http://msdn.microsoft.com/en-us/library/bb211737).
        */
        wxExcelChartCategory Item(long index);
        wxExcelChartCategory operator[](long index);
        //@}

        // ***** PROPERTIES *****

        /**
        Returns a Long value that represents the number of objects in the collection.

        [MSDN documentation for CategoryCollection.Count](http://msdn.microsoft.com/en-us/library/bb179622).
        */
        long GetCount();

        /**
        Returns "CategoryCollection".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("CategoryCollection"); }
    };


} // namespace wxAutoExcel

#endif // #if WXAUTOEXCEL_USE_CHARTS

#endif // _WXAUTOEXCEL_CHARTCATEGORY_H
