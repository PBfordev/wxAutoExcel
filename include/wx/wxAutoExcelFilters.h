/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// License:     MIT license
/////////////////////////////////////////////////////////////////////////////


#ifndef _WXAUTOEXCEL_FILTERS_H
#define _WXAUTOEXCEL_FILTERS_H

#include "wx/wxAutoExcel_defs.h"

#include "wx/wxAutoExcel_object.h"
#include "wx/wxAutoExcel_enums.h"

namespace wxAutoExcel {

    /**
    @brief Represents Microsoft Excel Filter object.
    */
    class WXDLLIMPEXP_WXAUTOEXCEL wxExcelFilter : public wxExcelObject
    {
    public:

        /**
        Read-only Since Excel 2007.

        [MSDN documentation for Filter.Count](http://msdn.microsoft.com/en-us/library/bb213162).
        */
        long GetCount();

        /**
        Returns the first filtered value for the specified column in a filtered range.

        [MSDN documentation for Filter.Criteria1](http://msdn.microsoft.com/en-us/library/bb177402).
        */
        wxVariant GetCriteria1();

        /**
        Returns the second filtered value for the specified column in a filtered range.

        [MSDN documentation for Filter.Criteria2](http://msdn.microsoft.com/en-us/library/bb177404).
        */
        wxVariant GetCriteria2();

        /**
        True if the specified filter is on.

        [MSDN documentation for Filter.On](http://msdn.microsoft.com/en-us/library/bb208878).
        */
        bool GetOn();

        /**
        Returns an XlAutoFilterOperator value that represents the operator that associates the two criteria applied by the specified filter.

        [MSDN documentation for Filter.Operator](http://msdn.microsoft.com/en-us/library/bb213164).
        */
        XlAutoFilterOperator  GetOperator();

        /**
        Returns "Filter".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("Filter"); }
    };

    /**
    @brief Represents Microsoft Excel Filters collection.
    */
    class WXDLLIMPEXP_WXAUTOEXCEL wxExcelFilters : public wxExcelObject
    {
    public:
        /**
        Returns a Long value that represents the number of objects in the collection.

        [MSDN documentation for Filters.Count](http://msdn.microsoft.com/en-us/library/bb213170).
        */
        long GetCount();

        //@{
        /**
        Returns a single object from a collection.

        [MSDN documentation for Filters.Item](http://msdn.microsoft.com/en-us/library/bb213172).
        */
        wxExcelFilter GetItem(long index);
        wxExcelFilter operator[](long index);
        //@}

        /**
        Returns "Filters".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("Filters"); }
    };


} // namespace wxAutoExcel

#endif //_WXAUTOEXCEL_FILTERS_H
