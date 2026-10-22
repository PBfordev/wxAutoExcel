/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// License:     MIT license
/////////////////////////////////////////////////////////////////////////////


#ifndef _WXAUTOEXCEL_AUTOFILTER_H
#define _WXAUTOEXCEL_AUTOFILTER_H

#include "wx/wxAutoExcel_defs.h"
#include "wx/wxAutoExcel_object.h"

namespace wxAutoExcel {

    /**
    @brief Represents Microsoft Excel AutoFilter object.
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelAutoFilter : public wxExcelObject
    {
    public:
        // ***** METHODS *****

        /**
        Applies the specified Autofilter object.

        [MSDN documentation for AutoFilter.ApplyFilter](https://learn.microsoft.com/en-us/office/vba/api/excel.autofilter.applyfilter).
        */
        void ApplyFilter();

        /**
        Displays all the data returned by the AutoFilter object.

        [MSDN documentation for AutoFilter.ShowAllData](https://learn.microsoft.com/en-us/office/vba/api/excel.autofilter.showalldata).
        */
        void ShowAllData();

        // ***** PROPERTIES *****

        /**
        Returns True if the worksheet is in the AutoFilter filter mode. Since Excel 2007.

        [MSDN documentation for AutoFilter.FilterMode](https://learn.microsoft.com/en-us/office/vba/api/excel.autofilter.filtermode).
        */
        bool GetFilterMode();

        /**
        Returns a Filters collection that represents all the filters in an autofiltered range.

        [MSDN documentation for AutoFilter.Filters](https://learn.microsoft.com/en-us/office/vba/api/excel.autofilter.filters).
        */
        wxExcelFilters GetFilters();

        /**
        Returns a Range Represents the range to which the specified AutoFilter applies.

        [MSDN documentation for AutoFilter.Range](https://learn.microsoft.com/en-us/office/vba/api/excel.autofilter.range).
        */
        wxExcelRange GetRange();

        /**
        Gets the sort column or columns, and sort order for the AutoFilter collection. Since Excel 2007.

        [MSDN documentation for AutoFilter.Sort](https://learn.microsoft.com/en-us/office/vba/api/excel.autofilter.sort).
        */
        wxExcelSort GetSort();

        /**
        Returns "AutoFilter".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("AutoFilter"); }
    };


} // namespace wxAutoExcel

#endif //_WXAUTOEXCEL_AUTOFILTER_H
