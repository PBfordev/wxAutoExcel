/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
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

        [MSDN documentation for AutoFilter.ApplyFilter](http://msdn.microsoft.com/en-us/library/bb238808).
        */
        void ApplyFilter();

        /**
        Displays all the data returned by the AutoFilter object.

        [MSDN documentation for AutoFilter.ShowAllData](http://msdn.microsoft.com/en-us/library/bb238813).
        */
        void ShowAllData();

        // ***** PROPERTIES *****

        /**
        Returns True if the worksheet is in the AutoFilter filter mode. Since Excel 2007.

        [MSDN documentation for AutoFilter.FilterMode](http://msdn.microsoft.com/en-us/library/bb179270).
        */
        bool GetFilterMode();

        /**
        Returns a Filters collection that represents all the filters in an autofiltered range.

        [MSDN documentation for AutoFilter.Filters](http://msdn.microsoft.com/en-us/library/bb208508).
        */
        wxExcelFilters GetFilters();

        /**
        Returns a Range Represents the range to which the specified AutoFilter applies.

        [MSDN documentation for AutoFilter.Range](http://msdn.microsoft.com/en-us/library/bb179272).
        */
        wxExcelRange GetRange();

        /**
        Gets the sort column or columns, and sort order for the AutoFilter collection. Since Excel 2007.

        [MSDN documentation for AutoFilter.Sort](http://msdn.microsoft.com/en-us/library/bb214465).
        */
        wxExcelSort GetSort();

        /**
        Returns "AutoFilter".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("AutoFilter"); }
    };


} // namespace wxAutoExcel

#endif //_WXAUTOEXCEL_AUTOFILTER_H
