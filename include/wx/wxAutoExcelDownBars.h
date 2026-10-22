/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// License:     MIT license
/////////////////////////////////////////////////////////////////////////////


#ifndef _WXAUTOEXCEL_DOWNBARS_H
#define _WXAUTOEXCEL_DOWNBARS_H

#include "wx/wxAutoExcel_defs.h"

#if WXAUTOEXCEL_USE_CHARTS

#include "wx/wxAutoExcel_object.h"

namespace wxAutoExcel {

    /**
    @brief Represents Microsoft Excel DownBars object.
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelDownBars : public wxExcelObject
    {
    public:
        // ***** METHODS *****

        /**
        Deletes the object.

        [MSDN documentation for DownBars.Delete](https://learn.microsoft.com/en-us/office/vba/api/excel.downbars.delete).
        */
        bool Delete();

        /**
        Selects the object.

        [MSDN documentation for DownBars.Select](https://learn.microsoft.com/en-us/office/vba/api/excel.downbars.select).
        */
        bool Select();

        // ***** PROPERTIES *****

        /**
        Returns the ChartFormat object. Since Excel 2007.

        [MSDN documentation for DownBars.Format](https://learn.microsoft.com/en-us/office/vba/api/excel.downbars.format).
        */
        wxExcelChartFormat GetFormat();

        /**
        Returns a String value that represents the name of the object.

        [MSDN documentation for DownBars.Name](https://learn.microsoft.com/en-us/office/vba/api/excel.downbars.name).
        */
        wxString GetName();

        /**
        Returns "DownBars".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("DownBars"); }
    };


} // namespace wxAutoExcel

#endif // #if WXAUTOEXCEL_USE_CHARTS

#endif // _WXAUTOEXCEL_DOWNBARS_H
