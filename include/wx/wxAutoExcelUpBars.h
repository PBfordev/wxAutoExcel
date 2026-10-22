/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// License:     MIT license
/////////////////////////////////////////////////////////////////////////////


#ifndef _WXAUTOEXCEL_UPBARS_H
#define _WXAUTOEXCEL_UPBARS_H

#include "wx/wxAutoExcel_defs.h"

#if WXAUTOEXCEL_USE_CHARTS

#include "wx/wxAutoExcel_object.h"

namespace wxAutoExcel {

    /**
    @brief Represents Microsoft Excel UpBars object.
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelUpBars : public wxExcelObject
    {
    public:
        // ***** METHODS *****

        /**
        bool the object.

        [MSDN documentation for UpBars.Delete](https://learn.microsoft.com/en-us/office/vba/api/excel.upbars.delete).
        */
        bool Delete();

        /**
        Selects the object.

        [MSDN documentation for UpBars.Select](https://learn.microsoft.com/en-us/office/vba/api/excel.upbars.select).
        */
        bool Select();

        // ***** PROPERTIES *****

        /**
        Returns the ChartFormat object.  Since Excel 2007.

        [MSDN documentation for UpBars.Format](https://learn.microsoft.com/en-us/office/vba/api/excel.upbars.format).
        */
        wxExcelChartFormat GetFormat();

        /**
        Returns a String value that represents the name of the object.

        [MSDN documentation for UpBars.Name](https://learn.microsoft.com/en-us/office/vba/api/excel.upbars.name).
        */
        wxString GetName();


        /**
        Returns "UpBars".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("UpBars"); }
    };


} // namespace wxAutoExcel

#endif // #if WXAUTOEXCEL_USE_CHARTS

#endif // _WXAUTOEXCEL_UPBARS_H
