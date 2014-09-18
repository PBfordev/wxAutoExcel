/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Modified by:
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////


#ifndef _WXAUTOEXCEL_CHARTVIEW_H
#define _WXAUTOEXCEL_CHARTVIEW_H

#include "wx/wxAutoExcel_defs.h"

#if WXAUTOEXCEL_USE_CHARTS

#include "wx/wxAutoExcelObject.h"

namespace wxAutoExcel {

    /**
    Represents Microsoft Excel ChartView object.
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelChartView : public wxExcelObject
    {
    public:

        // ***** PROPERTIES *****

        /**
        Returns the sheet name for the specified ChartView object. Since Excel 2007.

        [MSDN documentation for ChartView.Sheet](http://msdn.microsoft.com/en-us/library/bb224263).
        */
        wxExcelSheet GetSheet();

        /**
        Returns "ChartView".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("ChartView"); }
    };


} // namespace wxAutoExcel

#endif // #if WXAUTOEXCEL_USE_CHARTS

#endif // _WXAUTOEXCEL_CHARTVIEW_H
