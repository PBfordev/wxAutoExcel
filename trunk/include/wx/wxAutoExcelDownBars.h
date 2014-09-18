/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Modified by:
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////


#ifndef _WXAUTOEXCEL_DOWNBARS_H
#define _WXAUTOEXCEL_DOWNBARS_H

#include "wx/wxAutoExcel_defs.h"

#if WXAUTOEXCEL_USE_CHARTS

#include "wx/wxAutoExcelObject.h"

namespace wxAutoExcel {

    /**
    Represents Microsoft Excel DownBars object.
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelDownBars : public wxExcelObject
    {
    public:
        // ***** METHODS *****

        /**
        Deletes the object.

        [MSDN documentation for DownBars.Delete](http://msdn.microsoft.com/en-us/library/bb211747).
        */
        bool Delete();

        /**
        Selects the object.

        [MSDN documentation for DownBars.Select](http://msdn.microsoft.com/en-us/library/bb237875).
        */
        bool Select();

        // ***** PROPERTIES *****

        /**
        Returns the ChartFormat object. Since Excel 2007.

        [MSDN documentation for DownBars.Format](http://msdn.microsoft.com/en-us/library/bb242526).
        */
        wxExcelChartFormat GetFormat();

        /**
        Returns a String value that represents the name of the object.

        [MSDN documentation for DownBars.Name](http://msdn.microsoft.com/en-us/library/bb236956).
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
