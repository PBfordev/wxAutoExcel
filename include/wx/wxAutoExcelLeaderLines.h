/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////


#ifndef _WXAUTOEXCEL_LEADERLINES_H
#define _WXAUTOEXCEL_LEADERLINES_H

#include "wx/wxAutoExcel_defs.h"

#if WXAUTOEXCEL_USE_CHARTS

#include "wx/wxAutoExcel_object.h"

namespace wxAutoExcel {

    /**
    @brief Represents Microsoft Excel LeaderLines object.
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelLeaderLines : public wxExcelObject
    {
    public:

        // ***** METHODS *****

        /**
        Deletes the object.

        [MSDN documentation for LeaderLines.Delete](http://msdn.microsoft.com/en-us/library/bb211827).
        */
        bool Delete();

        /**
        Selects the object.

        [MSDN documentation for LeaderLines.Select](http://msdn.microsoft.com/en-us/library/bb237959).
        */
        bool Select();

        // ***** PROPERTIES *****


        /**
        Returns a Border object that represents the border of the object.

        [MSDN documentation for LeaderLines.Border](http://msdn.microsoft.com/en-us/library/bb148533).
        */
        wxExcelBorder GetBorder();

        /**
        Returns the ChartFormat object. Since Excel 2007.

        [MSDN documentation for LeaderLines.Format](http://msdn.microsoft.com/en-us/library/bb242532).
        */
        wxExcelChartFormat GetFormat();

        /**
        Returns "LeaderLines".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("LeaderLines"); }
    };


} // namespace wxAutoExcel

#endif // #if WXAUTOEXCEL_USE_CHARTS

#endif // _WXAUTOEXCEL_LEADERLINES_H
