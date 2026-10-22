/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// License:     MIT license
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

        [MSDN documentation for LeaderLines.Delete](https://learn.microsoft.com/en-us/office/vba/api/excel.leaderlines.delete).
        */
        bool Delete();

        /**
        Selects the object.

        [MSDN documentation for LeaderLines.Select](https://learn.microsoft.com/en-us/office/vba/api/excel.leaderlines.select).
        */
        bool Select();

        // ***** PROPERTIES *****


        /**
        Returns a Border object that represents the border of the object.

        [MSDN documentation for LeaderLines.Border](https://learn.microsoft.com/en-us/office/vba/api/excel.leaderlines.border).
        */
        wxExcelBorder GetBorder();

        /**
        Returns the ChartFormat object. Since Excel 2007.

        [MSDN documentation for LeaderLines.Format](https://learn.microsoft.com/en-us/office/vba/api/excel.leaderlines.format).
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
