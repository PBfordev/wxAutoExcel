/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// License:     MIT license
/////////////////////////////////////////////////////////////////////////////


#ifndef _WXAUTOEXCEL_SERIESLINES_H
#define _WXAUTOEXCEL_SERIESLINES_H

#include "wx/wxAutoExcel_defs.h"

#if WXAUTOEXCEL_USE_CHARTS

#include "wx/wxAutoExcel_object.h"

namespace wxAutoExcel {

    /**
    @brief Represents Microsoft Excel SeriesLines object.
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelSeriesLines : public wxExcelObject
    {
    public:
        // ***** METHODS *****

        /**
        Deletes the object.

        [MSDN documentation for SeriesLines.Delete](https://learn.microsoft.com/en-us/office/vba/api/excel.serieslines.delete).
        */
        bool Delete();

        /**
        Selects the object.

        [MSDN documentation for SeriesLines.Select](https://learn.microsoft.com/en-us/office/vba/api/excel.serieslines.select).
        */
        bool Select();

        // ***** PROPERTIES *****

        /**
        Returns a Border object that represents the border of the object.

        [MSDN documentation for SeriesLines.Border](https://learn.microsoft.com/en-us/office/vba/api/excel.serieslines.border).
        */
        wxExcelBorder GetBorder();

        /**
        Returns the ChartFormat object.  Since Excel 2007.

        [MSDN documentation for SeriesLines.Format](https://learn.microsoft.com/en-us/office/vba/api/excel.serieslines.format).
        */
        wxExcelChartFormat GetFormat();

        /**
        Returns a String value that represents the name of the object.

        [MSDN documentation for SeriesLines.Name](https://learn.microsoft.com/en-us/office/vba/api/excel.serieslines.name).
        */
        wxString GetName();

        /**
        Returns "SeriesLines".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("SeriesLines"); }
    };


} // namespace wxAutoExcel

#endif // #if WXAUTOEXCEL_USE_CHARTS

#endif // _WXAUTOEXCEL_SERIESLINES_H
