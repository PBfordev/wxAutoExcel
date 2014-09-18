/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Modified by:
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////


#ifndef _WXAUTOEXCEL_SERIESLINES_H
#define _WXAUTOEXCEL_SERIESLINES_H

#include "wx/wxAutoExcel_defs.h"

#if WXAUTOEXCEL_USE_CHARTS

#include "wx/wxAutoExcelObject.h"

namespace wxAutoExcel {

    /**
    Represents Microsoft Excel SeriesLines object.
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelSeriesLines : public wxExcelObject
    {
    public:
        // ***** METHODS *****

        /**
        Deletes the object.

        [MSDN documentation for SeriesLines.Delete](http://msdn.microsoft.com/en-us/library/bb178933).
        */
        bool Delete();

        /**
        Selects the object.

        [MSDN documentation for SeriesLines.Select](http://msdn.microsoft.com/en-us/library/bb238251).
        */
        bool Select();

        // ***** PROPERTIES *****

        /**
        Returns a Border object that represents the border of the object.

        [MSDN documentation for SeriesLines.Border](http://msdn.microsoft.com/en-us/library/bb237593).
        */
        wxExcelBorder GetBorder();

        /**
        Returns the ChartFormat object.  Since Excel 2007.

        [MSDN documentation for SeriesLines.Format](http://msdn.microsoft.com/en-us/library/bb242539).
        */
        wxExcelChartFormat GetFormat();

        /**
        Returns a String value that represents the name of the object.

        [MSDN documentation for SeriesLines.Name](http://msdn.microsoft.com/en-us/library/bb237596).
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
