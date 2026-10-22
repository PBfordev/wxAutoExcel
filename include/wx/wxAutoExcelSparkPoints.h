/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// License:     MIT license
/////////////////////////////////////////////////////////////////////////////


#ifndef _WXAUTOEXCEL_SPARKPOINTS_H
#define _WXAUTOEXCEL_SPARKPOINTS_H

#include "wx/wxAutoExcel_defs.h"

#if WXAUTOEXCEL_USE_CHARTS

#include "wx/wxAutoExcel_object.h"
#include "wx/wxAutoExcel_enums.h"


namespace wxAutoExcel {


    /**
    @brief Represents Microsoft Excel SparkPoints object. Since Excel 2010.
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelSparkPoints : public wxExcelObject
    {
    public:
        // ***** PROPERTIES *****

        /**
        Returns a SparkColor object that represents the color and visibility of the marker for the first point of data on a sparkline. Read-only Since Excel 2007.

        [MSDN documentation for SparkPoints.Firstpoint](https://learn.microsoft.com/en-us/office/vba/api/excel.sparkpoints.firstpoint).aspx).
        */
        wxExcelSparkColor GetFirstpoint();

        /**
        Returns a SparkColor object that represents the color and visibility of the marker for the highest point of data on a sparkline. Read-only Since Excel 2007.

        [MSDN documentation for SparkPoints.Highpoint](https://learn.microsoft.com/en-us/office/vba/api/excel.sparkpoints.highpoint).aspx).
        */
        wxExcelSparkColor GetHighpoint();

        /**
        Returns a SparkColor object that represents the color and visibility of the marker for the last point of data on a sparkline. Read-only Since Excel 2007.

        [MSDN documentation for SparkPoints.Lastpoint](https://learn.microsoft.com/en-us/office/vba/api/excel.sparkpoints.lastpoint).aspx).
        */
        wxExcelSparkColor GetLastpoint();

        /**
        Returns a SparkColor object that represents the color and visibility of the marker for the lowest point of data on a sparkline. Read-only Since Excel 2007.

        [MSDN documentation for SparkPoints.Lowpoint](https://learn.microsoft.com/en-us/office/vba/api/excel.sparkpoints.lowpoint).aspx).
        */
        wxExcelSparkColor GetLowpoint();

        /**
        Returns a SparkColor object that represents the color and visibility of markers for positive points of data on a sparkline. Read-only Since Excel 2007.

        [MSDN documentation for SparkPoints.Markers](https://learn.microsoft.com/en-us/office/vba/api/excel.sparkpoints.markers).aspx).
        */
        wxExcelSparkColor GetMarkers();

        /**
        Returns a SparkColor object that represents the color and visibility of markers for negative points of data on a sparkline. Read-only Since Excel 2007.

        [MSDN documentation for SparkPoints.Negative](https://learn.microsoft.com/en-us/office/vba/api/excel.sparkpoints.negative).aspx).
        */
        wxExcelSparkColor GetNegative();

        /**
        Returns the parent SparklineGroup object for the specified object. Since Excel 2007.

        [MSDN documentation for SparkPoints.Parent](https://learn.microsoft.com/en-us/office/vba/api/excel.sparkpoints.parent).aspx).
        */
        wxExcelSparklineGroup GetParent();


        /**
        Returns "SparkPoints".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("SparkPoints"); }
    };


} // namespace wxAutoExcel

#endif // #if WXAUTOEXCEL_USE_CHARTS

#endif // _WXAUTOEXCEL_SPARKPOINTS_H
