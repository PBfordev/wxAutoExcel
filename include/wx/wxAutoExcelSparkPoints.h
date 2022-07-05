/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Modified by:
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
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

        [MSDN documentation for SparkPoints.Firstpoint](http://msdn.microsoft.com/en-us/library/office/ff198353(v=office.14).aspx).
        */
        wxExcelSparkColor GetFirstpoint();

        /**
        Returns a SparkColor object that represents the color and visibility of the marker for the highest point of data on a sparkline. Read-only Since Excel 2007.

        [MSDN documentation for SparkPoints.Highpoint](http://msdn.microsoft.com/en-us/library/office/ff196431(v=office.14).aspx).
        */
        wxExcelSparkColor GetHighpoint();

        /**
        Returns a SparkColor object that represents the color and visibility of the marker for the last point of data on a sparkline. Read-only Since Excel 2007.

        [MSDN documentation for SparkPoints.Lastpoint](http://msdn.microsoft.com/en-us/library/office/ff196316(v=office.14).aspx).
        */
        wxExcelSparkColor GetLastpoint();

        /**
        Returns a SparkColor object that represents the color and visibility of the marker for the lowest point of data on a sparkline. Read-only Since Excel 2007.

        [MSDN documentation for SparkPoints.Lowpoint](http://msdn.microsoft.com/en-us/library/office/ff837055(v=office.14).aspx).
        */
        wxExcelSparkColor GetLowpoint();

        /**
        Returns a SparkColor object that represents the color and visibility of markers for positive points of data on a sparkline. Read-only Since Excel 2007.

        [MSDN documentation for SparkPoints.Markers](http://msdn.microsoft.com/en-us/library/office/ff194201(v=office.14).aspx).
        */
        wxExcelSparkColor GetMarkers();

        /**
        Returns a SparkColor object that represents the color and visibility of markers for negative points of data on a sparkline. Read-only Since Excel 2007.

        [MSDN documentation for SparkPoints.Negative](http://msdn.microsoft.com/en-us/library/office/ff821824(v=office.14).aspx).
        */
        wxExcelSparkColor GetNegative();

        /**
        Returns the parent SparklineGroup object for the specified object. Since Excel 2007.

        [MSDN documentation for SparkPoints.Parent](http://msdn.microsoft.com/en-us/library/office/ff193308(v=office.14).aspx).
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
