/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Modified by:
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////


#ifndef _WXAUTOEXCEL_CHARTAREA_H
#define _WXAUTOEXCEL_CHARTAREA_H

#include "wx/wxAutoExcel_defs.h"

#if WXAUTOEXCEL_USE_CHARTS

#include "wx/wxAutoExcel_object.h"

namespace wxAutoExcel {

    /**
    Represents Microsoft Excel ChartArea object.
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelChartArea : public wxExcelObject
    {
    public:
        // ***** METHODS *****

        /**
        Clears the entire object.

        [MSDN documentation for ChartArea.Clear](http://msdn.microsoft.com/en-us/library/bb211641).
        */
        bool Clear();

        /**
        Clears the data from a chart but leaves the formatting.

        [MSDN documentation for ChartArea.ClearContents](http://msdn.microsoft.com/en-us/library/bb211644).
        */
        bool ClearContents();

        /**
        Clears the formatting of the object.

        [MSDN documentation for ChartArea.ClearFormats](http://msdn.microsoft.com/en-us/library/bb148184).
        */
        bool ClearFormats();

        /**
        Copies the object to the Clipboard.

        [MSDN documentation for ChartArea.Copy](http://msdn.microsoft.com/en-us/library/bb148190).
        */
        bool Copy();

        /**
        Selects the object.

        [MSDN documentation for ChartArea.Select](http://msdn.microsoft.com/en-us/library/bb213874).
        */
        bool Select();

        // ***** PROPERTIES *****

        /**
        True if the text in the object changes font size when the object size changes. The default value is True. Read/write Variant.

        [MSDN documentation for ChartArea.AutoScaleFont](http://msdn.microsoft.com/en-us/library/bb179412).
        */
        bool GetAutoScaleFont();

        /**
        True if the text in the object changes font size when the object size changes. The default value is True. Read/write Variant.

        [MSDN documentation for ChartArea.AutoScaleFont](http://msdn.microsoft.com/en-us/library/bb179412).
        */
        void SetAutoScaleFont(bool autoScaleFont);

        /**
        Returns the ChartFormat object Since Excel 2007.

        [MSDN documentation for ChartArea.Format](http://msdn.microsoft.com/en-us/library/bb242493).
        */
        wxExcelChartFormat GetFormat();

        /**
        Returns a Double value that represents the height, in points, of the object.

        [MSDN documentation for ChartArea.Height](http://msdn.microsoft.com/en-us/library/bb179415).
        */
        double GetHeight();

        /**
        Sets a Double value that represents the height, in points, of the object.

        [MSDN documentation for ChartArea.Height](http://msdn.microsoft.com/en-us/library/bb179415).
        */
        void SetHeight(double height);

        /**
        Returns a Double value that represents the distance, in points, from the left edge of the object to the left edge of column A (on a worksheet) or the left edge of the chart area (on a chart).

        [MSDN documentation for ChartArea.Left](http://msdn.microsoft.com/en-us/library/bb179416).
        */
        double GetLeft();

        /**
        Sets a Double value that represents the distance, in points, from the left edge of the object to the left edge of column A (on a worksheet) or the left edge of the chart area (on a chart).

        [MSDN documentation for ChartArea.Left](http://msdn.microsoft.com/en-us/library/bb179416).
        */
        void SetLeft(double left);

        /**
        Returns a String value that represents the name of the object.

        [MSDN documentation for ChartArea.Name](http://msdn.microsoft.com/en-us/library/bb179419).
        */
        wxString GetName();

        /**
        True if the chart area of the chart has rounded corners. Since Excel 2010.
        [MSDN documentation for ChartArea.RoundedCorners](http://msdn.microsoft.com/en-us/library/office/ff194918%28v=office.14%29.aspx).
        */
        bool GetRoundedCorners();

        /**
        True if the chart area of the chart has rounded corners. Since Excel 2010.
        [MSDN documentation for ChartArea.RoundedCorners](http://msdn.microsoft.com/en-us/library/office/ff194918%28v=office.14%29.aspx).
        */
        void SetRoundedCorners(bool roundedCorners);

        /**
        Returns a boolean value that determines if the object has a shadow.

        [MSDN documentation for ChartArea.Shadow](http://msdn.microsoft.com/en-us/library/bb148823).
        */
        bool GetShadow();

        /**
        Sets a Boolean value that determines if the object has a shadow.

        [MSDN documentation for ChartArea.Shadow](http://msdn.microsoft.com/en-us/library/bb148823).
        */
        void SetShadow(bool shadow);

        /**
        Returns a Double value that represents the distance, in points, from the top edge of the object to the top of row 1 (on a worksheet) or the top of the chart area (on a chart).

        [MSDN documentation for ChartArea.Top](http://msdn.microsoft.com/en-us/library/bb148830).
        */
        double GetTop();

        /**
        Returns a Double value that represents the distance, in points, from the top edge of the object to the top of row 1 (on a worksheet) or the top of the chart area (on a chart).

        [MSDN documentation for ChartArea.Top](http://msdn.microsoft.com/en-us/library/bb148830).
        */
        void SetTop(double top);

        /**
        Returns a Double value that represents the width, in points, of the object.

        [MSDN documentation for ChartArea.Width](http://msdn.microsoft.com/en-us/library/bb148833).
        */
        double GetWidth();

        /**
        Sets a Double value that represents the width, in points, of the object.

        [MSDN documentation for ChartArea.Width](http://msdn.microsoft.com/en-us/library/bb148833).
        */
        void SetWidth(double width);

        /**
        Returns "ChartArea".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("ChartArea"); }
    };


} // namespace wxAutoExcel

#endif // #if WXAUTOEXCEL_USE_CHARTS

#endif // _WXAUTOEXCEL_CHARTAREA_H
