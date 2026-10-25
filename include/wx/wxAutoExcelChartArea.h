/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// License:     MIT license
/////////////////////////////////////////////////////////////////////////////


#ifndef _WXAUTOEXCEL_CHARTAREA_H
#define _WXAUTOEXCEL_CHARTAREA_H

#include "wx/wxAutoExcel_defs.h"

#if WXAUTOEXCEL_USE_CHARTS

#include "wx/wxAutoExcel_object.h"

namespace wxAutoExcel {

    /**
    @brief Represents Microsoft Excel ChartArea object.
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelChartArea : public wxExcelObject
    {
    public:
        // ***** METHODS *****

        /**
        Clears the entire object.

        [MSDN documentation for ChartArea.Clear](https://learn.microsoft.com/en-us/office/vba/api/excel.chartarea.clear).
        */
        bool Clear();

        /**
        Clears the data from a chart but leaves the formatting.

        [MSDN documentation for ChartArea.ClearContents](https://learn.microsoft.com/en-us/office/vba/api/excel.chartarea.clearcontents).
        */
        bool ClearContents();

        /**
        Clears the formatting of the object.

        [MSDN documentation for ChartArea.ClearFormats](https://learn.microsoft.com/en-us/office/vba/api/excel.chartarea.clearformats).
        */
        bool ClearFormats();

        /**
        Copies the object to the Clipboard.

        [MSDN documentation for ChartArea.Copy](https://learn.microsoft.com/en-us/office/vba/api/excel.chartarea.copy).
        */
        bool Copy();

        /**
        Selects the object.

        [MSDN documentation for ChartArea.Select](https://learn.microsoft.com/en-us/office/vba/api/excel.chartarea.select).
        */
        bool Select();

        // ***** PROPERTIES *****

        /**
        True if the text in the object changes font size when the object size changes. The default value is True. Read/write Variant.

        [MSDN documentation for ChartArea.AutoScaleFont](https://learn.microsoft.com/en-us/previous-versions/office/developer/office-2007/bb179412%28v=office.12%29).
        */
        bool GetAutoScaleFont();

        /**
        True if the text in the object changes font size when the object size changes. The default value is True. Read/write Variant.

        [MSDN documentation for ChartArea.AutoScaleFont](https://learn.microsoft.com/en-us/previous-versions/office/developer/office-2007/bb179412%28v=office.12%29).
        */
        void SetAutoScaleFont(bool autoScaleFont);

        /**
        Returns a ChartFillFormat object that contains fill formatting properties for the object.

        [MSDN documentation for ChartArea.Fill](https://learn.microsoft.com/en-us/office/vba/api/excel.chartarea.fill).
        */
        wxExcelChartFillFormat GetFill();

        /**
        Returns the ChartFormat object Since Excel 2007.

        [MSDN documentation for ChartArea.Format](https://learn.microsoft.com/en-us/office/vba/api/excel.chartarea.format).
        */
        wxExcelChartFormat GetFormat();

        /**
        Returns a Double value that represents the height, in points, of the object.

        [MSDN documentation for ChartArea.Height](https://learn.microsoft.com/en-us/office/vba/api/excel.chartarea.height).
        */
        double GetHeight();

        /**
        Sets a Double value that represents the height, in points, of the object.

        [MSDN documentation for ChartArea.Height](https://learn.microsoft.com/en-us/office/vba/api/excel.chartarea.height).
        */
        void SetHeight(double height);

        /**
        Returns a Double value that represents the distance, in points, from the left edge of the object to the left edge of column A (on a worksheet) or the left edge of the chart area (on a chart).

        [MSDN documentation for ChartArea.Left](https://learn.microsoft.com/en-us/office/vba/api/excel.chartarea.left).
        */
        double GetLeft();

        /**
        Sets a Double value that represents the distance, in points, from the left edge of the object to the left edge of column A (on a worksheet) or the left edge of the chart area (on a chart).

        [MSDN documentation for ChartArea.Left](https://learn.microsoft.com/en-us/office/vba/api/excel.chartarea.left).
        */
        void SetLeft(double left);

        /**
        Returns a String value that represents the name of the object.

        [MSDN documentation for ChartArea.Name](https://learn.microsoft.com/en-us/office/vba/api/excel.chartarea.name).
        */
        wxString GetName();

        /**
        True if the chart area of the chart has rounded corners. Since Excel 2010.
        [MSDN documentation for ChartArea.RoundedCorners](https://learn.microsoft.com/en-us/office/vba/api/excel.chartarea.roundedcorners).
        */
        bool GetRoundedCorners();

        /**
        True if the chart area of the chart has rounded corners. Since Excel 2010.
        [MSDN documentation for ChartArea.RoundedCorners](https://learn.microsoft.com/en-us/office/vba/api/excel.chartarea.roundedcorners).
        */
        void SetRoundedCorners(bool roundedCorners);

        /**
        Returns a boolean value that determines if the object has a shadow.

        [MSDN documentation for ChartArea.Shadow](https://learn.microsoft.com/en-us/office/vba/api/excel.chartarea.shadow).
        */
        bool GetShadow();

        /**
        Sets a Boolean value that determines if the object has a shadow.

        [MSDN documentation for ChartArea.Shadow](https://learn.microsoft.com/en-us/office/vba/api/excel.chartarea.shadow).
        */
        void SetShadow(bool shadow);

        /**
        Returns a Double value that represents the distance, in points, from the top edge of the object to the top of row 1 (on a worksheet) or the top of the chart area (on a chart).

        [MSDN documentation for ChartArea.Top](https://learn.microsoft.com/en-us/office/vba/api/excel.chartarea.top).
        */
        double GetTop();

        /**
        Returns a Double value that represents the distance, in points, from the top edge of the object to the top of row 1 (on a worksheet) or the top of the chart area (on a chart).

        [MSDN documentation for ChartArea.Top](https://learn.microsoft.com/en-us/office/vba/api/excel.chartarea.top).
        */
        void SetTop(double top);

        /**
        Returns a Double value that represents the width, in points, of the object.

        [MSDN documentation for ChartArea.Width](https://learn.microsoft.com/en-us/office/vba/api/excel.chartarea.width).
        */
        double GetWidth();

        /**
        Sets a Double value that represents the width, in points, of the object.

        [MSDN documentation for ChartArea.Width](https://learn.microsoft.com/en-us/office/vba/api/excel.chartarea.width).
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
