/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// License:     MIT license
/////////////////////////////////////////////////////////////////////////////


#ifndef _WXAUTOEXCEL_LEGEND_H
#define _WXAUTOEXCEL_LEGEND_H

#include "wx/wxAutoExcel_defs.h"

#if WXAUTOEXCEL_USE_CHARTS

#include "wx/wxAutoExcel_object.h"

namespace wxAutoExcel {

    /**
    @brief Represents Microsoft Excel Legend object.
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelLegend : public wxExcelObject
    {
    public:
        // ***** METHODS *****

        /**
        Clears the entire object.

        [MSDN documentation for Legend.Clear](https://learn.microsoft.com/en-us/office/vba/api/excel.legend.clear).
        */
        bool Clear();

        /**
        Deletes the object.

        [MSDN documentation for Legend.Delete](https://learn.microsoft.com/en-us/office/vba/api/excel.legend.delete).
        */
        bool Delete();

        /**
        Returns an object that represents either a collection of legend entries (a LegendEntries object) for the legend.

        [MSDN documentation for Legend.LegendEntries](https://learn.microsoft.com/en-us/office/vba/api/excel.legend.legendentries).
        */
        wxExcelLegendEntries LegendEntries();

        /**
        Selects the object.

        [MSDN documentation for Legend.Select](https://learn.microsoft.com/en-us/office/vba/api/excel.legend.select).
        */
        bool Select();

        // ***** PROPERTIES *****

        /**
        True if the text in the object changes font size when the object size changes. The default value is True. Read/write Variant.

        [MSDN documentation for Legend.AutoScaleFont](https://learn.microsoft.com/en-us/previous-versions/office/developer/office-2007/bb148535%28v=office.12%29).
        */
        bool GetAutoScaleFont();

        /**
        True if the text in the object changes font size when the object size changes. The default value is True. Read/write Variant.

        [MSDN documentation for Legend.AutoScaleFont](https://learn.microsoft.com/en-us/previous-versions/office/developer/office-2007/bb148535%28v=office.12%29).
        */
        void SetAutoScaleFont(bool autoScaleFont);

        /**
        Returns a ChartFillFormat object that contains fill formatting properties for the object.

        [MSDN documentation for Legend.Fill](https://learn.microsoft.com/en-us/office/vba/api/excel.legend.fill).
        */
        wxExcelChartFillFormat GetFill();

        /**
        Returns the ChartFormat object. Since Excel 2007.

        [MSDN documentation for Legend.Format](https://learn.microsoft.com/en-us/office/vba/api/excel.legend.format).
        */
        wxExcelChartFormat GetFormat();

        /**
        Returns a Double value that represents the height, in points, of the object.

        [MSDN documentation for Legend.Height](https://learn.microsoft.com/en-us/office/vba/api/excel.legend.height).
        */
        double GetHeight();

        /**
        Sets a Double value that represents the height, in points, of the object.

        [MSDN documentation for Legend.Height](https://learn.microsoft.com/en-us/office/vba/api/excel.legend.height).
        */
        void SetHeight(double height);

        /**
        True if a legend will occupy the chart layout space when a chart layout is being determined.  Since Excel 2007.

        [MSDN documentation for Legend.IncludeInLayout](https://learn.microsoft.com/en-us/office/vba/api/excel.legend.includeinlayout).
        */
        bool GetIncludeInLayout();

        /**
        True if a legend will occupy the chart layout space when a chart layout is being determined.  Since Excel 2007.

        [MSDN documentation for Legend.IncludeInLayout](https://learn.microsoft.com/en-us/office/vba/api/excel.legend.includeinlayout).
        */
        void SetIncludeInLayout(bool includeInLayout);

        /**
        Returns a Double value that represents the distance, in points, from the left edge of the object to the left edge of column A (on a worksheet) or the left edge of the chart area (on a chart).

        [MSDN documentation for Legend.Left](https://learn.microsoft.com/en-us/office/vba/api/excel.legend.left).
        */
        double GetLeft();

        /**
        Sets a Double value that represents the distance, in points, from the left edge of the object to the left edge of column A (on a worksheet) or the left edge of the chart area (on a chart).

        [MSDN documentation for Legend.Left](https://learn.microsoft.com/en-us/office/vba/api/excel.legend.left).
        */
        void SetLeft(double left);

        /**
        Returns a String value that represents the name of the object.

        [MSDN documentation for Legend.Name](https://learn.microsoft.com/en-us/office/vba/api/excel.legend.name).
        */
        wxString GetName();

        /**
        Returns a XlLegendPosition value that represents the position of the legend on the chart.

        [MSDN documentation for Legend.Position](https://learn.microsoft.com/en-us/office/vba/api/excel.legend.position).
        */
        XlLegendPosition GetPosition();

        /**
        Sets a XlLegendPosition value that represents the position of the legend on the chart.

        [MSDN documentation for Legend.Position](https://learn.microsoft.com/en-us/office/vba/api/excel.legend.position).
        */
        void SetPosition(XlLegendPosition position);

        /**
        Returns a Boolean value that determines if the object has a shadow.

        [MSDN documentation for Legend.Shadow](https://learn.microsoft.com/en-us/office/vba/api/excel.legend.shadow).
        */
        bool GetShadow();

        /**
        Sets a Boolean value that determines if the object has a shadow.

        [MSDN documentation for Legend.Shadow](https://learn.microsoft.com/en-us/office/vba/api/excel.legend.shadow).
        */
        void SetShadow(bool shadow);

        /**
        Returns a Double value that represents the distance, in points, from the top edge of the object to the top of row 1 (on a worksheet) or the top of the chart area (on a chart).

        [MSDN documentation for Legend.Top](https://learn.microsoft.com/en-us/office/vba/api/excel.legend.top).
        */
        double GetTop();

        /**
        Sets a Double value that represents the distance, in points, from the top edge of the object to the top of row 1 (on a worksheet) or the top of the chart area (on a chart).

        [MSDN documentation for Legend.Top](https://learn.microsoft.com/en-us/office/vba/api/excel.legend.top).
        */
        void SetTop(double top);

        /**
        Returns a Double value that represents the width, in points, of the object.

        [MSDN documentation for Legend.Width](https://learn.microsoft.com/en-us/office/vba/api/excel.legend.width).
        */
        double GetWidth();

        /**
        Sets a Double value that represents the width, in points, of the object.

        [MSDN documentation for Legend.Width](https://learn.microsoft.com/en-us/office/vba/api/excel.legend.width).
        */
        void SetWidth(double width);

        /**
        Returns "Legend".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("Legend"); }
    };


} // namespace wxAutoExcel

#endif // #if WXAUTOEXCEL_USE_CHARTS

#endif // _WXAUTOEXCEL_LEGEND_H
