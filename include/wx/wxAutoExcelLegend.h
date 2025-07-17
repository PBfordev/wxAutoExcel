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

        [MSDN documentation for Legend.Clear](http://msdn.microsoft.com/en-us/library/bb211830).
        */
        bool Clear();

        /**
        Deletes the object.

        [MSDN documentation for Legend.Delete](http://msdn.microsoft.com/en-us/library/bb211832).
        */
        bool Delete();

        /**
        Returns an object that represents either a collection of legend entries (a LegendEntries object) for the legend.

        [MSDN documentation for Legend.LegendEntries](http://msdn.microsoft.com/en-us/library/bb209969).
        */
        wxExcelLegendEntries LegendEntries();

        /**
        Selects the object.

        [MSDN documentation for Legend.Select](http://msdn.microsoft.com/en-us/library/bb237985).
        */
        bool Select();

        // ***** PROPERTIES *****

        /**
        True if the text in the object changes font size when the object size changes. The default value is True. Read/write Variant.

        [MSDN documentation for Legend.AutoScaleFont](http://msdn.microsoft.com/en-us/library/bb148535).
        */
        bool GetAutoScaleFont();

        /**
        True if the text in the object changes font size when the object size changes. The default value is True. Read/write Variant.

        [MSDN documentation for Legend.AutoScaleFont](http://msdn.microsoft.com/en-us/library/bb148535).
        */
        void SetAutoScaleFont(bool autoScaleFont);

        /**
        Returns the ChartFormat object. Since Excel 2007.

        [MSDN documentation for Legend.Format](http://msdn.microsoft.com/en-us/library/bb242534).
        */
        wxExcelChartFormat GetFormat();

        /**
        Returns a Double value that represents the height, in points, of the object.

        [MSDN documentation for Legend.Height](http://msdn.microsoft.com/en-us/library/bb148546).
        */
        double GetHeight();

        /**
        Sets a Double value that represents the height, in points, of the object.

        [MSDN documentation for Legend.Height](http://msdn.microsoft.com/en-us/library/bb148546).
        */
        void SetHeight(double height);

        /**
        True if a legend will occupy the chart layout space when a chart layout is being determined.  Since Excel 2007.

        [MSDN documentation for Legend.IncludeInLayout](http://msdn.microsoft.com/en-us/library/bb148548).
        */
        bool GetIncludeInLayout();

        /**
        True if a legend will occupy the chart layout space when a chart layout is being determined.  Since Excel 2007.

        [MSDN documentation for Legend.IncludeInLayout](http://msdn.microsoft.com/en-us/library/bb148548).
        */
        void SetIncludeInLayout(bool includeInLayout);

        /**
        Returns a Double value that represents the distance, in points, from the left edge of the object to the left edge of column A (on a worksheet) or the left edge of the chart area (on a chart).

        [MSDN documentation for Legend.Left](http://msdn.microsoft.com/en-us/library/bb148582).
        */
        double GetLeft();

        /**
        Sets a Double value that represents the distance, in points, from the left edge of the object to the left edge of column A (on a worksheet) or the left edge of the chart area (on a chart).

        [MSDN documentation for Legend.Left](http://msdn.microsoft.com/en-us/library/bb148582).
        */
        void SetLeft(double left);

        /**
        Returns a String value that represents the name of the object.

        [MSDN documentation for Legend.Name](http://msdn.microsoft.com/en-us/library/bb148585).
        */
        wxString GetName();

        /**
        Returns a XlLegendPosition value that represents the position of the legend on the chart.

        [MSDN documentation for Legend.Position](http://msdn.microsoft.com/en-us/library/bb148586).
        */
        XlLegendPosition GetPosition();

        /**
        Sets a XlLegendPosition value that represents the position of the legend on the chart.

        [MSDN documentation for Legend.Position](http://msdn.microsoft.com/en-us/library/bb148586).
        */
        void SetPosition(XlLegendPosition position);

        /**
        Returns a Boolean value that determines if the object has a shadow.

        [MSDN documentation for Legend.Shadow](http://msdn.microsoft.com/en-us/library/bb214649).
        */
        bool GetShadow();

        /**
        Sets a Boolean value that determines if the object has a shadow.

        [MSDN documentation for Legend.Shadow](http://msdn.microsoft.com/en-us/library/bb214649).
        */
        void SetShadow(bool shadow);

        /**
        Returns a Double value that represents the distance, in points, from the top edge of the object to the top of row 1 (on a worksheet) or the top of the chart area (on a chart).

        [MSDN documentation for Legend.Top](http://msdn.microsoft.com/en-us/library/bb214652).
        */
        double GetTop();

        /**
        Sets a Double value that represents the distance, in points, from the top edge of the object to the top of row 1 (on a worksheet) or the top of the chart area (on a chart).

        [MSDN documentation for Legend.Top](http://msdn.microsoft.com/en-us/library/bb214652).
        */
        void SetTop(double top);

        /**
        Returns a Double value that represents the width, in points, of the object.

        [MSDN documentation for Legend.Width](http://msdn.microsoft.com/en-us/library/bb214829).
        */
        double GetWidth();

        /**
        Sets a Double value that represents the width, in points, of the object.

        [MSDN documentation for Legend.Width](http://msdn.microsoft.com/en-us/library/bb214829).
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
