/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// License:     MIT license
/////////////////////////////////////////////////////////////////////////////


#ifndef _WXAUTOEXCEL_PLOTAREA_H
#define _WXAUTOEXCEL_PLOTAREA_H

#include "wx/wxAutoExcel_defs.h"

#if WXAUTOEXCEL_USE_CHARTS

#include "wx/wxAutoExcel_object.h"

namespace wxAutoExcel {

    /**
    @brief Represents Microsoft Excel PlotArea object.
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelPlotArea : public wxExcelObject
    {
    public:
        // ***** METHODS *****

        /**
        Clears the formatting of the object.

        [MSDN documentation for PlotArea.ClearFormats](https://learn.microsoft.com/en-us/office/vba/api/excel.plotarea.clearformats).
        */
        bool ClearFormats();

        /**
        Selects the object.

        [MSDN documentation for PlotArea.Select](https://learn.microsoft.com/en-us/office/vba/api/excel.plotarea.select).
        */
        bool Select();

        // ***** PROPERTIES *****

        /**
        Returns the ChartFormat object. Since Excel 2007.

        [MSDN documentation for PlotArea.Format](https://learn.microsoft.com/en-us/office/vba/api/excel.plotarea.format).
        */
        wxExcelChartFormat GetFormat();

        /**
        Returns a Double value that represents the height, in points, of the object.

        [MSDN documentation for PlotArea.Height](https://learn.microsoft.com/en-us/office/vba/api/excel.plotarea.height).
        */
        double GetHeight();

        /**
        Sets a Double value that represents the height, in points, of the object.

        [MSDN documentation for PlotArea.Height](https://learn.microsoft.com/en-us/office/vba/api/excel.plotarea.height).
        */
        void SetHeight(double height);

        /**
        Returns the inside height of the plot area, in points.

        [MSDN documentation for PlotArea.InsideHeight](https://learn.microsoft.com/en-us/office/vba/api/excel.plotarea.insideheight).
        */
        double GetInsideHeight();

        /**
        Sets the inside height of the plot area, in points.

        [MSDN documentation for PlotArea.InsideHeight](https://learn.microsoft.com/en-us/office/vba/api/excel.plotarea.insideheight).
        */
        void SetInsideHeight(double insideHeight);

        /**
        Returns the distance from the chart edge to the inside left edge of the plot area, in points.

        [MSDN documentation for PlotArea.InsideLeft](https://learn.microsoft.com/en-us/office/vba/api/excel.plotarea.insideleft).
        */
        double GetInsideLeft();

        /**
        Sets the distance from the chart edge to the inside left edge of the plot area, in points.

        [MSDN documentation for PlotArea.InsideLeft](https://learn.microsoft.com/en-us/office/vba/api/excel.plotarea.insideleft).
        */
        void SetInsideLeft(double insideLeft);

        /**
        Returns the distance from the chart edge to the inside top edge of the plot area, in points.

        [MSDN documentation for PlotArea.InsideTop](https://learn.microsoft.com/en-us/office/vba/api/excel.plotarea.insidetop).
        */
        double GetInsideTop();

        /**
        Sets the distance from the chart edge to the inside top edge of the plot area, in points.

        [MSDN documentation for PlotArea.InsideTop](https://learn.microsoft.com/en-us/office/vba/api/excel.plotarea.insidetop).
        */
        void SetInsideTop(double insideTop);

        /**
        Returns the inside width of the plot area, in points.

        [MSDN documentation for PlotArea.InsideWidth](https://learn.microsoft.com/en-us/office/vba/api/excel.plotarea.insidewidth).
        */
        double GetInsideWidth();

        /**
        Sets the inside width of the plot area, in points.

        [MSDN documentation for PlotArea.InsideWidth](https://learn.microsoft.com/en-us/office/vba/api/excel.plotarea.insidewidth).
        */
        void SetInsideWidth(double insideWidth);

        /**
        Returns a Double value that represents the distance, in points, from the left edge of the object to the left edge of column A (on a worksheet) or the left edge of the chart area (on a chart).

        [MSDN documentation for PlotArea.Left](https://learn.microsoft.com/en-us/office/vba/api/excel.plotarea.left).
        */
        double GetLeft();

        /**
        Sets a Double value that represents the distance, in points, from the left edge of the object to the left edge of column A (on a worksheet) or the left edge of the chart area (on a chart).

        [MSDN documentation for PlotArea.Left](https://learn.microsoft.com/en-us/office/vba/api/excel.plotarea.left).
        */
        void SetLeft(double left);

        /**
        Returns a String value that represents the name of the object.

        [MSDN documentation for PlotArea.Name](https://learn.microsoft.com/en-us/office/vba/api/excel.plotarea.name).
        */
        wxString GetName();

        /**
        Returns the position of the plot area on the chart. Since Excel 2007.

        [MSDN documentation for PlotArea.Position](https://learn.microsoft.com/en-us/office/vba/api/excel.plotarea.position).
        */
        XlChartElementPosition GetPosition();

        /**
        Sets the position of the plot area on the chart. Since Excel 2007.

        [MSDN documentation for PlotArea.Position](https://learn.microsoft.com/en-us/office/vba/api/excel.plotarea.position).
        */
        void SetPosition(XlChartElementPosition position);

        /**
        Returns a Double value that represents the distance, in points, from the top edge of the object to the top of row 1 (on a worksheet) or the top of the chart area (on a chart).

        [MSDN documentation for PlotArea.Top](https://learn.microsoft.com/en-us/office/vba/api/excel.plotarea.top).
        */
        double GetTop();

        /**
        Sets a Double value that represents the distance, in points, from the top edge of the object to the top of row 1 (on a worksheet) or the top of the chart area (on a chart).

        [MSDN documentation for PlotArea.Top](https://learn.microsoft.com/en-us/office/vba/api/excel.plotarea.top).
        */
        void SetTop(double top);

        /**
        Returns a Double value that represents the width, in points, of the object.

        [MSDN documentation for PlotArea.Width](https://learn.microsoft.com/en-us/office/vba/api/excel.plotarea.width).
        */
        double GetWidth();

        /**
        Sets a Double value that represents the width, in points, of the object.

        [MSDN documentation for PlotArea.Width](https://learn.microsoft.com/en-us/office/vba/api/excel.plotarea.width).
        */
        void SetWidth(double width);

        /**
        Returns "PlotArea".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("PlotArea"); }
    };


} // namespace wxAutoExcel

#endif // #if WXAUTOEXCEL_USE_CHARTS

#endif // _WXAUTOEXCEL_CHART_H
