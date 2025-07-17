/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
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

        [MSDN documentation for PlotArea.ClearFormats](http://msdn.microsoft.com/en-us/library/bb178773).
        */
        bool ClearFormats();

        /**
        Selects the object.

        [MSDN documentation for PlotArea.Select](http://msdn.microsoft.com/en-us/library/bb238173).
        */
        bool Select();

        // ***** PROPERTIES *****

        /**
        Returns the ChartFormat object. Since Excel 2007.

        [MSDN documentation for PlotArea.Format](http://msdn.microsoft.com/en-us/library/bb242536).
        */
        wxExcelChartFormat GetFormat();

        /**
        Returns a Double value that represents the height, in points, of the object.

        [MSDN documentation for PlotArea.Height](http://msdn.microsoft.com/en-us/library/bb237446).
        */
        double GetHeight();

        /**
        Sets a Double value that represents the height, in points, of the object.

        [MSDN documentation for PlotArea.Height](http://msdn.microsoft.com/en-us/library/bb237446).
        */
        void SetHeight(double height);

        /**
        Returns the inside height of the plot area, in points. Read-only Double.

        [MSDN documentation for PlotArea.InsideHeight](http://msdn.microsoft.com/en-us/library/bb177653).
        */
        double GetInsideHeight();

        /**
        Returns the distance from the chart edge to the inside left edge of the plot area, in points. Read-only Double.

        [MSDN documentation for PlotArea.InsideLeft](http://msdn.microsoft.com/en-us/library/bb177655).
        */
        double GetInsideLeft();

        /**
        Returns the distance from the chart edge to the inside top edge of the plot area, in points. Read-only Double.

        [MSDN documentation for PlotArea.InsideTop](http://msdn.microsoft.com/en-us/library/bb177658).
        */
        double GetInsideTop();

        /**
        Returns the inside width of the plot area, in points. Read-only Double.

        [MSDN documentation for PlotArea.InsideWidth](http://msdn.microsoft.com/en-us/library/bb177660).
        */
        double GetInsideWidth();

        /**
        Returns a Double value that represents the distance, in points, from the left edge of the object to the left edge of column A (on a worksheet) or the left edge of the chart area (on a chart).

        [MSDN documentation for PlotArea.Left](http://msdn.microsoft.com/en-us/library/bb237450).
        */
        double GetLeft();

        /**
        Sets a Double value that represents the distance, in points, from the left edge of the object to the left edge of column A (on a worksheet) or the left edge of the chart area (on a chart).

        [MSDN documentation for PlotArea.Left](http://msdn.microsoft.com/en-us/library/bb237450).
        */
        void SetLeft(double left);

        /**
        Returns a String value that represents the name of the object.

        [MSDN documentation for PlotArea.Name](http://msdn.microsoft.com/en-us/library/bb237452).
        */
        wxString GetName();

        /**
        Returns the position of the plot area on the chart. Since Excel 2007.

        [MSDN documentation for PlotArea.Position](http://msdn.microsoft.com/en-us/library/bb240589).
        */
        XlChartElementPosition GetPosition();

        /**
        Sets the position of the plot area on the chart. Since Excel 2007.

        [MSDN documentation for PlotArea.Position](http://msdn.microsoft.com/en-us/library/bb240589).
        */
        void SetPosition(XlChartElementPosition position);

        /**
        Returns a Double value that represents the distance, in points, from the top edge of the object to the top of row 1 (on a worksheet) or the top of the chart area (on a chart).

        [MSDN documentation for PlotArea.Top](http://msdn.microsoft.com/en-us/library/bb238554).
        */
        double GetTop();

        /**
        Sets a Double value that represents the distance, in points, from the top edge of the object to the top of row 1 (on a worksheet) or the top of the chart area (on a chart).

        [MSDN documentation for PlotArea.Top](http://msdn.microsoft.com/en-us/library/bb238554).
        */
        void SetTop(double top);

        /**
        Returns a Double value that represents the width, in points, of the object.

        [MSDN documentation for PlotArea.Width](http://msdn.microsoft.com/en-us/library/bb238560).
        */
        double GetWidth();

        /**
        Sets a Double value that represents the width, in points, of the object.

        [MSDN documentation for PlotArea.Width](http://msdn.microsoft.com/en-us/library/bb238560).
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
