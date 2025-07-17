/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////


#ifndef _WXAUTOEXCEL_CHARTGROUPS_H
#define _WXAUTOEXCEL_CHARTGROUPS_H

#include "wx/wxAutoExcel_defs.h"

#if WXAUTOEXCEL_USE_CHARTS

#include "wx/wxAutoExcel_object.h"
#include "wx/wxAutoExcel_enums.h"

namespace wxAutoExcel {

    /**
    @brief Represents Microsoft Excel ChartGroup object.
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelChartGroup : public wxExcelObject
    {
    public:
        // ***** METHODS *****

        //@{
        /**
        Returns an object that represents a collection of all of the visible categories (a CategoryCollection collection) in the chart group.

        [MSDN documentation for ChartGroup.CategoryCollection](https://msdn.microsoft.com/vba/excel-vba/articles/chartgroup-categorycollection-method-excel).
        */

        wxExcelCategoryCollection CategoryCollection();
        wxExcelChartCategory CategoryCollection(long index);
        wxExcelChartCategory CategoryCollection(const wxString& name);
        //@}

        //@{
        /**
        Returns an object that represents a collection of all of the visible and filtered categories (a CategoryCollection collection) in the chart group.

        [MSDN documentation for ChartGroup.FullCategoryCollection](https://msdn.microsoft.com/vba/excel-vba/articles/chartgroup-fullcategorycollection-method-excel).
        */

        wxExcelCategoryCollection FullCategoryCollection();
        wxExcelChartCategory FullCategoryCollection(long index);
        wxExcelChartCategory FullCategoryCollection(const wxString& name);
        //@}

        //@{
        /**
        Returns an object that represents either a single series (a Series object) or a collection of all the series (a SeriesCollection collection) in the chart or chart group.

        [MSDN documentation for ChartGroup.SeriesCollection](http://msdn.microsoft.com/en-us/library/bb213903).
        */

        wxExcelSeriesCollection SeriesCollection();
        wxExcelSeries SeriesCollection(long index);
        wxExcelSeries SeriesCollection(const wxString& name);
        //@}

        // ***** PROPERTIES *****

        /**
        Returns an XlAxisGroup value that represents the type of axis group.

        [MSDN documentation for ChartGroup.AxisGroup](http://msdn.microsoft.com/en-us/library/dd787724).
        */
        XlAxisGroup GetAxisGroup();

        /**
        Specifies the number of bins in the histogram chart. Since Excel 2016.

        [MSDN documentation for ChartGroup.BinsCountValue](https://msdn.microsoft.com/VBA/Excel-VBA/articles/chartgroup-binscountvalue-property-excel).
        */
        long GetBinsCountValue();

        /**
        Specifies the number of bins in the histogram chart. Since Excel 2016.

        [MSDN documentation for ChartGroup.BinsCountValue](https://msdn.microsoft.com/VBA/Excel-VBA/articles/chartgroup-binscountvalue-property-excel).
        */
        void SetBinsCountValue(long value);

        /**
        Specifies whether a bin for values above the BinsOverflowValue is enabled. Since Excel 2016.

        [MSDN documentation for ChartGroup.BinsOverflowEnabled](https://msdn.microsoft.com/VBA/Excel-VBA/articles/chartgroup-binsoverflowenabled-property-excel).
        */
        bool GetBinsOverflowEnabled();

        /**
        Specifies whether a bin for values above the BinsOverflowValue is enabled. Since Excel 2016.

        [MSDN documentation for ChartGroup.BinsOverflowEnabled](https://msdn.microsoft.com/VBA/Excel-VBA/articles/chartgroup-binsoverflowenabled-property-excel).
        */
        void SetBinsOverflowEnabled(bool enabled);

        /**
        If an BinsOverflowEnabled is True, specifies the value above which an overflow bin is displayed. Since Excel 2016.

        [MSDN documentation for ChartGroup.BinsOverflowValue](https://msdn.microsoft.com/VBA/Excel-VBA/articles/chartgroup-binsoverflowvalue-property-excel).
        */
        double GetBinsOverflowValue();

        /**
        If an BinsOverflowEnabled is True, specifies the value above which an overflow bin is displayed. Since Excel 2016.

        [MSDN documentation for ChartGroup.BinsOverflowValue](https://msdn.microsoft.com/VBA/Excel-VBA/articles/chartgroup-binsoverflowvalue-property-excel).
        */
        void SetBinsOverflowValue(double value);

        /**
       Specifies how the horizontal axis of the histogram chart is formatted, by bins type. Since Excel 2016.

        [MSDN documentation for ChartGroup.BinsOverflowValue](https://msdn.microsoft.com/VBA/Excel-VBA/articles/chartgroup-binstype-property-excel).
        */
        XlBinsType GetBinsType();

        /**
        Specifies how the horizontal axis of the histogram chart is formatted, by bins type. Since Excel 2016.

        [MSDN documentation for ChartGroup.BinsOverflowValue](https://msdn.microsoft.com/VBA/Excel-VBA/articles/chartgroup-binstype-property-excel).
        */
        void SetBinsType(XlBinsType type);

        /**
        Specifies whether a bin for values below the BinsUnderflowValue is enabled. Since Excel 2016.

        [MSDN documentation for ChartGroup.BinsUnderflowEnabled ](https://msdn.microsoft.com/VBA/Excel-VBA/articles/chartgroup-binsunderflowenabled-property-excel).
        */
        bool GetBinsUnderflowEnabled();

        /**
        Specifies whether a bin for values below the BinsUnderflowValue is enabled. Since Excel 2016.

        [MSDN documentation for ChartGroup.BinsUnderflowEnabled ](https://msdn.microsoft.com/VBA/Excel-VBA/articles/chartgroup-binsunderflowenabled-property-excel).
        */
        void SetBinsUnderflowEnabled(bool enabled);

        /**
        If an BinsUnderflowEnabled is True, specifies the value below which an underflow bin is displayed. Since Excel 2016.

        [MSDN documentation for ChartGroup.BinsUnderflowValue](https://msdn.microsoft.com/VBA/Excel-VBA/articles/chartgroup-binsunderflowvalue-property-excel).
        */
        double GetBinsUnderflowValue();

        /**
        If an BinsUnderflowEnabled is True, specifies the value below which an underflow bin is displayed. Since Excel 2016.

        [MSDN documentation for ChartGroup.BinsUnderflowValue](https://msdn.microsoft.com/VBA/Excel-VBA/articles/chartgroup-binsunderflowvalue-property-excel).
        */
        void SetBinsUnderflowValue(double value);

        /**
        Specifies the number of points in each range. Since Excel 2016.

        [MSDN documentation for ChartGroup.BinWidthValue](https://msdn.microsoft.com/VBA/Excel-VBA/articles/chartgroup-binwidthvalue-property-excel).
        */
        double GetBinWidthValue();

        /**
        Specifies the number of points in each range. Since Excel 2016.

        [MSDN documentation for ChartGroup.BinWidthValue](https://msdn.microsoft.com/VBA/Excel-VBA/articles/chartgroup-binwidthvalue-property-excel).
        */
        void SetBinWidthValue(double value);

        /**
        Returns the scale factor for bubbles in the specified chart group. Can be an integer value from 0 (zero) to 300, corresponding to a percentage of the default size. Applies only to bubble charts.

        [MSDN documentation for ChartGroup.BubbleScale](http://msdn.microsoft.com/en-us/library/bb220893).
        */
        long GetBubbleScale();

        /**
        Sets the scale factor for bubbles in the specified chart group. Can be an integer value from 0 (zero) to 300, corresponding to a percentage of the default size. Applies only to bubble charts.

        [MSDN documentation for ChartGroup.BubbleScale](http://msdn.microsoft.com/en-us/library/bb220893).
        */
        void SetBubbleScale(long bubbleScale);

        /**
        Returns the size of the hole in a doughnut chart group. The hole size is expressed as a percentage of the chart size, between 10 and 90 percent.

        [MSDN documentation for ChartGroup.DoughnutHoleSize](http://msdn.microsoft.com/en-us/library/bb221035).
        */
        long GetDoughnutHoleSize();

        /**
        Returns a DownBars object that represents the down bars on a line chart. Applies only to line charts.

        [MSDN documentation for ChartGroup.DownBars](http://msdn.microsoft.com/en-us/library/bb221039).
        */
        wxExcelDownBars GetDownBars();

        /**
        Returns a DropLines object that represents the drop lines for a series on a line chart or area chart. Applies only to line charts or area charts.

        [MSDN documentation for ChartGroup.DropLines](http://msdn.microsoft.com/en-us/library/bb221071).
        */
        wxExcelDropLines GetDropLines();

        /**
        Returns the angle of the first pie-chart or doughnut-chart slice, in degrees (clockwise from vertical). Applies only to pie, 3-D pie, and doughnut charts. Can be a value from 0 through 360.

        [MSDN documentation for ChartGroup.FirstSliceAngle](http://msdn.microsoft.com/en-us/library/bb208513).
        */
        long GetFirstSliceAngle();

        /**
        Sets the angle of the first pie-chart or doughnut-chart slice, in degrees (clockwise from vertical). Applies only to pie, 3-D pie, and doughnut charts. Can be a value from 0 through 360.

        [MSDN documentation for ChartGroup.FirstSliceAngle](http://msdn.microsoft.com/en-us/library/bb208513).
        */
        void SetFirstSliceAngle(long firstSliceAngle);

        /**
        Bar and Column charts: Returns or sets the space between bar or column clusters, as a percentage of the bar or column width. Pie of Pie and Bar of Pie charts: Returns or sets the space between the primary and secondary sections of the chart.

        [MSDN documentation for ChartGroup.GapWidth](http://msdn.microsoft.com/en-us/library/bb208574).
        */
        long GetGapWidth();

        /**
        Bar and Column charts: Returns or sets the space between bar or column clusters, as a percentage of the bar or column width. Pie of Pie and Bar of Pie charts: Returns or sets the space between the primary and secondary sections of the chart.

        [MSDN documentation for ChartGroup.GapWidth](http://msdn.microsoft.com/en-us/library/bb208574).
        */
        void SetGapWidth(long gapWidth);

        /**
        True if the chart group has three-dimensional shading. This property only applies to surface charts and will return a run-time error if you try to set it to a non-surface chart.

        [MSDN documentation for ChartGroup.Has3DShading](http://msdn.microsoft.com/en-us/library/bb208591).
        */
        bool GetHas3DShading();

        /**
        True if the chart group has three-dimensional shading. This property only applies to surface charts and will return a run-time error if you try to set it to a non-surface chart.

        [MSDN documentation for ChartGroup.Has3DShading](http://msdn.microsoft.com/en-us/library/bb208591).
        */
        void SetHas3DShading(bool has3DShading);

        /**
        True if the line chart or area chart has drop lines. Applies only to line and area charts.

        [MSDN documentation for ChartGroup.HasDropLines](http://msdn.microsoft.com/en-us/library/bb208617).
        */
        bool GetHasDropLines();

        /**
        True if the line chart or area chart has drop lines. Applies only to line and area charts.

        [MSDN documentation for ChartGroup.HasDropLines](http://msdn.microsoft.com/en-us/library/bb208617).
        */
        void SetHasDropLines(bool hasDropLines);

        /**
        True if the line chart has high-low lines. Applies only to line charts.

        [MSDN documentation for ChartGroup.HasHiLoLines](http://msdn.microsoft.com/en-us/library/bb208630).
        */
        bool GetHasHiLoLines();

        /**
        True if the line chart has high-low lines. Applies only to line charts.

        [MSDN documentation for ChartGroup.HasHiLoLines](http://msdn.microsoft.com/en-us/library/bb208630).
        */
        void SetHasHiLoLines(bool hasHiLoLines);

        /**
        True if a radar chart has axis labels. Applies only to radar charts.

        [MSDN documentation for ChartGroup.HasRadarAxisLabels](http://msdn.microsoft.com/en-us/library/bb208652).
        */
        bool GetHasRadarAxisLabels();

        /**
        True if a radar chart has axis labels. Applies only to radar charts.

        [MSDN documentation for ChartGroup.HasRadarAxisLabels](http://msdn.microsoft.com/en-us/library/bb208652).
        */
        void SetHasRadarAxisLabels(bool hasRadarAxisLabels);

        /**
        True if a stacked column chart or bar chart has series lines or if a Pie of Pie chart or Bar of Pie chart has connector lines between the two sections. Applies only to 2-D stacked bar, 2-D stacked column, pie of pie, or bar of pie charts.

        [MSDN documentation for ChartGroup.HasSeriesLines](http://msdn.microsoft.com/en-us/library/bb208657).
        */
        bool GetHasSeriesLines();

        /**
        True if a stacked column chart or bar chart has series lines or if a Pie of Pie chart or Bar of Pie chart has connector lines between the two sections. Applies only to 2-D stacked bar, 2-D stacked column, pie of pie, or bar of pie charts.

        [MSDN documentation for ChartGroup.HasSeriesLines](http://msdn.microsoft.com/en-us/library/bb208657).
        */
        void SetHasSeriesLines(bool hasSeriesLines);

        /**
        True if a line chart has up and down bars. Applies only to line charts.

        [MSDN documentation for ChartGroup.HasUpDownBars](http://msdn.microsoft.com/en-us/library/bb208661).
        */
        bool GetHasUpDownBars();

        /**
        True if a line chart has up and down bars. Applies only to line charts.

        [MSDN documentation for ChartGroup.HasUpDownBars](http://msdn.microsoft.com/en-us/library/bb208661).
        */
        void SetHasUpDownBars(bool hasUpDownBars);

        /**
        Returns a HiLoLines object that represents the high-low lines for a series on a line chart. Applies only to line charts.

        [MSDN documentation for ChartGroup.HiLoLines](http://msdn.microsoft.com/en-us/library/bb208697).
        */
        wxExcelHiLoLines GetHiLoLines();

        /**
        Returns a Long value that represents the index number of the object within the collection of similar objects.

        [MSDN documentation for ChartGroup.Index](http://msdn.microsoft.com/en-us/library/bb179451).
        */
        long GetIndex();

        /**
        Specifies how bars and columns are positioned. Can be a value between – 100 and 100. Applies only to 2-D bar and 2-D column charts.

        [MSDN documentation for ChartGroup.Overlap](http://msdn.microsoft.com/en-us/library/bb208910).
        */
        long GetOverlap();

        /**
        Specifies how bars and columns are positioned. Can be a value between – 100 and 100. Applies only to 2-D bar and 2-D column charts.

        [MSDN documentation for ChartGroup.Overlap](http://msdn.microsoft.com/en-us/library/bb208910).
        */
        void SetOverlap(long overlap);

        /**
        Returns a TickLabels object that represents the radar axis labels for the specified chart group.

        [MSDN documentation for ChartGroup.RadarAxisLabels](http://msdn.microsoft.com/en-us/library/bb209052).
        */
        wxExcelTickLabels GetRadarAxisLabels();

        /**
        Returns the size of the secondary section of either a pie of pie chart or a bar of pie chart, as a percentage of the size of the primary pie. Can be a value from 5 to 200.

        [MSDN documentation for ChartGroup.SecondPlotSize](http://msdn.microsoft.com/en-us/library/bb221629).
        */
        long GetSecondPlotSize();

        /**
        Sets the size of the secondary section of either a pie of pie chart or a bar of pie chart, as a percentage of the size of the primary pie. Can be a value from 5 to 200.

        [MSDN documentation for ChartGroup.SecondPlotSize](http://msdn.microsoft.com/en-us/library/bb221629).
        */
        void SetSecondPlotSize(long secondPlotSize);

        /**
        Returns a SeriesLines object that represents the series lines for a 2-D stacked bar, 2-D stacked column, pie of pie, or bar of pie chart.

        [MSDN documentation for ChartGroup.SeriesLines](http://msdn.microsoft.com/en-us/library/bb221655).
        */
        wxExcelSeriesLines GetSeriesLines();

        /**
        True if negative bubbles are shown for the chart group. Valid only for bubble charts.

        [MSDN documentation for ChartGroup.ShowNegativeBubbles](http://msdn.microsoft.com/en-us/library/bb209225).
        */
        bool GetShowNegativeBubbles();

        /**
        True if negative bubbles are shown for the chart group. Valid only for bubble charts.

        [MSDN documentation for ChartGroup.ShowNegativeBubbles](http://msdn.microsoft.com/en-us/library/bb209225).
        */
        void SetShowNegativeBubbles(bool showNegativeBubbles);

        /**
        Returns what the bubble size represents on a bubble chart. Can be either of the following XlSizeRepresents constants: xlSizeIsArea or xlSizeIsWidth.

        [MSDN documentation for ChartGroup.SizeRepresents](http://msdn.microsoft.com/en-us/library/bb209240).
        */
        XlSizeRepresents GetSizeRepresents();

        /**
        Sets what the bubble size represents on a bubble chart. Can be either of the following XlSizeRepresents constants: xlSizeIsArea or xlSizeIsWidth.

        [MSDN documentation for ChartGroup.SizeRepresents](http://msdn.microsoft.com/en-us/library/bb209240).
        */
        void SetSizeRepresents(XlSizeRepresents sizeRepresents);

        /**
        Returns the way the two sections of either a pie of pie chart or a bar of pie chart are split. Read/write XlChartSplitType.

        [MSDN documentation for ChartGroup.SplitType](http://msdn.microsoft.com/en-us/library/bb209278).
        */
        XlChartSplitType GetSplitType();

        /**
        Sets the way the two sections of either a pie of pie chart or a bar of pie chart are split. Read/write XlChartSplitType.

        [MSDN documentation for ChartGroup.SplitType](http://msdn.microsoft.com/en-us/library/bb209278).
        */
        void SetSplitType(XlChartSplitType splitType);

        /**
        Returns the threshold value separating the two sections of either a pie of pie chart or a bar of pie chart. Read/write Variant.

        [MSDN documentation for ChartGroup.SplitValue](http://msdn.microsoft.com/en-us/library/bb209281).
        */
        double GetSplitValue();

        /**
        Sets the threshold value separating the two sections of either a pie of pie chart or a bar of pie chart. Read/write Variant.

        [MSDN documentation for ChartGroup.SplitValue](http://msdn.microsoft.com/en-us/library/bb209281).
        */
        void SetSplitValue(double splitValue);

        /**
        Returns an UpBars object that represents the up bars on a line chart. Applies only to line charts.

        [MSDN documentation for ChartGroup.UpBars](http://msdn.microsoft.com/en-us/library/bb221946).
        */
        wxExcelUpBars GetUpBars();

        /**
        True if Microsoft Excel assigns a different color or pattern to each data marker. The chart must contain only one series.

        [MSDN documentation for ChartGroup.VaryByCategories](http://msdn.microsoft.com/en-us/library/bb223009).
        */
        bool GetVaryByCategories();

        /**
        True if Microsoft Excel assigns a different color or pattern to each data marker. The chart must contain only one series.

        [MSDN documentation for ChartGroup.VaryByCategories](http://msdn.microsoft.com/en-us/library/bb223009).
        */
        void SetVaryByCategories(bool varyByCategories);

        /**
        Returns "ChartGroup".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("ChartGroup"); }
    };

    /**
    @brief Represents Microsoft Excel ChartGroups collection.
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelChartGroups : public wxExcelObject
    {
    public:
        // ***** METHODS *****

        //@{
        /**
        Returns a single object from a collection.

        [MSDN documentation for ChartGroups.Item](http://msdn.microsoft.com/en-us/library/bb148246).
        */
        wxExcelChartGroup Item(long index);
        wxExcelChartGroup operator[](long index);
        //@}

        // ***** PROPERTIES *****

        /**
        Returns a Long value that represents the number of objects in the collection.

        [MSDN documentation for ChartGroups.Count](http://msdn.microsoft.com/en-us/library/bb179453).
        */
        long GetCount();
        /**
        Returns "ChartGroups".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("ChartGroups"); }
    };


} // namespace wxAutoExcel

#endif // #if WXAUTOEXCEL_USE_CHARTS

#endif // _WXAUTOEXCEL_CHARTGROUPS_H
