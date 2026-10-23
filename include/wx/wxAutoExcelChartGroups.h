/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// License:     MIT license
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

        [MSDN documentation for ChartGroup.CategoryCollection](https://learn.microsoft.com/en-us/office/vba/api/excel.chartgroup.categorycollection).
        */

        wxExcelCategoryCollection CategoryCollection();
        wxExcelChartCategory CategoryCollection(long index);
        wxExcelChartCategory CategoryCollection(const wxString& name);
        //@}

        //@{
        /**
        Returns an object that represents a collection of all of the visible and filtered categories (a CategoryCollection collection) in the chart group.

        [MSDN documentation for ChartGroup.FullCategoryCollection](https://learn.microsoft.com/en-us/office/vba/api/excel.chartgroup.fullcategorycollection).
        */

        wxExcelCategoryCollection FullCategoryCollection();
        wxExcelChartCategory FullCategoryCollection(long index);
        wxExcelChartCategory FullCategoryCollection(const wxString& name);
        //@}

        //@{
        /**
        Returns an object that represents either a single series (a Series object) or a collection of all the series (a SeriesCollection collection) in the chart or chart group.

        [MSDN documentation for ChartGroup.SeriesCollection](https://learn.microsoft.com/en-us/office/vba/api/excel.chartgroup.seriescollection).
        */

        wxExcelSeriesCollection SeriesCollection();
        wxExcelSeries SeriesCollection(long index);
        wxExcelSeries SeriesCollection(const wxString& name);
        //@}

        // ***** PROPERTIES *****

        /**
        Returns an XlAxisGroup value that represents the type of axis group.

        [MSDN documentation for ChartGroup.AxisGroup](https://learn.microsoft.com/en-us/office/vba/api/excel.chartgroup.axisgroup).
        */
        XlAxisGroup GetAxisGroup();

        /**
        Sets the group for the specified chart.

        [MSDN documentation for ChartGroup.AxisGroup](https://learn.microsoft.com/en-us/office/vba/api/excel.chartgroup.axisgroup).
        */
        void SetAxisGroup(XlAxisGroup axisGroup);

        /**
        Specifies the number of bins in the histogram chart. Since Excel 2016.

        [MSDN documentation for ChartGroup.BinsCountValue](https://learn.microsoft.com/en-us/office/vba/api/excel.chartgroup.binscountvalue).
        */
        long GetBinsCountValue();

        /**
        Specifies the number of bins in the histogram chart. Since Excel 2016.

        [MSDN documentation for ChartGroup.BinsCountValue](https://learn.microsoft.com/en-us/office/vba/api/excel.chartgroup.binscountvalue).
        */
        void SetBinsCountValue(long value);

        /**
        Specifies whether a bin for values above the BinsOverflowValue is enabled. Since Excel 2016.

        [MSDN documentation for ChartGroup.BinsOverflowEnabled](https://learn.microsoft.com/en-us/office/vba/api/excel.chartgroup.binsoverflowenabled).
        */
        bool GetBinsOverflowEnabled();

        /**
        Specifies whether a bin for values above the BinsOverflowValue is enabled. Since Excel 2016.

        [MSDN documentation for ChartGroup.BinsOverflowEnabled](https://learn.microsoft.com/en-us/office/vba/api/excel.chartgroup.binsoverflowenabled).
        */
        void SetBinsOverflowEnabled(bool enabled);

        /**
        If an BinsOverflowEnabled is True, specifies the value above which an overflow bin is displayed. Since Excel 2016.

        [MSDN documentation for ChartGroup.BinsOverflowValue](https://learn.microsoft.com/en-us/office/vba/api/excel.chartgroup.binsoverflowvalue).
        */
        double GetBinsOverflowValue();

        /**
        If an BinsOverflowEnabled is True, specifies the value above which an overflow bin is displayed. Since Excel 2016.

        [MSDN documentation for ChartGroup.BinsOverflowValue](https://learn.microsoft.com/en-us/office/vba/api/excel.chartgroup.binsoverflowvalue).
        */
        void SetBinsOverflowValue(double value);

        /**
       Specifies how the horizontal axis of the histogram chart is formatted, by bins type. Since Excel 2016.

        [MSDN documentation for ChartGroup.BinsType](https://learn.microsoft.com/en-us/office/vba/api/excel.chartgroup.binstype).
        */
        XlBinsType GetBinsType();

        /**
        Specifies how the horizontal axis of the histogram chart is formatted, by bins type. Since Excel 2016.

        [MSDN documentation for ChartGroup.BinsType](https://learn.microsoft.com/en-us/office/vba/api/excel.chartgroup.binstype).
        */
        void SetBinsType(XlBinsType type);

        /**
        Specifies whether a bin for values below the BinsUnderflowValue is enabled. Since Excel 2016.

        [MSDN documentation for ChartGroup.BinsUnderflowEnabled](https://learn.microsoft.com/en-us/office/vba/api/excel.chartgroup.binsunderflowenabled).
        */
        bool GetBinsUnderflowEnabled();

        /**
        Specifies whether a bin for values below the BinsUnderflowValue is enabled. Since Excel 2016.

        [MSDN documentation for ChartGroup.BinsUnderflowEnabled](https://learn.microsoft.com/en-us/office/vba/api/excel.chartgroup.binsunderflowenabled).
        */
        void SetBinsUnderflowEnabled(bool enabled);

        /**
        If an BinsUnderflowEnabled is True, specifies the value below which an underflow bin is displayed. Since Excel 2016.

        [MSDN documentation for ChartGroup.BinsUnderflowValue](https://learn.microsoft.com/en-us/office/vba/api/excel.chartgroup.binsunderflowvalue).
        */
        double GetBinsUnderflowValue();

        /**
        If an BinsUnderflowEnabled is True, specifies the value below which an underflow bin is displayed. Since Excel 2016.

        [MSDN documentation for ChartGroup.BinsUnderflowValue](https://learn.microsoft.com/en-us/office/vba/api/excel.chartgroup.binsunderflowvalue).
        */
        void SetBinsUnderflowValue(double value);

        /**
        Specifies the number of points in each range. Since Excel 2016.

        [MSDN documentation for ChartGroup.BinWidthValue](https://learn.microsoft.com/en-us/office/vba/api/excel.chartgroup.binwidthvalue).
        */
        double GetBinWidthValue();

        /**
        Specifies the number of points in each range. Since Excel 2016.

        [MSDN documentation for ChartGroup.BinWidthValue](https://learn.microsoft.com/en-us/office/vba/api/excel.chartgroup.binwidthvalue).
        */
        void SetBinWidthValue(double value);

        /**
        Returns the scale factor for bubbles in the specified chart group. Can be an integer value from 0 (zero) to 300, corresponding to a percentage of the default size. Applies only to bubble charts.

        [MSDN documentation for ChartGroup.BubbleScale](https://learn.microsoft.com/en-us/office/vba/api/excel.chartgroup.bubblescale).
        */
        long GetBubbleScale();

        /**
        Sets the scale factor for bubbles in the specified chart group. Can be an integer value from 0 (zero) to 300, corresponding to a percentage of the default size. Applies only to bubble charts.

        [MSDN documentation for ChartGroup.BubbleScale](https://learn.microsoft.com/en-us/office/vba/api/excel.chartgroup.bubblescale).
        */
        void SetBubbleScale(long bubbleScale);

        /**
        Returns the size of the hole in a doughnut chart group. The hole size is expressed as a percentage of the chart size, between 10 and 90 percent.

        [MSDN documentation for ChartGroup.DoughnutHoleSize](https://learn.microsoft.com/en-us/office/vba/api/excel.chartgroup.doughnutholesize).
        */
        long GetDoughnutHoleSize();

        /**
        Sets the size of the hole in a doughnut chart group, as a percentage of the chart size, between 10 and 90 percent.

        [MSDN documentation for ChartGroup.DoughnutHoleSize](https://learn.microsoft.com/en-us/office/vba/api/excel.chartgroup.doughnutholesize).
        */
        void SetDoughnutHoleSize(long doughnutHoleSize);

        /**
        Returns a DownBars object that represents the down bars on a line chart. Applies only to line charts.

        [MSDN documentation for ChartGroup.DownBars](https://learn.microsoft.com/en-us/office/vba/api/excel.chartgroup.downbars).
        */
        wxExcelDownBars GetDownBars();

        /**
        Returns a DropLines object that represents the drop lines for a series on a line chart or area chart. Applies only to line charts or area charts.

        [MSDN documentation for ChartGroup.DropLines](https://learn.microsoft.com/en-us/office/vba/api/excel.chartgroup.droplines).
        */
        wxExcelDropLines GetDropLines();

        /**
        Returns the angle of the first pie-chart or doughnut-chart slice, in degrees (clockwise from vertical). Applies only to pie, 3-D pie, and doughnut charts. Can be a value from 0 through 360.

        [MSDN documentation for ChartGroup.FirstSliceAngle](https://learn.microsoft.com/en-us/office/vba/api/excel.chartgroup.firstsliceangle).
        */
        long GetFirstSliceAngle();

        /**
        Sets the angle of the first pie-chart or doughnut-chart slice, in degrees (clockwise from vertical). Applies only to pie, 3-D pie, and doughnut charts. Can be a value from 0 through 360.

        [MSDN documentation for ChartGroup.FirstSliceAngle](https://learn.microsoft.com/en-us/office/vba/api/excel.chartgroup.firstsliceangle).
        */
        void SetFirstSliceAngle(long firstSliceAngle);

        /**
        Bar and Column charts: Returns or sets the space between bar or column clusters, as a percentage of the bar or column width. Pie of Pie and Bar of Pie charts: Returns or sets the space between the primary and secondary sections of the chart.

        [MSDN documentation for ChartGroup.GapWidth](https://learn.microsoft.com/en-us/office/vba/api/excel.chartgroup.gapwidth).
        */
        long GetGapWidth();

        /**
        Bar and Column charts: Returns or sets the space between bar or column clusters, as a percentage of the bar or column width. Pie of Pie and Bar of Pie charts: Returns or sets the space between the primary and secondary sections of the chart.

        [MSDN documentation for ChartGroup.GapWidth](https://learn.microsoft.com/en-us/office/vba/api/excel.chartgroup.gapwidth).
        */
        void SetGapWidth(long gapWidth);

        /**
        True if the chart group has three-dimensional shading. This property only applies to surface charts and will return a run-time error if you try to set it to a non-surface chart.

        [MSDN documentation for ChartGroup.Has3DShading](https://learn.microsoft.com/en-us/office/vba/api/excel.chartgroup.has3dshading).
        */
        bool GetHas3DShading();

        /**
        True if the chart group has three-dimensional shading. This property only applies to surface charts and will return a run-time error if you try to set it to a non-surface chart.

        [MSDN documentation for ChartGroup.Has3DShading](https://learn.microsoft.com/en-us/office/vba/api/excel.chartgroup.has3dshading).
        */
        void SetHas3DShading(bool has3DShading);

        /**
        True if the line chart or area chart has drop lines. Applies only to line and area charts.

        [MSDN documentation for ChartGroup.HasDropLines](https://learn.microsoft.com/en-us/office/vba/api/excel.chartgroup.hasdroplines).
        */
        bool GetHasDropLines();

        /**
        True if the line chart or area chart has drop lines. Applies only to line and area charts.

        [MSDN documentation for ChartGroup.HasDropLines](https://learn.microsoft.com/en-us/office/vba/api/excel.chartgroup.hasdroplines).
        */
        void SetHasDropLines(bool hasDropLines);

        /**
        True if the line chart has high-low lines. Applies only to line charts.

        [MSDN documentation for ChartGroup.HasHiLoLines](https://learn.microsoft.com/en-us/office/vba/api/excel.chartgroup.hashilolines).
        */
        bool GetHasHiLoLines();

        /**
        True if the line chart has high-low lines. Applies only to line charts.

        [MSDN documentation for ChartGroup.HasHiLoLines](https://learn.microsoft.com/en-us/office/vba/api/excel.chartgroup.hashilolines).
        */
        void SetHasHiLoLines(bool hasHiLoLines);

        /**
        True if a radar chart has axis labels. Applies only to radar charts.

        [MSDN documentation for ChartGroup.HasRadarAxisLabels](https://learn.microsoft.com/en-us/office/vba/api/excel.chartgroup.hasradaraxislabels).
        */
        bool GetHasRadarAxisLabels();

        /**
        True if a radar chart has axis labels. Applies only to radar charts.

        [MSDN documentation for ChartGroup.HasRadarAxisLabels](https://learn.microsoft.com/en-us/office/vba/api/excel.chartgroup.hasradaraxislabels).
        */
        void SetHasRadarAxisLabels(bool hasRadarAxisLabels);

        /**
        True if a stacked column chart or bar chart has series lines or if a Pie of Pie chart or Bar of Pie chart has connector lines between the two sections. Applies only to 2-D stacked bar, 2-D stacked column, pie of pie, or bar of pie charts.

        [MSDN documentation for ChartGroup.HasSeriesLines](https://learn.microsoft.com/en-us/office/vba/api/excel.chartgroup.hasserieslines).
        */
        bool GetHasSeriesLines();

        /**
        True if a stacked column chart or bar chart has series lines or if a Pie of Pie chart or Bar of Pie chart has connector lines between the two sections. Applies only to 2-D stacked bar, 2-D stacked column, pie of pie, or bar of pie charts.

        [MSDN documentation for ChartGroup.HasSeriesLines](https://learn.microsoft.com/en-us/office/vba/api/excel.chartgroup.hasserieslines).
        */
        void SetHasSeriesLines(bool hasSeriesLines);

        /**
        True if a line chart has up and down bars. Applies only to line charts.

        [MSDN documentation for ChartGroup.HasUpDownBars](https://learn.microsoft.com/en-us/office/vba/api/excel.chartgroup.hasupdownbars).
        */
        bool GetHasUpDownBars();

        /**
        True if a line chart has up and down bars. Applies only to line charts.

        [MSDN documentation for ChartGroup.HasUpDownBars](https://learn.microsoft.com/en-us/office/vba/api/excel.chartgroup.hasupdownbars).
        */
        void SetHasUpDownBars(bool hasUpDownBars);

        /**
        Returns a HiLoLines object that represents the high-low lines for a series on a line chart. Applies only to line charts.

        [MSDN documentation for ChartGroup.HiLoLines](https://learn.microsoft.com/en-us/office/vba/api/excel.chartgroup.hilolines).
        */
        wxExcelHiLoLines GetHiLoLines();

        /**
        Returns a Long value that represents the index number of the object within the collection of similar objects.

        [MSDN documentation for ChartGroup.Index](https://learn.microsoft.com/en-us/office/vba/api/excel.chartgroup.index).
        */
        long GetIndex();

        /**
        Specifies how bars and columns are positioned. Can be a value between – 100 and 100. Applies only to 2-D bar and 2-D column charts.

        [MSDN documentation for ChartGroup.Overlap](https://learn.microsoft.com/en-us/office/vba/api/excel.chartgroup.overlap).
        */
        long GetOverlap();

        /**
        Specifies how bars and columns are positioned. Can be a value between – 100 and 100. Applies only to 2-D bar and 2-D column charts.

        [MSDN documentation for ChartGroup.Overlap](https://learn.microsoft.com/en-us/office/vba/api/excel.chartgroup.overlap).
        */
        void SetOverlap(long overlap);

        /**
        Returns a TickLabels object that represents the radar axis labels for the specified chart group.

        [MSDN documentation for ChartGroup.RadarAxisLabels](https://learn.microsoft.com/en-us/office/vba/api/excel.chartgroup.radaraxislabels).
        */
        wxExcelTickLabels GetRadarAxisLabels();

        /**
        Returns the size of the secondary section of either a pie of pie chart or a bar of pie chart, as a percentage of the size of the primary pie. Can be a value from 5 to 200.

        [MSDN documentation for ChartGroup.SecondPlotSize](https://learn.microsoft.com/en-us/office/vba/api/excel.chartgroup.secondplotsize).
        */
        long GetSecondPlotSize();

        /**
        Sets the size of the secondary section of either a pie of pie chart or a bar of pie chart, as a percentage of the size of the primary pie. Can be a value from 5 to 200.

        [MSDN documentation for ChartGroup.SecondPlotSize](https://learn.microsoft.com/en-us/office/vba/api/excel.chartgroup.secondplotsize).
        */
        void SetSecondPlotSize(long secondPlotSize);

        /**
        Returns a SeriesLines object that represents the series lines for a 2-D stacked bar, 2-D stacked column, pie of pie, or bar of pie chart.

        [MSDN documentation for ChartGroup.SeriesLines](https://learn.microsoft.com/en-us/office/vba/api/excel.chartgroup.serieslines).
        */
        wxExcelSeriesLines GetSeriesLines();

        /**
        True if negative bubbles are shown for the chart group. Valid only for bubble charts.

        [MSDN documentation for ChartGroup.ShowNegativeBubbles](https://learn.microsoft.com/en-us/office/vba/api/excel.chartgroup.shownegativebubbles).
        */
        bool GetShowNegativeBubbles();

        /**
        True if negative bubbles are shown for the chart group. Valid only for bubble charts.

        [MSDN documentation for ChartGroup.ShowNegativeBubbles](https://learn.microsoft.com/en-us/office/vba/api/excel.chartgroup.shownegativebubbles).
        */
        void SetShowNegativeBubbles(bool showNegativeBubbles);

        /**
        Returns what the bubble size represents on a bubble chart. Can be either of the following XlSizeRepresents constants: xlSizeIsArea or xlSizeIsWidth.

        [MSDN documentation for ChartGroup.SizeRepresents](https://learn.microsoft.com/en-us/office/vba/api/excel.chartgroup.sizerepresents).
        */
        XlSizeRepresents GetSizeRepresents();

        /**
        Sets what the bubble size represents on a bubble chart. Can be either of the following XlSizeRepresents constants: xlSizeIsArea or xlSizeIsWidth.

        [MSDN documentation for ChartGroup.SizeRepresents](https://learn.microsoft.com/en-us/office/vba/api/excel.chartgroup.sizerepresents).
        */
        void SetSizeRepresents(XlSizeRepresents sizeRepresents);

        /**
        Returns the way the two sections of either a pie of pie chart or a bar of pie chart are split. Read/write XlChartSplitType.

        [MSDN documentation for ChartGroup.SplitType](https://learn.microsoft.com/en-us/office/vba/api/excel.chartgroup.splittype).
        */
        XlChartSplitType GetSplitType();

        /**
        Sets the way the two sections of either a pie of pie chart or a bar of pie chart are split. Read/write XlChartSplitType.

        [MSDN documentation for ChartGroup.SplitType](https://learn.microsoft.com/en-us/office/vba/api/excel.chartgroup.splittype).
        */
        void SetSplitType(XlChartSplitType splitType);

        /**
        Returns the threshold value separating the two sections of either a pie of pie chart or a bar of pie chart. Read/write Variant.

        [MSDN documentation for ChartGroup.SplitValue](https://learn.microsoft.com/en-us/office/vba/api/excel.chartgroup.splitvalue).
        */
        double GetSplitValue();

        /**
        Sets the threshold value separating the two sections of either a pie of pie chart or a bar of pie chart. Read/write Variant.

        [MSDN documentation for ChartGroup.SplitValue](https://learn.microsoft.com/en-us/office/vba/api/excel.chartgroup.splitvalue).
        */
        void SetSplitValue(double splitValue);

        /**
        Returns an UpBars object that represents the up bars on a line chart. Applies only to line charts.

        [MSDN documentation for ChartGroup.UpBars](https://learn.microsoft.com/en-us/office/vba/api/excel.chartgroup.upbars).
        */
        wxExcelUpBars GetUpBars();

        /**
        True if Microsoft Excel assigns a different color or pattern to each data marker. The chart must contain only one series.

        [MSDN documentation for ChartGroup.VaryByCategories](https://learn.microsoft.com/en-us/office/vba/api/excel.chartgroup.varybycategories).
        */
        bool GetVaryByCategories();

        /**
        True if Microsoft Excel assigns a different color or pattern to each data marker. The chart must contain only one series.

        [MSDN documentation for ChartGroup.VaryByCategories](https://learn.microsoft.com/en-us/office/vba/api/excel.chartgroup.varybycategories).
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

        [MSDN documentation for ChartGroups.Item](https://learn.microsoft.com/en-us/office/vba/api/excel.chartgroups.item).
        */
        wxExcelChartGroup Item(long index);
        wxExcelChartGroup operator[](long index);
        //@}

        // ***** PROPERTIES *****

        /**
        Returns a Long value that represents the number of objects in the collection.

        [MSDN documentation for ChartGroups.Count](https://learn.microsoft.com/en-us/office/vba/api/excel.chartgroups.count).
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
