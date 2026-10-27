/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// License:     MIT license
/////////////////////////////////////////////////////////////////////////////


#ifndef _WXAUTOEXCEL_SPARKLINEGROUPS_H
#define _WXAUTOEXCEL_SPARKLINEGROUPS_H

#include "wx/wxAutoExcel_defs.h"

#if WXAUTOEXCEL_USE_CHARTS

#include "wx/wxAutoExcel_object.h"
#include "wx/wxAutoExcel_enums.h"


namespace wxAutoExcel {

    /**
    @brief Represents Microsoft Excel SparklineGroup object. Since Excel 2010.
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelSparklineGroup: public wxExcelObject
    {
    public:
        // ***** METHODS *****

        /**
        Deletes the sparkline group.

        [MSDN documentation for SparklineGroup.Delete](https://learn.microsoft.com/en-us/office/vba/api/excel.sparklinegroup.delete).
        */
        void Delete();

        /**
        Sets the location and the source data for the sparkline group.

        [MSDN documentation for SparklineGroup.Modify](https://learn.microsoft.com/en-us/office/vba/api/excel.sparklinegroup.modify).
        */
        void Modify(wxExcelRange range, const wxString& sourceData);

        /**
        Sets the date range for the sparkline group.

        [MSDN documentation for SparklineGroup.ModifyDateRange](https://learn.microsoft.com/en-us/office/vba/api/excel.sparklinegroup.modifydaterange).
        */
        void ModifyDateRange(const wxString& dateRange);

        /**
        Sets the associated Range object to modify the location of the sparkline group.

        [MSDN documentation for SparklineGroup.ModifyLocation](https://learn.microsoft.com/en-us/office/vba/api/excel.sparklinegroup.modifylocation).
        */
        void ModifyLocation(wxExcelRange location);

        /**
        Sets the range that represents the source data for the sparkline group.

        [MSDN documentation for SparklineGroup.ModifySourceData](https://learn.microsoft.com/en-us/office/vba/api/excel.sparklinegroup.modifysourcedata).
        */
        void ModifySourceData(const wxString& sourceData);

        // ***** PROPERTIES *****

        /**
        Returns the associated SparkAxes object. Since Excel 2007.

        [MSDN documentation for SparklineGroup.Axes](https://learn.microsoft.com/en-us/office/vba/api/excel.sparklinegroup.axes).
        */
        wxExcelSparkAxes GetAxes();

        /**
        Returns the number of sparklines in the sparkline group. Since Excel 2007.

        [MSDN documentation for SparklineGroup.Count](https://learn.microsoft.com/en-us/office/vba/api/excel.sparklinegroup.count).
        */
        long GetCount();

        /**
        Gets or sets the date range for the sparkline group. Since Excel 2007.

        [MSDN documentation for SparklineGroup.DateRange](https://learn.microsoft.com/en-us/office/vba/api/excel.sparklinegroup.daterange).
        */
        wxString GetDateRange();

        /**
        Gets or sets the date range for the sparkline group. Since Excel 2007.

        [MSDN documentation for SparklineGroup.DateRange](https://learn.microsoft.com/en-us/office/vba/api/excel.sparklinegroup.daterange).
        */
        void SetDateRange(const wxString& dateRange);

        /**
        Returns the way that blank cells are plotted on a chart. Can be one of the XlDisplayBlanksAs constants.

        [MSDN documentation for SparklineGroup.DisplayBlanksAs]().
        */
        XlDisplayBlanksAs GetDisplayBlanksAs();

        /**
        Sets the way that blank cells are plotted on a chart. Can be one of the XlDisplayBlanksAs constants.

        [MSDN documentation for SparklineGroup.DisplayBlanksAs]().
        */
        void SetDisplayBlanksAs(XlDisplayBlanksAs displayBlanksAs);

        /**
        Specifies if hidden cells are plotted in the sparkline group. Since Excel 2007.

        [MSDN documentation for SparklineGroup.DisplayHidden](https://learn.microsoft.com/en-us/office/vba/api/excel.sparklinegroup.displayhidden).
        */
        bool GetDisplayHidden();

        /**
        Specifies if hidden cells are plotted in the sparkline group. Since Excel 2007.

        [MSDN documentation for SparklineGroup.DisplayHidden](https://learn.microsoft.com/en-us/office/vba/api/excel.sparklinegroup.displayhidden).
        */
        void SetDisplayHidden(bool displayHidden);

        //@{
        /**
        Returns a Sparkline object. Read-only Since Excel 2007.

        [MSDN documentation for SparklineGroup.Item](https://learn.microsoft.com/en-us/office/vba/api/excel.sparklinegroup.item).
        */
        wxExcelSparkline GetItem(long index);
        wxExcelSparkline operator[](long index);
        //@}

        /**
        Gets or sets the thickness of the sparklines in the sparkline group. Since Excel 2007.

        [MSDN documentation for SparklineGroup.LineWeight](https://learn.microsoft.com/en-us/office/vba/api/excel.sparklinegroup.lineweight).
        */
        double GetLineWeight();

        /**
        Gets or sets the thickness of the sparklines in the sparkline group. Since Excel 2007.

        [MSDN documentation for SparklineGroup.LineWeight](https://learn.microsoft.com/en-us/office/vba/api/excel.sparklinegroup.lineweight).
        */
        void SetLineWeight(double lineWeight);

        /**
        Gets or sets the Range object that represents the location of the sparkline group. Since Excel 2007.

        [MSDN documentation for SparklineGroup.Location](https://learn.microsoft.com/en-us/office/vba/api/excel.sparklinegroup.location).
        */
        wxExcelRange GetLocation();

        /**
        Gets or sets the Range object that represents the location of the sparkline group. Since Excel 2007.

        [MSDN documentation for SparklineGroup.Location](https://learn.microsoft.com/en-us/office/vba/api/excel.sparklinegroup.location).
        */
        void SetLocation(const wxExcelRange& location);

        /**
        Returns how to plot the sparkline when the data on which it is based is in a square-shaped range. Since Excel 2007.

        [MSDN documentation for SparklineGroup.PlotBy](https://learn.microsoft.com/en-us/office/vba/api/excel.sparklinegroup.plotby).
        */
        XlSparklineRowCol GetPlotBy();

        /**
        Sets how to plot the sparkline when the data on which it is based is in a square-shaped range. Since Excel 2007.

        [MSDN documentation for SparklineGroup.PlotBy](https://learn.microsoft.com/en-us/office/vba/api/excel.sparklinegroup.plotby).
        */
        void SetPlotBy(XlSparklineRowCol plotBy);

        /**
        Returns the position of the specified node as a coordinate pair. Each coordinate is expressed in points. Since Excel 2007.

        [MSDN documentation for SparklineGroup.Points](https://learn.microsoft.com/en-us/office/vba/api/excel.sparklinegroup.points).
        */
        wxExcelSparkPoints GetPoints();

        /**
        Returns a FormatColor object that represents the main series color for the sparkline group. Since Excel 2007.

        [MSDN documentation for SparklineGroup.SeriesColor](https://learn.microsoft.com/en-us/office/vba/api/excel.sparklinegroup.seriescolor).
        */
        wxExcelFormatColor GetSeriesColor();

        /**
        Returns the range that contains the source data for the sparkline group. Since Excel 2007.

        [MSDN documentation for SparklineGroup.SourceData](https://learn.microsoft.com/en-us/office/vba/api/excel.sparklinegroup.sourcedata).
        */
        wxString GetSourceData();

        /**
        Sets the range that contains the source data for the sparkline group. Since Excel 2007.

        [MSDN documentation for SparklineGroup.SourceData](https://learn.microsoft.com/en-us/office/vba/api/excel.sparklinegroup.sourcedata).
        */
        void SetSourceData(const wxString& sourceData);

        /**
        Gets or sets the type of sparkline for the group. Since Excel 2007.

        [MSDN documentation for SparklineGroup.Type](https://learn.microsoft.com/en-us/office/vba/api/excel.sparklinegroup.type).
        */
        XlSparkType GetType();

        /**
        Gets or sets the type of sparkline for the group. Since Excel 2007.

        [MSDN documentation for SparklineGroup.Type](https://learn.microsoft.com/en-us/office/vba/api/excel.sparklinegroup.type).
        */
        void SetType(XlSparkType type);

        /**
        The iterator over the items, see wxExcelCollectionIterator.
        */
        typedef wxExcelCollectionIterator<wxExcelSparklineGroup, wxExcelSparkline> iterator;

        //@{
        /**
        Returns the iterator referring to the first item, or past the last item,
        which allows iterating over the items with a range-based for loop,
        see wxExcelCollectionIterator.
        */
        iterator begin() { return iterator(this, 1); }
        iterator end() { return iterator(this, GetCount() + 1); }
        //@}

        /**
        Returns "SparklineGroup".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("SparklineGroup"); }
    };

    /**
    @brief Represents Microsoft Excel SparklineGroups collection. Since Excel 2010.
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelSparklineGroups: public wxExcelObject
    {
    public:
        // ***** METHODS *****

        /**
        Creates a new sparkline group and returns a SparklineGroup object.

        [MSDN documentation for SparklineGroups.Add](https://learn.microsoft.com/en-us/office/vba/api/excel.sparklinegroups.add).
        */
        wxExcelSparklineGroup Add(XlSparkType sparkType, const wxString& sourceData);

        /**
        Clears the selected sparklines.

        [MSDN documentation for SparklineGroups.Clear](https://learn.microsoft.com/en-us/office/vba/api/excel.sparklinegroups.clear).
        */
        void Clear();

        /**
        Clears the selected sparkline groups.

        [MSDN documentation for SparklineGroups.ClearGroups](https://learn.microsoft.com/en-us/office/vba/api/excel.sparklinegroups.cleargroups).
        */
        void ClearGroups();

        /**
        Groups the selected sparklines.

        [MSDN documentation for SparklineGroups.Group](https://learn.microsoft.com/en-us/office/vba/api/excel.sparklinegroups.group).
        */
        void Group(wxExcelRange location);

        /**
        Ungroups the sparklines in the selected sparkline group.

        [MSDN documentation for SparklineGroups.Ungroup](https://learn.microsoft.com/en-us/office/vba/api/excel.sparklinegroups.ungroup).
        */
        void Ungroup();

        // ***** PROPERTIES *****

        /**
        Returns the count of sparkline groups in the associated Range object.

        [MSDN documentation for SparklineGroups.Count](https://learn.microsoft.com/en-us/office/vba/api/excel.sparklinegroups.count).
        */
        long GetCount();

        //@{
        /**
        Returns a SparklineGroup object from a collection.

        [MSDN documentation for SparklineGroups.Item](https://learn.microsoft.com/en-us/office/vba/api/excel.sparklinegroups.item).
        */
        wxExcelSparklineGroup GetItem(long index);
        wxExcelSparklineGroup operator[](long index);
        //@}

        /**
        The iterator over the items, see wxExcelCollectionIterator.
        */
        typedef wxExcelCollectionIterator<wxExcelSparklineGroups, wxExcelSparklineGroup> iterator;

        //@{
        /**
        Returns the iterator referring to the first item, or past the last item,
        which allows iterating over the items with a range-based for loop,
        see wxExcelCollectionIterator.
        */
        iterator begin() { return iterator(this, 1); }
        iterator end() { return iterator(this, GetCount() + 1); }
        //@}

        /**
        Returns "SparklineGroups".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("SparklineGroups"); }
    };


} // namespace wxAutoExcel

#endif // #if WXAUTOEXCEL_USE_CHARTS

#endif // _WXAUTOEXCEL_SPARKLINEGROUPS_H
