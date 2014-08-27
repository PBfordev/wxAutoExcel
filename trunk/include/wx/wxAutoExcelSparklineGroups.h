/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Modified by:
// RCS-ID:      $Id: $
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////


#ifndef _WXAUTOEXCEL_SPARKLINEGROUPS_H
#define _WXAUTOEXCEL_SPARKLINEGROUPS_H

#include "wx/wxAutoExcel_defs.h"

#if WXAUTOEXCEL_USE_CHARTS

#include "wx/wxAutoExcelObject.h"
#include "wx/wxAutoExcel_enums.h"


namespace wxAutoExcel {

    /**
    Represents Microsoft Excel SparklineGroup object. Since Excel 2010.
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelSparklineGroup: public wxExcelObject
    {
    public:        
        // ***** METHODS *****

        /**
        Deletes the sparkline group.

        [MSDN documentation for SparklineGroup.Delete](http://msdn.microsoft.com/en-us/library/office/ff835534(v=office.14).aspx).
        */
        void Delete();

        /**
        Sets the location and the source data for the sparkline group.

        [MSDN documentation for SparklineGroup.Modify](http://msdn.microsoft.com/en-us/library/office/ff821200(v=office.14).aspx).
        */
        void Modify(wxExcelRange range, const wxString& sourceData);

        /**
        Sets the date range for the sparkline group.

        [MSDN documentation for SparklineGroup.ModifyDateRange](http://msdn.microsoft.com/en-us/library/office/ff195942(v=office.14).aspx).
        */
        void ModifyDateRange(const wxString& dateRange);

        /**
        Sets the associated Range object to modify the location of the sparkline group.

        [MSDN documentation for SparklineGroup.ModifyLocation](http://msdn.microsoft.com/en-us/library/office/ff835853(v=office.14).aspx).
        */
        void ModifyLocation(wxExcelRange location);

        /**
        Sets the range that represents the source data for the sparkline group.

        [MSDN documentation for SparklineGroup.ModifySourceData](http://msdn.microsoft.com/en-us/library/office/ff196516(v=office.14).aspx).
        */
        void ModifySourceData(const wxString& sourceData);

        // ***** PROPERTIES *****

        /**
        Returns the associated SparkAxes object. Since Excel 2007.

        [MSDN documentation for SparklineGroup.Axes](http://msdn.microsoft.com/en-us/library/office/ff836224(v=office.14).aspx).
        */
        wxExcelSparkAxes GetAxes();

        /**
        Returns the number of sparklines in the sparkline group. Since Excel 2007.

        [MSDN documentation for SparklineGroup.Count](http://msdn.microsoft.com/en-us/library/office/ff837620(v=office.14).aspx).
        */
        long GetCount();

        /**
        Gets or sets the date range for the sparkline group. Since Excel 2007.

        [MSDN documentation for SparklineGroup.DateRange](http://msdn.microsoft.com/en-us/library/office/ff197813(v=office.14).aspx).
        */
        wxString GetDateRange();

        /**
        Gets or sets the date range for the sparkline group. Since Excel 2007.

        [MSDN documentation for SparklineGroup.DateRange](http://msdn.microsoft.com/en-us/library/office/ff197813(v=office.14).aspx).
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

        [MSDN documentation for SparklineGroup.DisplayHidden](http://msdn.microsoft.com/en-us/library/office/ff838230(v=office.14).aspx).
        */
        bool GetDisplayHidden();

        /**
        Specifies if hidden cells are plotted in the sparkline group. Since Excel 2007.

        [MSDN documentation for SparklineGroup.DisplayHidden](http://msdn.microsoft.com/en-us/library/office/ff838230(v=office.14).aspx).
        */
        void SetDisplayHidden(bool displayHidden);

        //@{
        /**
        Returns a Sparkline object. Read-only Since Excel 2007.

        [MSDN documentation for SparklineGroup.Item](http://msdn.microsoft.com/en-us/library/office/ff195996(v=office.14).aspx).
        */
        wxExcelSparkline GetItem(long index);
        wxExcelSparkline operator[](long index);
        //@}

        /**
        Gets or sets the thickness of the sparklines in the sparkline group. Since Excel 2007.

        [MSDN documentation for SparklineGroup.LineWeight](http://msdn.microsoft.com/en-us/library/office/ff821580(v=office.14).aspx).
        */
        double GetLineWeight();

        /**
        Gets or sets the thickness of the sparklines in the sparkline group. Since Excel 2007.

        [MSDN documentation for SparklineGroup.LineWeight](http://msdn.microsoft.com/en-us/library/office/ff821580(v=office.14).aspx).
        */
        void SetLineWeight(double lineWeight);

        /**
        Gets or sets the Range object that represents the location of the sparkline group. Since Excel 2007.

        [MSDN documentation for SparklineGroup.Location](http://msdn.microsoft.com/en-us/library/office/ff196423(v=office.14).aspx).
        */
        wxExcelRange GetLocation();

        /**
        Gets or sets the Range object that represents the location of the sparkline group. Since Excel 2007.

        [MSDN documentation for SparklineGroup.Location](http://msdn.microsoft.com/en-us/library/office/ff196423(v=office.14).aspx).
        */
        void SetLocation(const wxExcelRange& location);

        /**
        Returns how to plot the sparkline when the data on which it is based is in a square-shaped range. Since Excel 2007.

        [MSDN documentation for SparklineGroup.PlotBy](http://msdn.microsoft.com/en-us/library/office/ff838586(v=office.14).aspx).
        */
        XlSparklineRowCol GetPlotBy();

        /**
        Sets how to plot the sparkline when the data on which it is based is in a square-shaped range. Since Excel 2007.

        [MSDN documentation for SparklineGroup.PlotBy](http://msdn.microsoft.com/en-us/library/office/ff838586(v=office.14).aspx).
        */
        void SetPlotBy(XlSparklineRowCol plotBy);

        /**
        Returns the position of the specified node as a coordinate pair. Each coordinate is expressed in points. Since Excel 2007.

        [MSDN documentation for SparklineGroup.Points](http://msdn.microsoft.com/en-us/library/office/ff837867(v=office.14).aspx).
        */
        wxExcelSparkPoints GetPoints();

        /**
        Returns a FormatColor object that represents the main series color for the sparkline group. Since Excel 2007.

        [MSDN documentation for SparklineGroup.SeriesColor](http://msdn.microsoft.com/en-us/library/office/ff194071(v=office.14).aspx).
        */
        wxExcelFormatColor GetSeriesColor();

        /**
        Returns the range that contains the source data for the sparkline group. Since Excel 2007.

        [MSDN documentation for SparklineGroup.SourceData](http://msdn.microsoft.com/en-us/library/office/ff838187(v=office.14).aspx).
        */
        wxString GetSourceData();

        /**
        Sets the range that contains the source data for the sparkline group. Since Excel 2007.

        [MSDN documentation for SparklineGroup.SourceData](http://msdn.microsoft.com/en-us/library/office/ff838187(v=office.14).aspx).
        */
        void SetSourceData(const wxString& sourceData);

        /**
        Gets or sets the type of sparkline for the group. Since Excel 2007.

        [MSDN documentation for SparklineGroup.Type](http://msdn.microsoft.com/en-us/library/office/ff196366(v=office.14).aspx).
        */
        XlSparkType GetType();

        /**
        Gets or sets the type of sparkline for the group. Since Excel 2007.

        [MSDN documentation for SparklineGroup.Type](http://msdn.microsoft.com/en-us/library/office/ff196366(v=office.14).aspx).
        */
        void SetType(XlSparkType type);

        /**
        Returns "SparklineGroup".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("SparklineGroup"); }
    };

    /**
    Represents Microsoft Excel SparklineGroups collection. Since Excel 2010.
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelSparklineGroups: public wxExcelObject
    {
    public:        
        // ***** METHODS *****

        /**
        Creates a new sparkline group and returns a SparklineGroup object.

        [MSDN documentation for SparklineGroups.Add](http://msdn.microsoft.com/en-us/library/office/ff837827(v=office.14).aspx).
        */
        wxExcelSparklineGroup Add(XlSparkType sparkType, const wxString& sourceData);

        /**
        Clears the selected sparklines.

        [MSDN documentation for SparklineGroups.Clear](http://msdn.microsoft.com/en-us/library/office/ff837606(v=office.14).aspx).
        */
        void Clear();

        /**
        Clears the selected sparkline groups.

        [MSDN documentation for SparklineGroups.ClearGroups](http://msdn.microsoft.com/en-us/library/office/ff835254(v=office.14).aspx).
        */
        void ClearGroups();

        /**
        Groups the selected sparklines.

        [MSDN documentation for SparklineGroups.Group](http://msdn.microsoft.com/en-us/library/office/ff837423(v=office.14).aspx).
        */
        void Group(wxExcelRange location);

        /**
        Ungroups the sparklines in the selected sparkline group.

        [MSDN documentation for SparklineGroups.Ungroup](http://msdn.microsoft.com/en-us/library/office/ff838848(v=office.14).aspx).
        */
        void Ungroup();

        // ***** PROPERTIES *****

        /**
        Returns the count of sparkline groups in the associated Range object.

        [MSDN documentation for SparklineGroups.Count](http://msdn.microsoft.com/en-us/library/office/ff197912(v=office.14).aspx).
        */
        long GetCount();

        //@{
        /**
        Returns a SparklineGroup object from a collection. 

        [MSDN documentation for SparklineGroups.Item](http://msdn.microsoft.com/en-us/library/office/ff822785(v=office.14).aspx).
        */
        wxExcelSparklineGroup GetItem(long index);
        wxExcelSparklineGroup operator[](long index);
        //@}

        /**
        Returns "SparklineGroups".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("SparklineGroups"); }
    };


} // namespace wxAutoExcel

#endif // #if WXAUTOEXCEL_USE_CHARTS

#endif // _WXAUTOEXCEL_SPARKLINEGROUPS_H
