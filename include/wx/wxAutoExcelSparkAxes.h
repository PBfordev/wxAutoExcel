/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// License:     MIT license
/////////////////////////////////////////////////////////////////////////////


#ifndef _WXAUTOEXCEL_SPARKAXES_H
#define _WXAUTOEXCEL_SPARKAXES_H

#include "wx/wxAutoExcel_defs.h"

#if WXAUTOEXCEL_USE_CHARTS

#include "wx/wxAutoExcel_object.h"
#include "wx/wxAutoExcel_enums.h"


namespace wxAutoExcel {

    /**
    @brief Represents Microsoft Excel SparkHorizontalAxis object. Since Excel 2010.
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelSparkHorizontalAxis : public wxExcelObject
    {
    public:
        // ***** PROPERTIES *****

        /**
        Returns a SparkColor object that specifies the color of the horizontal axis of the sparkline.

        [MSDN documentation for SparkHorizontalAxis.Axis](https://learn.microsoft.com/en-us/office/vba/api/excel.sparkhorizontalaxis.axis).aspx).
        */
        wxExcelSparkColor GetAxis();

        /**
        Returns whether the horizontal axis of the sparkline is based on date values.

        [MSDN documentation for SparkHorizontalAxis.IsDateAxis](https://learn.microsoft.com/en-us/office/vba/api/excel.sparkhorizontalaxis.isdateaxis).aspx).
        */
        bool GetIsDateAxis();

        /**
        Returns the parent SparklineGroup object for the specified object.

        [MSDN documentation for SparkHorizontalAxis.Parent](https://learn.microsoft.com/en-us/office/vba/api/excel.sparkhorizontalaxis.parent).aspx).
        */
        wxExcelSparklineGroup GetParent();

        /**
        Returns whether the points on the horizontal axis are plotted in right-to-left order.

        [MSDN documentation for SparkHorizontalAxis.RightToLeftPlotOrder](https://learn.microsoft.com/en-us/office/vba/api/excel.sparkhorizontalaxis.righttoleftplotorder).aspx).
        */
        bool GetRightToLeftPlotOrder();

        /**
        Sets whether the points on the horizontal axis are plotted in right-to-left order.

        [MSDN documentation for SparkHorizontalAxis.RightToLeftPlotOrder](https://learn.microsoft.com/en-us/office/vba/api/excel.sparkhorizontalaxis.righttoleftplotorder).
        */
        void SetRightToLeftPlotOrder(bool rightToLeftPlotOrder);


        /**
        Returns "SparkHorizontalAxis".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("SparkHorizontalAxis"); }
    };

   /**
    @brief Represents Microsoft Excel SparkVerticalAxis object. Since Excel 2010.
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelSparkVerticalAxis : public wxExcelObject
    {
    public:
        // ***** PROPERTIES *****

        /**
        Returns the custom maximum value for the vertical axis of a sparkline.

        [MSDN documentation for SparkVerticalAxis.CustomMaxScaleValue](https://learn.microsoft.com/en-us/office/vba/api/excel.sparkverticalaxis.custommaxscalevalue).aspx).
        */
        double GetCustomMaxScaleValue();

        /**
        Sets the custom maximum value for the vertical axis of a sparkline.

        [MSDN documentation for SparkVerticalAxis.CustomMaxScaleValue](https://learn.microsoft.com/en-us/office/vba/api/excel.sparkverticalaxis.custommaxscalevalue).aspx).
        */
        void SetCustomMaxScaleValue(double customMaxScaleValue);

        /**
        Returns the custom minimum value for the vertical axis of a sparkline.

        [MSDN documentation for SparkVerticalAxis.CustomMinScaleValue](https://learn.microsoft.com/en-us/office/vba/api/excel.sparkverticalaxis.customminscalevalue).aspx).
        */
        double GetCustomMinScaleValue();

        /**
        Sets the custom minimum value for the vertical axis of a sparkline.

        [MSDN documentation for SparkVerticalAxis.CustomMinScaleValue](https://learn.microsoft.com/en-us/office/vba/api/excel.sparkverticalaxis.customminscalevalue).aspx).
        */
        void SetCustomMinScaleValue(double customMinScaleValue);

        /**
        Returns how the maximum value of the vertical axis of the sparkline is scaled relative to other sparklines in the group.

        [MSDN documentation for SparkVerticalAxis.MaxScaleType](https://learn.microsoft.com/en-us/office/vba/api/excel.sparkverticalaxis.maxscaletype).aspx).
        */
        XlSparkScale GetMaxScaleType();

        /**
        Sets how the maximum value of the vertical axis of the sparkline is scaled relative to other sparklines in the group.

        [MSDN documentation for SparkVerticalAxis.MaxScaleType](https://learn.microsoft.com/en-us/office/vba/api/excel.sparkverticalaxis.maxscaletype).aspx).
        */
        void SetMaxScaleType(XlSparkScale maxScaleType);

        /**
        Returns how the minimum value of the vertical axis of the sparkline is scaled relative to other sparklines in the group.

        [MSDN documentation for SparkVerticalAxis.MinScaleType](https://learn.microsoft.com/en-us/office/vba/api/excel.sparkverticalaxis.minscaletype).aspx).
        */
        XlSparkScale GetMinScaleType();

        /**
        Sets how the minimum value of the vertical axis of the sparkline is scaled relative to other sparklines in the group.

        [MSDN documentation for SparkVerticalAxis.MinScaleType](https://learn.microsoft.com/en-us/office/vba/api/excel.sparkverticalaxis.minscaletype).aspx).
        */
        void SetMinScaleType(XlSparkScale minScaleType);

        /**
        Returns the parent SparklineGroup object for the specified object.

        [MSDN documentation for SparkVerticalAxis.Parent](https://learn.microsoft.com/en-us/office/vba/api/excel.sparkverticalaxis.parent).aspx).
        */
        wxExcelSparklineGroup GetParent();


        /**
        Returns "SparkVerticalAxis".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("SparkVerticalAxis"); }
    };

    /**
    @brief Represents Microsoft Excel SparkAxes object. Since Excel 2010.
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelSparkAxes: public wxExcelObject
    {
    public:
        // ***** METHODS *****

        // ***** PROPERTIES *****

        /**
        Returns the SparkHorizontalAxis object for the specified SparkAxes object.

        [MSDN documentation for SparkAxes.Horizontal](https://learn.microsoft.com/en-us/office/vba/api/excel.sparkaxes.horizontal).aspx).
        */
        wxExcelSparkHorizontalAxis GetHorizontal();

        /**
        Returns the parent SparklineGroup object for the specified SparkAxes object.

        [MSDN documentation for SparkAxes.Parent](https://learn.microsoft.com/en-us/office/vba/api/excel.sparkaxes.parent).aspx).
        */
        wxExcelSparklineGroup GetParent();

        /**
        Returns the SparkVerticalAxis object for the specified SparkAxes object.

        [MSDN documentation for SparkAxes.Vertical](https://learn.microsoft.com/en-us/office/vba/api/excel.sparkaxes.vertical).aspx).
        */
        wxExcelSparkVerticalAxis GetVertical();


        /**
        Returns "SparkAxes".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("SparkAxes"); }
    };


} // namespace wxAutoExcel

#endif // #if WXAUTOEXCEL_USE_CHARTS

#endif // _WXAUTOEXCEL_SPARKAXES_H
