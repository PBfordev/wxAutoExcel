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

        [MSDN documentation for SparkHorizontalAxis.Axis](http://msdn.microsoft.com/en-us/library/office/ff821267(v=office.14).aspx).
        */
        wxExcelSparkColor GetAxis();

        /**
        Returns whether the horizontal axis of the sparkline is based on date values.

        [MSDN documentation for SparkHorizontalAxis.IsDateAxis](http://msdn.microsoft.com/en-us/library/office/ff840524(v=office.14).aspx).
        */
        bool GetIsDateAxis();

        /**
        Returns the parent SparklineGroup object for the specified object.

        [MSDN documentation for SparkHorizontalAxis.Parent](http://msdn.microsoft.com/en-us/library/office/ff196240(v=office.14).aspx).
        */
        wxExcelSparklineGroup GetParent();

        /**
        Returns whether the points on the horizontal axis are plotted in right-to-left order.

        [MSDN documentation for SparkHorizontalAxis.RightToLeftPlotOrder](http://msdn.microsoft.com/en-us/library/office/ff823107(v=office.14).aspx).
        */
        bool GetRightToLeftPlotOrder();


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

        [MSDN documentation for SparkVerticalAxis.CustomMaxScaleValue](http://msdn.microsoft.com/en-us/library/office/ff821960(v=office.14).aspx).
        */
        double GetCustomMaxScaleValue();

        /**
        Sets the custom maximum value for the vertical axis of a sparkline.

        [MSDN documentation for SparkVerticalAxis.CustomMaxScaleValue](http://msdn.microsoft.com/en-us/library/office/ff821960(v=office.14).aspx).
        */
        void SetCustomMaxScaleValue(double customMaxScaleValue);

        /**
        Returns the custom minimum value for the vertical axis of a sparkline.

        [MSDN documentation for SparkVerticalAxis.CustomMinScaleValue](http://msdn.microsoft.com/en-us/library/office/ff194837(v=office.14).aspx).
        */
        double GetCustomMinScaleValue();

        /**
        Sets the custom minimum value for the vertical axis of a sparkline.

        [MSDN documentation for SparkVerticalAxis.CustomMinScaleValue](http://msdn.microsoft.com/en-us/library/office/ff194837(v=office.14).aspx).
        */
        void SetCustomMinScaleValue(double customMinScaleValue);

        /**
        Returns how the maximum value of the vertical axis of the sparkline is scaled relative to other sparklines in the group.

        [MSDN documentation for SparkVerticalAxis.MaxScaleType](http://msdn.microsoft.com/en-us/library/office/ff194274(v=office.14).aspx).
        */
        XlSparkScale GetMaxScaleType();

        /**
        Sets how the maximum value of the vertical axis of the sparkline is scaled relative to other sparklines in the group.

        [MSDN documentation for SparkVerticalAxis.MaxScaleType](http://msdn.microsoft.com/en-us/library/office/ff194274(v=office.14).aspx).
        */
        void SetMaxScaleType(XlSparkScale maxScaleType);

        /**
        Returns how the minimum value of the vertical axis of the sparkline is scaled relative to other sparklines in the group.

        [MSDN documentation for SparkVerticalAxis.MinScaleType](http://msdn.microsoft.com/en-us/library/office/ff840224(v=office.14).aspx).
        */
        XlSparkScale GetMinScaleType();

        /**
        Sets how the minimum value of the vertical axis of the sparkline is scaled relative to other sparklines in the group.

        [MSDN documentation for SparkVerticalAxis.MinScaleType](http://msdn.microsoft.com/en-us/library/office/ff840224(v=office.14).aspx).
        */
        void SetMinScaleType(XlSparkScale minScaleType);

        /**
        Returns the parent SparklineGroup object for the specified object.

        [MSDN documentation for SparkVerticalAxis.Parent](http://msdn.microsoft.com/en-us/library/office/ff837439(v=office.14).aspx).
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

        [MSDN documentation for SparkAxes.Horizontal](http://msdn.microsoft.com/en-us/library/office/ff823093(v=office.14).aspx).
        */
        wxExcelSparkHorizontalAxis GetHorizontal();

        /**
        Returns the parent SparklineGroup object for the specified SparkAxes object.

        [MSDN documentation for SparkAxes.Parent](http://msdn.microsoft.com/en-us/library/office/ff198368(v=office.14).aspx).
        */
        wxExcelSparklineGroup GetParent();

        /**
        Returns the SparkVerticalAxis object for the specified SparkAxes object.

        [MSDN documentation for SparkAxes.Vertical](http://msdn.microsoft.com/en-us/library/office/ff841183(v=office.14).aspx).
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
