/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Modified by:
// RCS-ID:      $Id: $
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////


#ifndef _WXAUTOEXCEL_AXES_H
#define _WXAUTOEXCEL_AXES_H

#include "wx/wxAutoExcel_defs.h"

#if WXAUTOEXCEL_USE_CHARTS

#include "wx/wxAutoExcelObject.h"
#include "wx/wxAutoExcel_enums.h"

class wxArrayString;

namespace wxAutoExcel {


    /**
    Represents Microsoft Excel Axis object.
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelAxis: public wxExcelObject
    {
    public:
        // ***** METHODS *****

        /**
        Deletes the object.

        [MSDN documentation for Axis.Delete](http://msdn.microsoft.com/en-us/library/bb211573).
        */
        bool Delete();

        /**
        Selects the object.

        [MSDN documentation for Axis.Select](http://msdn.microsoft.com/en-us/library/bb213854).
        */
        bool Select();

        // ***** PROPERTIES *****

        /**
        True if the value axis crosses the category axis between categories.

        [MSDN documentation for Axis.AxisBetweenCategories](http://msdn.microsoft.com/en-us/library/bb220870).
        */
        bool GetAxisBetweenCategories();

        /**
        True if the value axis crosses the category axis between categories.

        [MSDN documentation for Axis.AxisBetweenCategories](http://msdn.microsoft.com/en-us/library/bb220870).
        */
        void SetAxisBetweenCategories(bool axisBetweenCategories);

        /**
        Returns an XlAxisGroup value that represents the type of axis group.

        [MSDN documentation for Axis.AxisGroup](http://msdn.microsoft.com/en-us/library/dd787723).
        */
        XlAxisGroup GetAxisGroup();

        /**
        Returns an AxisTitle object that represents the title of the specified axis.

        [MSDN documentation for Axis.AxisTitle](http://msdn.microsoft.com/en-us/library/bb220871).
        */
        wxExcelAxisTitle GetAxisTitle();

        /**
        Returns the base unit for the specified category axis.

        [MSDN documentation for Axis.BaseUnit](http://msdn.microsoft.com/en-us/library/bb220880).
        */
        XlTimeUnit GetBaseUnit();

        /**
        Sets the base unit for the specified category axis. 

        [MSDN documentation for Axis.BaseUnit](http://msdn.microsoft.com/en-us/library/bb220880).
        */
        void SetBaseUnit(XlTimeUnit baseUnit);

        /**
        True if Microsoft Excel chooses appropriate base units for the specified category axis. The default value is True.

        [MSDN documentation for Axis.BaseUnitIsAuto](http://msdn.microsoft.com/en-us/library/bb220881).
        */
        bool GetBaseUnitIsAuto();

        /**
        True if Microsoft Excel chooses appropriate base units for the specified category axis. The default value is True.

        [MSDN documentation for Axis.BaseUnitIsAuto](http://msdn.microsoft.com/en-us/library/bb220881).
        */
        void SetBaseUnitIsAuto(bool baseUnitIsAuto);

        /**
        Returns a Border object that represents the border of the object.

        [MSDN documentation for Axis.Border](http://msdn.microsoft.com/en-us/library/bb179279).
        */
        wxExcelBorder GetBorder();

        /**
        Returns all the category names for the specified axis, as a text array. When you set this property, you can set it to either an array or a Range object that contains the category names. 

        [MSDN documentation for Axis.CategoryNames](http://msdn.microsoft.com/en-us/library/bb220908).
        */
        wxArrayString GetCategoryNames();

        //@{
        /**
        Sets all the category names for the specified axis, as a text array. When you set this property, you can set it to either an array or a Range object that contains the category names. 

        [MSDN documentation for Axis.CategoryNames](http://msdn.microsoft.com/en-us/library/bb220908).
        */
        void SetCategoryNames(wxExcelRange categoryNames);
        void SetCategoryNames(const wxArrayString& categoryNames);
        //@}

        /**
        Returns the category axis type. 

        [MSDN documentation for Axis.CategoryType](http://msdn.microsoft.com/en-us/library/bb220909).
        */
        XlCategoryType GetCategoryType();

        /**
        Sets the category axis type. 

        [MSDN documentation for Axis.CategoryType](http://msdn.microsoft.com/en-us/library/bb220909).
        */
        void SetCategoryType(XlCategoryType categoryType);

        /**
        Returns the point on the specified axis where the other axis crosses.

        [MSDN documentation for Axis.Crosses](http://msdn.microsoft.com/en-us/library/bb177405).
        */
        long GetCrosses();

        /**
        Sets the point on the specified axis where the other axis crosses.

        [MSDN documentation for Axis.Crosses](http://msdn.microsoft.com/en-us/library/bb177405).
        */
        void SetCrosses(long crosses);

        /**
        Returns the point on the value axis where the category axis crosses it. Applies only to the value axis. 

        [MSDN documentation for Axis.CrossesAt](http://msdn.microsoft.com/en-us/library/bb177407).
        */
        double GetCrossesAt();

        /**
        Sets the point on the value axis where the category axis crosses it. Applies only to the value axis. 

        [MSDN documentation for Axis.CrossesAt](http://msdn.microsoft.com/en-us/library/bb177407).
        */
        void SetCrossesAt(double crossesAt);

        /**
        Returns the unit label for the value axis. Can be XlDisplayUnit, xlCustom, or xlNone.

        [MSDN documentation for Axis.DisplayUnit](http://msdn.microsoft.com/en-us/library/bb221003).
        */
        long GetDisplayUnit();

        /**
        Sets the unit label for the value axis. Can be XlDisplayUnit, xlCustom, or xlNone.

        [MSDN documentation for Axis.DisplayUnit](http://msdn.microsoft.com/en-us/library/bb221003).
        */
        void SetDisplayUnit(long displayUnit);

        /**
        If the value of the DisplayUnit property is xlCustom, the DisplayUnitCustom property returns or sets the value of the displayed units. The value must be from 0 through 10E307. Read/write Double.

        [MSDN documentation for Axis.DisplayUnitCustom](http://msdn.microsoft.com/en-us/library/bb221009).
        */
        double GetDisplayUnitCustom();

        /**
        If the value of the DisplayUnit property is xlCustom, the DisplayUnitCustom property returns or sets the value of the displayed units. The value must be from 0 through 10E307. Read/write Double.

        [MSDN documentation for Axis.DisplayUnitCustom](http://msdn.microsoft.com/en-us/library/bb221009).
        */
        void SetDisplayUnitCustom(double displayUnitCustom);

        /**
        Returns the DisplayUnitLabel object for the specified axis. Returns null if the HasDisplayUnitLabel property is set to False.

        [MSDN documentation for Axis.DisplayUnitLabel](http://msdn.microsoft.com/en-us/library/bb221013).
        */
        wxExcelDisplayUnitLabel GetDisplayUnitLabel();

        /**
        Returns the ChartFormat object. Since Excel 2007.

        [MSDN documentation for Axis.Format](http://msdn.microsoft.com/en-us/library/bb242488).
        */
        wxExcelChartFormat GetFormat();

        /**
        True if the label specified by the DisplayUnit or DisplayUnitCustom property is displayed on the specified axis. The default value is True.

        [MSDN documentation for Axis.HasDisplayUnitLabel](http://msdn.microsoft.com/en-us/library/bb208614).
        */
        bool GetHasDisplayUnitLabel();

        /**
        True if the label specified by the DisplayUnit or DisplayUnitCustom property is displayed on the specified axis. The default value is True.

        [MSDN documentation for Axis.HasDisplayUnitLabel](http://msdn.microsoft.com/en-us/library/bb208614).
        */
        void SetHasDisplayUnitLabel(bool hasDisplayUnitLabel);

        /**
        True if the axis has major gridlines. Only axes in the primary axis group can have gridlines.

        [MSDN documentation for Axis.HasMajorGridlines](http://msdn.microsoft.com/en-us/library/bb208640).
        */
        bool GetHasMajorGridlines();

        /**
        True if the axis has major gridlines. Only axes in the primary axis group can have gridlines.

        [MSDN documentation for Axis.HasMajorGridlines](http://msdn.microsoft.com/en-us/library/bb208640).
        */
        void SetHasMajorGridlines(bool hasMajorGridlines);

        /**
        True if the axis has minor gridlines. Only axes in the primary axis group can have gridlines.

        [MSDN documentation for Axis.HasMinorGridlines](http://msdn.microsoft.com/en-us/library/bb208647).
        */
        bool GetHasMinorGridlines();

        /**
        True if the axis has minor gridlines. Only axes in the primary axis group can have gridlines.

        [MSDN documentation for Axis.HasMinorGridlines](http://msdn.microsoft.com/en-us/library/bb208647).
        */
        void SetHasMinorGridlines(bool hasMinorGridlines);

        /**
        True if the axis or chart has a visible title.

        [MSDN documentation for Axis.HasTitle](http://msdn.microsoft.com/en-us/library/bb179281).
        */
        bool GetHasTitle();

        /**
        True if the axis or chart has a visible title.

        [MSDN documentation for Axis.HasTitle](http://msdn.microsoft.com/en-us/library/bb179281).
        */
        void SetHasTitle(bool hasTitle);

        /**
        Returns a Double value that represents the height, in points, of the object.

        [MSDN documentation for Axis.Height](http://msdn.microsoft.com/en-us/library/bb179284).
        */
        double GetHeight();

        /**
        Returns a Double value that represents the distance, in points, from the left edge of the object to the left edge of the chart area.

        [MSDN documentation for Axis.Left](http://msdn.microsoft.com/en-us/library/bb179285).
        */
        double GetLeft();

        /**
        Returns the base of the logarithm when you are using log scales. Since Excel 2007.

        [MSDN documentation for Axis.LogBase](http://msdn.microsoft.com/en-us/library/bb224816).
        */
        double GetLogBase();

        /**
        Sets the base of the logarithm when you are using log scales. Since Excel 2007.

        [MSDN documentation for Axis.LogBase](http://msdn.microsoft.com/en-us/library/bb224816).
        */
        void SetLogBase(double logBase);

        /**
        Returns a Gridlines object that represents the major gridlines for the specified axis. Only axes in the primary axis group can have gridlines.

        [MSDN documentation for Axis.MajorGridlines](http://msdn.microsoft.com/en-us/library/bb208711).
        */
        wxExcelGridlines GetMajorGridlines();

        /**
        Returns the type of major tick mark for the specified axis. Read/write XlTickMark.

        [MSDN documentation for Axis.MajorTickMark](http://msdn.microsoft.com/en-us/library/bb208713).
        */
        XlTickMark GetMajorTickMark();

        /**
        Sets the type of major tick mark for the specified axis. Read/write XlTickMark.

        [MSDN documentation for Axis.MajorTickMark](http://msdn.microsoft.com/en-us/library/bb208713).
        */
        void SetMajorTickMark(XlTickMark majorTickMark);

        /**
        Returns the major units for the value axis. Read/write Double.

        [MSDN documentation for Axis.MajorUnit](http://msdn.microsoft.com/en-us/library/bb208716).
        */
        double GetMajorUnit();

        /**
        Sets the major units for the value axis. Read/write Double.

        [MSDN documentation for Axis.MajorUnit](http://msdn.microsoft.com/en-us/library/bb208716).
        */
        void SetMajorUnit(double majorUnit);

        /**
        True if Microsoft Excel calculates the major units for the value axis.

        [MSDN documentation for Axis.MajorUnitIsAuto](http://msdn.microsoft.com/en-us/library/bb208717).
        */
        bool GetMajorUnitIsAuto();

        /**
        True if Microsoft Excel calculates the major units for the value axis.

        [MSDN documentation for Axis.MajorUnitIsAuto](http://msdn.microsoft.com/en-us/library/bb208717).
        */
        void SetMajorUnitIsAuto(bool majorUnitIsAuto);

        /**
        Returns the major unit scale value for the category axis when the CategoryType property is set to xlTimeScale. Read/write XlTimeUnit.

        [MSDN documentation for Axis.MajorUnitScale](http://msdn.microsoft.com/en-us/library/bb208719).
        */
        XlTimeUnit GetMajorUnitScale();

        /**
        Sets the major unit scale value for the category axis when the CategoryType property is set to xlTimeScale. Read/write XlTimeUnit.

        [MSDN documentation for Axis.MajorUnitScale](http://msdn.microsoft.com/en-us/library/bb208719).
        */
        void SetMajorUnitScale(XlTimeUnit majorUnitScale);

        /**
        Returns the maximum value on the value axis. Read/write Double.

        [MSDN documentation for Axis.MaximumScale](http://msdn.microsoft.com/en-us/library/bb208744).
        */
        double GetMaximumScale();

        /**
        Sets the maximum value on the value axis. Read/write Double.

        [MSDN documentation for Axis.MaximumScale](http://msdn.microsoft.com/en-us/library/bb208744).
        */
        void SetMaximumScale(double maximumScale);

        /**
        True if Microsoft Excel calculates the maximum value for the value axis.

        [MSDN documentation for Axis.MaximumScaleIsAuto](http://msdn.microsoft.com/en-us/library/bb208747).
        */
        bool GetMaximumScaleIsAuto();

        /**
        True if Microsoft Excel calculates the maximum value for the value axis.

        [MSDN documentation for Axis.MaximumScaleIsAuto](http://msdn.microsoft.com/en-us/library/bb208747).
        */
        void SetMaximumScaleIsAuto(bool maximumScaleIsAuto);

        /**
        Returns the minimum value on the value axis. Read/write Double.

        [MSDN documentation for Axis.MinimumScale](http://msdn.microsoft.com/en-us/library/bb208764).
        */
        double GetMinimumScale();

        /**
        Sets the minimum value on the value axis. Read/write Double.

        [MSDN documentation for Axis.MinimumScale](http://msdn.microsoft.com/en-us/library/bb208764).
        */
        void SetMinimumScale(double minimumScale);

        /**
        True if Microsoft Excel calculates the minimum value for the value axis.

        [MSDN documentation for Axis.MinimumScaleIsAuto](http://msdn.microsoft.com/en-us/library/bb208766).
        */
        bool GetMinimumScaleIsAuto();

        /**
        True if Microsoft Excel calculates the minimum value for the value axis.

        [MSDN documentation for Axis.MinimumScaleIsAuto](http://msdn.microsoft.com/en-us/library/bb208766).
        */
        void SetMinimumScaleIsAuto(bool minimumScaleIsAuto);

        /**
        Returns a Gridlines object that represents the minor gridlines for the specified axis. Only axes in the primary axis group can have gridlines.

        [MSDN documentation for Axis.MinorGridlines](http://msdn.microsoft.com/en-us/library/bb208773).
        */
        wxExcelGridlines GetMinorGridlines();

        /**
        Returns the type of minor tick mark for the specified axis. Read/write XlTickMark.

        [MSDN documentation for Axis.MinorTickMark](http://msdn.microsoft.com/en-us/library/bb208776).
        */
        XlTickMark GetMinorTickMark();

        /**
        Sets the type of minor tick mark for the specified axis. Read/write XlTickMark.

        [MSDN documentation for Axis.MinorTickMark](http://msdn.microsoft.com/en-us/library/bb208776).
        */
        void SetMinorTickMark(XlTickMark minorTickMark);

        /**
        Returns the minor units on the value axis. Read/write Double.

        [MSDN documentation for Axis.MinorUnit](http://msdn.microsoft.com/en-us/library/bb208781).
        */
        double GetMinorUnit();

        /**
        Sets the minor units on the value axis. Read/write Double.

        [MSDN documentation for Axis.MinorUnit](http://msdn.microsoft.com/en-us/library/bb208781).
        */
        void SetMinorUnit(double minorUnit);

        /**
        True if Microsoft Excel calculates minor units for the value axis.

        [MSDN documentation for Axis.MinorUnitIsAuto](http://msdn.microsoft.com/en-us/library/bb208784).
        */
        bool GetMinorUnitIsAuto();

        /**
        True if Microsoft Excel calculates minor units for the value axis.

        [MSDN documentation for Axis.MinorUnitIsAuto](http://msdn.microsoft.com/en-us/library/bb208784).
        */
        void SetMinorUnitIsAuto(bool minorUnitIsAuto);

        /**
        Returns the minor unit scale value for the category axis when the CategoryType property is set to xlTimeScale. Read/write XlTimeUnit.

        [MSDN documentation for Axis.MinorUnitScale](http://msdn.microsoft.com/en-us/library/bb208786).
        */
        XlTimeUnit GetMinorUnitScale();

        /**
        Sets the minor unit scale value for the category axis when the CategoryType property is set to xlTimeScale. Read/write XlTimeUnit.

        [MSDN documentation for Axis.MinorUnitScale](http://msdn.microsoft.com/en-us/library/bb208786).
        */
        void SetMinorUnitScale(XlTimeUnit minorUnitScale);        

        /**
        True if Microsoft Excel plots data points from last to first.

        [MSDN documentation for Axis.ReversePlotOrder](http://msdn.microsoft.com/en-us/library/bb209149).
        */
        bool GetReversePlotOrder();

        /**
        True if Microsoft Excel plots data points from last to first.

        [MSDN documentation for Axis.ReversePlotOrder](http://msdn.microsoft.com/en-us/library/bb209149).
        */
        void SetReversePlotOrder(bool reversePlotOrder);

        /**
        Returns the value axis scale type. Read/write XlScaleType.

        [MSDN documentation for Axis.ScaleType](http://msdn.microsoft.com/en-us/library/bb221595).
        */
        XlScaleType GetScaleType();

        /**
        Sets the value axis scale type. Read/write XlScaleType.

        [MSDN documentation for Axis.ScaleType](http://msdn.microsoft.com/en-us/library/bb221595).
        */
        void SetScaleType(XlScaleType scaleType);

        /**
        Describes the position of tick-mark labels on the specified axis. Read/write XlTickLabelPosition.

        [MSDN documentation for Axis.TickLabelPosition](http://msdn.microsoft.com/en-us/library/bb221823).
        */
        XlTickLabelPosition GetTickLabelPosition();

        /**
        Describes the position of tick-mark labels on the specified axis. Read/write XlTickLabelPosition.

        [MSDN documentation for Axis.TickLabelPosition](http://msdn.microsoft.com/en-us/library/bb221823).
        */
        void SetTickLabelPosition(XlTickLabelPosition tickLabelPosition);

        /**
        Returns a TickLabels object that represents the tick-mark labels for the specified axis.

        [MSDN documentation for Axis.TickLabels](http://msdn.microsoft.com/en-us/library/bb221829).
        */
        wxExcelTickLabels GetTickLabels();

        /**
        Returns the number of categories or series between tick-mark labels. Applies only to category and series axes. Can be a value from 1 through 31999.

        [MSDN documentation for Axis.TickLabelSpacing](http://msdn.microsoft.com/en-us/library/bb221833).
        */
        long GetTickLabelSpacing();

        /**
        Sets the number of categories or series between tick-mark labels. Applies only to category and series axes. Can be a value from 1 through 31999.

        [MSDN documentation for Axis.TickLabelSpacing](http://msdn.microsoft.com/en-us/library/bb221833).
        */
        void SetTickLabelSpacing(long tickLabelSpacing);

        /**
        Returns whether or not the tick label spacing is automatic.  Since Excel 2007.

        [MSDN documentation for Axis.TickLabelSpacingIsAuto](http://msdn.microsoft.com/en-us/library/bb224818).
        */
        bool GetTickLabelSpacingIsAuto();

        /**
        Sets whether or not the tick label spacing is automatic.  Since Excel 2007.

        [MSDN documentation for Axis.TickLabelSpacingIsAuto](http://msdn.microsoft.com/en-us/library/bb224818).
        */
        void SetTickLabelSpacingIsAuto(bool tickLabelSpacingIsAuto);

        /**
        Returns the number of categories or series between tick marks. Applies only to category and series axes. Can be a value from 1 through 31999.

        [MSDN documentation for Axis.TickMarkSpacing](http://msdn.microsoft.com/en-us/library/bb221838).
        */
        long GetTickMarkSpacing();

        /**
        Sets the number of categories or series between tick marks. Applies only to category and series axes. Can be a value from 1 through 31999.

        [MSDN documentation for Axis.TickMarkSpacing](http://msdn.microsoft.com/en-us/library/bb221838).
        */
        void SetTickMarkSpacing(long tickMarkSpacing);

        /**
        Returns a Double value that represents the distance, in points, from the top edge of the object to the top of row 1 (on a worksheet) or the top of the chart area (on a chart).

        [MSDN documentation for Axis.Top](http://msdn.microsoft.com/en-us/library/bb148697).
        */
        double GetTop();

        /**
        Returns an XlAxisType value that represents the Axis type.

        [MSDN documentation for Axis.Type](http://msdn.microsoft.com/en-us/library/bb148702).
        */
        XlAxisType GetType();

        /**
        Returns a Double value that represents the width, in points, of the object.

        [MSDN documentation for Axis.Width](http://msdn.microsoft.com/en-us/library/bb148704).
        */
        double GetWidth();

        /**
        Returns "Axis".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("Axis"); }
    };

    /**
    Represents Microsoft Excel Axes collection.
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelAxes: public wxExcelObject
    {
    public:
        // ***** METHODS *****

        //@{
        /**
        Returns a single Axis object from an Axes collection.

        [MSDN documentation for Axes.Item](http://msdn.microsoft.com/en-us/library/bb211568).
        */
        wxExcelAxis Item(long index);
        wxExcelAxis operator[](long index);
        //@}

        // ***** PROPERTIES *****

        /**
        Returns a Long value that represents the number of objects in the collection.

        [MSDN documentation for Axes.Count](http://msdn.microsoft.com/en-us/library/bb179278).
        */
        long GetCount();


        /**
        Returns "Axes".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("Axes"); }
    };


} // namespace wxAutoExcel

#endif // #if WXAUTOEXCEL_USE_CHARTS

#endif // _WXAUTOEXCEL_AXES_H
