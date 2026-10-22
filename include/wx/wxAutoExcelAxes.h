/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// License:     MIT license
/////////////////////////////////////////////////////////////////////////////


#ifndef _WXAUTOEXCEL_AXES_H
#define _WXAUTOEXCEL_AXES_H

#include "wx/wxAutoExcel_defs.h"

#if WXAUTOEXCEL_USE_CHARTS

#include "wx/wxAutoExcel_object.h"
#include "wx/wxAutoExcel_enums.h"

class wxArrayString;

namespace wxAutoExcel {


    /**
    @brief Represents Microsoft Excel Axis object.
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelAxis: public wxExcelObject
    {
    public:
        // ***** METHODS *****

        /**
        Deletes the object.

        [MSDN documentation for Axis.Delete](https://learn.microsoft.com/en-us/office/vba/api/excel.axis.delete).
        */
        bool Delete();

        /**
        Selects the object.

        [MSDN documentation for Axis.Select](https://learn.microsoft.com/en-us/office/vba/api/excel.axis.select).
        */
        bool Select();

        // ***** PROPERTIES *****

        /**
        True if the value axis crosses the category axis between categories.

        [MSDN documentation for Axis.AxisBetweenCategories](https://learn.microsoft.com/en-us/office/vba/api/excel.axis.axisbetweencategories).
        */
        bool GetAxisBetweenCategories();

        /**
        True if the value axis crosses the category axis between categories.

        [MSDN documentation for Axis.AxisBetweenCategories](https://learn.microsoft.com/en-us/office/vba/api/excel.axis.axisbetweencategories).
        */
        void SetAxisBetweenCategories(bool axisBetweenCategories);

        /**
        Returns an XlAxisGroup value that represents the type of axis group.

        [MSDN documentation for Axis.AxisGroup](https://learn.microsoft.com/en-us/office/vba/api/excel.axis.axisgroup).
        */
        XlAxisGroup GetAxisGroup();

        /**
        Returns an AxisTitle object that represents the title of the specified axis.

        [MSDN documentation for Axis.AxisTitle](https://learn.microsoft.com/en-us/office/vba/api/excel.axis.axistitle).
        */
        wxExcelAxisTitle GetAxisTitle();

        /**
        Returns the base unit for the specified category axis.

        [MSDN documentation for Axis.BaseUnit](https://learn.microsoft.com/en-us/office/vba/api/excel.axis.baseunit).
        */
        XlTimeUnit GetBaseUnit();

        /**
        Sets the base unit for the specified category axis.

        [MSDN documentation for Axis.BaseUnit](https://learn.microsoft.com/en-us/office/vba/api/excel.axis.baseunit).
        */
        void SetBaseUnit(XlTimeUnit baseUnit);

        /**
        True if Microsoft Excel chooses appropriate base units for the specified category axis. The default value is True.

        [MSDN documentation for Axis.BaseUnitIsAuto](https://learn.microsoft.com/en-us/office/vba/api/excel.axis.baseunitisauto).
        */
        bool GetBaseUnitIsAuto();

        /**
        True if Microsoft Excel chooses appropriate base units for the specified category axis. The default value is True.

        [MSDN documentation for Axis.BaseUnitIsAuto](https://learn.microsoft.com/en-us/office/vba/api/excel.axis.baseunitisauto).
        */
        void SetBaseUnitIsAuto(bool baseUnitIsAuto);

        /**
        Returns a Border object that represents the border of the object.

        [MSDN documentation for Axis.Border](https://learn.microsoft.com/en-us/office/vba/api/excel.axis.border).
        */
        wxExcelBorder GetBorder();

        /**
        Returns all the category names for the specified axis, as a text array. When you set this property, you can set it to either an array or a Range object that contains the category names.

        [MSDN documentation for Axis.CategoryNames](https://learn.microsoft.com/en-us/office/vba/api/excel.axis.categorynames).
        */
        wxArrayString GetCategoryNames();

        //@{
        /**
        Sets all the category names for the specified axis, as a text array. When you set this property, you can set it to either an array or a Range object that contains the category names.

        [MSDN documentation for Axis.CategoryNames](https://learn.microsoft.com/en-us/office/vba/api/excel.axis.categorynames).
        */
        void SetCategoryNames(wxExcelRange categoryNames);
        void SetCategoryNames(const wxArrayString& categoryNames);
        //@}

        /**
        Returns the category axis type.

        [MSDN documentation for Axis.CategoryType](https://learn.microsoft.com/en-us/office/vba/api/excel.axis.categorytype).
        */
        XlCategoryType GetCategoryType();

        /**
        Sets the category axis type.

        [MSDN documentation for Axis.CategoryType](https://learn.microsoft.com/en-us/office/vba/api/excel.axis.categorytype).
        */
        void SetCategoryType(XlCategoryType categoryType);

        /**
        Returns the point on the specified axis where the other axis crosses.

        [MSDN documentation for Axis.Crosses](https://learn.microsoft.com/en-us/office/vba/api/excel.axis.crosses).
        */
        long GetCrosses();

        /**
        Sets the point on the specified axis where the other axis crosses.

        [MSDN documentation for Axis.Crosses](https://learn.microsoft.com/en-us/office/vba/api/excel.axis.crosses).
        */
        void SetCrosses(long crosses);

        /**
        Returns the point on the value axis where the category axis crosses it. Applies only to the value axis.

        [MSDN documentation for Axis.CrossesAt](https://learn.microsoft.com/en-us/office/vba/api/excel.axis.crossesat).
        */
        double GetCrossesAt();

        /**
        Sets the point on the value axis where the category axis crosses it. Applies only to the value axis.

        [MSDN documentation for Axis.CrossesAt](https://learn.microsoft.com/en-us/office/vba/api/excel.axis.crossesat).
        */
        void SetCrossesAt(double crossesAt);

        /**
        Returns the unit label for the value axis. Can be XlDisplayUnit, xlCustom, or xlNone.

        [MSDN documentation for Axis.DisplayUnit](https://learn.microsoft.com/en-us/office/vba/api/excel.axis.displayunit).
        */
        long GetDisplayUnit();

        /**
        Sets the unit label for the value axis. Can be XlDisplayUnit, xlCustom, or xlNone.

        [MSDN documentation for Axis.DisplayUnit](https://learn.microsoft.com/en-us/office/vba/api/excel.axis.displayunit).
        */
        void SetDisplayUnit(long displayUnit);

        /**
        If the value of the DisplayUnit property is xlCustom, the DisplayUnitCustom property returns or sets the value of the displayed units. The value must be from 0 through 10E307. Read/write Double.

        [MSDN documentation for Axis.DisplayUnitCustom](https://learn.microsoft.com/en-us/office/vba/api/excel.axis.displayunitcustom).
        */
        double GetDisplayUnitCustom();

        /**
        If the value of the DisplayUnit property is xlCustom, the DisplayUnitCustom property returns or sets the value of the displayed units. The value must be from 0 through 10E307. Read/write Double.

        [MSDN documentation for Axis.DisplayUnitCustom](https://learn.microsoft.com/en-us/office/vba/api/excel.axis.displayunitcustom).
        */
        void SetDisplayUnitCustom(double displayUnitCustom);

        /**
        Returns the DisplayUnitLabel object for the specified axis. Returns null if the HasDisplayUnitLabel property is set to False.

        [MSDN documentation for Axis.DisplayUnitLabel](https://learn.microsoft.com/en-us/office/vba/api/excel.axis.displayunitlabel).
        */
        wxExcelDisplayUnitLabel GetDisplayUnitLabel();

        /**
        Returns the ChartFormat object. Since Excel 2007.

        [MSDN documentation for Axis.Format](https://learn.microsoft.com/en-us/office/vba/api/excel.axis.format).
        */
        wxExcelChartFormat GetFormat();

        /**
        True if the label specified by the DisplayUnit or DisplayUnitCustom property is displayed on the specified axis. The default value is True.

        [MSDN documentation for Axis.HasDisplayUnitLabel](https://learn.microsoft.com/en-us/office/vba/api/excel.axis.hasdisplayunitlabel).
        */
        bool GetHasDisplayUnitLabel();

        /**
        True if the label specified by the DisplayUnit or DisplayUnitCustom property is displayed on the specified axis. The default value is True.

        [MSDN documentation for Axis.HasDisplayUnitLabel](https://learn.microsoft.com/en-us/office/vba/api/excel.axis.hasdisplayunitlabel).
        */
        void SetHasDisplayUnitLabel(bool hasDisplayUnitLabel);

        /**
        True if the axis has major gridlines. Only axes in the primary axis group can have gridlines.

        [MSDN documentation for Axis.HasMajorGridlines](https://learn.microsoft.com/en-us/office/vba/api/excel.axis.hasmajorgridlines).
        */
        bool GetHasMajorGridlines();

        /**
        True if the axis has major gridlines. Only axes in the primary axis group can have gridlines.

        [MSDN documentation for Axis.HasMajorGridlines](https://learn.microsoft.com/en-us/office/vba/api/excel.axis.hasmajorgridlines).
        */
        void SetHasMajorGridlines(bool hasMajorGridlines);

        /**
        True if the axis has minor gridlines. Only axes in the primary axis group can have gridlines.

        [MSDN documentation for Axis.HasMinorGridlines](https://learn.microsoft.com/en-us/office/vba/api/excel.axis.hasminorgridlines).
        */
        bool GetHasMinorGridlines();

        /**
        True if the axis has minor gridlines. Only axes in the primary axis group can have gridlines.

        [MSDN documentation for Axis.HasMinorGridlines](https://learn.microsoft.com/en-us/office/vba/api/excel.axis.hasminorgridlines).
        */
        void SetHasMinorGridlines(bool hasMinorGridlines);

        /**
        True if the axis or chart has a visible title.

        [MSDN documentation for Axis.HasTitle](https://learn.microsoft.com/en-us/office/vba/api/excel.axis.hastitle).
        */
        bool GetHasTitle();

        /**
        True if the axis or chart has a visible title.

        [MSDN documentation for Axis.HasTitle](https://learn.microsoft.com/en-us/office/vba/api/excel.axis.hastitle).
        */
        void SetHasTitle(bool hasTitle);

        /**
        Returns a Double value that represents the height, in points, of the object.

        [MSDN documentation for Axis.Height](https://learn.microsoft.com/en-us/office/vba/api/excel.axis.height).
        */
        double GetHeight();

        /**
        Returns a Double value that represents the distance, in points, from the left edge of the object to the left edge of the chart area.

        [MSDN documentation for Axis.Left](https://learn.microsoft.com/en-us/office/vba/api/excel.axis.left).
        */
        double GetLeft();

        /**
        Returns the base of the logarithm when you are using log scales. Since Excel 2007.

        [MSDN documentation for Axis.LogBase](https://learn.microsoft.com/en-us/office/vba/api/excel.axis.logbase).
        */
        double GetLogBase();

        /**
        Sets the base of the logarithm when you are using log scales. Since Excel 2007.

        [MSDN documentation for Axis.LogBase](https://learn.microsoft.com/en-us/office/vba/api/excel.axis.logbase).
        */
        void SetLogBase(double logBase);

        /**
        Returns a Gridlines object that represents the major gridlines for the specified axis. Only axes in the primary axis group can have gridlines.

        [MSDN documentation for Axis.MajorGridlines](https://learn.microsoft.com/en-us/office/vba/api/excel.axis.majorgridlines).
        */
        wxExcelGridlines GetMajorGridlines();

        /**
        Returns the type of major tick mark for the specified axis. Read/write XlTickMark.

        [MSDN documentation for Axis.MajorTickMark](https://learn.microsoft.com/en-us/office/vba/api/excel.axis.majortickmark).
        */
        XlTickMark GetMajorTickMark();

        /**
        Sets the type of major tick mark for the specified axis. Read/write XlTickMark.

        [MSDN documentation for Axis.MajorTickMark](https://learn.microsoft.com/en-us/office/vba/api/excel.axis.majortickmark).
        */
        void SetMajorTickMark(XlTickMark majorTickMark);

        /**
        Returns the major units for the value axis. Read/write Double.

        [MSDN documentation for Axis.MajorUnit](https://learn.microsoft.com/en-us/office/vba/api/excel.axis.majorunit).
        */
        double GetMajorUnit();

        /**
        Sets the major units for the value axis. Read/write Double.

        [MSDN documentation for Axis.MajorUnit](https://learn.microsoft.com/en-us/office/vba/api/excel.axis.majorunit).
        */
        void SetMajorUnit(double majorUnit);

        /**
        True if Microsoft Excel calculates the major units for the value axis.

        [MSDN documentation for Axis.MajorUnitIsAuto](https://learn.microsoft.com/en-us/office/vba/api/excel.axis.majorunitisauto).
        */
        bool GetMajorUnitIsAuto();

        /**
        True if Microsoft Excel calculates the major units for the value axis.

        [MSDN documentation for Axis.MajorUnitIsAuto](https://learn.microsoft.com/en-us/office/vba/api/excel.axis.majorunitisauto).
        */
        void SetMajorUnitIsAuto(bool majorUnitIsAuto);

        /**
        Returns the major unit scale value for the category axis when the CategoryType property is set to xlTimeScale. Read/write XlTimeUnit.

        [MSDN documentation for Axis.MajorUnitScale](https://learn.microsoft.com/en-us/office/vba/api/excel.axis.majorunitscale).
        */
        XlTimeUnit GetMajorUnitScale();

        /**
        Sets the major unit scale value for the category axis when the CategoryType property is set to xlTimeScale. Read/write XlTimeUnit.

        [MSDN documentation for Axis.MajorUnitScale](https://learn.microsoft.com/en-us/office/vba/api/excel.axis.majorunitscale).
        */
        void SetMajorUnitScale(XlTimeUnit majorUnitScale);

        /**
        Returns the maximum value on the value axis. Read/write Double.

        [MSDN documentation for Axis.MaximumScale](https://learn.microsoft.com/en-us/office/vba/api/excel.axis.maximumscale).
        */
        double GetMaximumScale();

        /**
        Sets the maximum value on the value axis. Read/write Double.

        [MSDN documentation for Axis.MaximumScale](https://learn.microsoft.com/en-us/office/vba/api/excel.axis.maximumscale).
        */
        void SetMaximumScale(double maximumScale);

        /**
        True if Microsoft Excel calculates the maximum value for the value axis.

        [MSDN documentation for Axis.MaximumScaleIsAuto](https://learn.microsoft.com/en-us/office/vba/api/excel.axis.maximumscaleisauto).
        */
        bool GetMaximumScaleIsAuto();

        /**
        True if Microsoft Excel calculates the maximum value for the value axis.

        [MSDN documentation for Axis.MaximumScaleIsAuto](https://learn.microsoft.com/en-us/office/vba/api/excel.axis.maximumscaleisauto).
        */
        void SetMaximumScaleIsAuto(bool maximumScaleIsAuto);

        /**
        Returns the minimum value on the value axis. Read/write Double.

        [MSDN documentation for Axis.MinimumScale](https://learn.microsoft.com/en-us/office/vba/api/excel.axis.minimumscale).
        */
        double GetMinimumScale();

        /**
        Sets the minimum value on the value axis. Read/write Double.

        [MSDN documentation for Axis.MinimumScale](https://learn.microsoft.com/en-us/office/vba/api/excel.axis.minimumscale).
        */
        void SetMinimumScale(double minimumScale);

        /**
        True if Microsoft Excel calculates the minimum value for the value axis.

        [MSDN documentation for Axis.MinimumScaleIsAuto](https://learn.microsoft.com/en-us/office/vba/api/excel.axis.minimumscaleisauto).
        */
        bool GetMinimumScaleIsAuto();

        /**
        True if Microsoft Excel calculates the minimum value for the value axis.

        [MSDN documentation for Axis.MinimumScaleIsAuto](https://learn.microsoft.com/en-us/office/vba/api/excel.axis.minimumscaleisauto).
        */
        void SetMinimumScaleIsAuto(bool minimumScaleIsAuto);

        /**
        Returns a Gridlines object that represents the minor gridlines for the specified axis. Only axes in the primary axis group can have gridlines.

        [MSDN documentation for Axis.MinorGridlines](https://learn.microsoft.com/en-us/office/vba/api/excel.axis.minorgridlines).
        */
        wxExcelGridlines GetMinorGridlines();

        /**
        Returns the type of minor tick mark for the specified axis. Read/write XlTickMark.

        [MSDN documentation for Axis.MinorTickMark](https://learn.microsoft.com/en-us/office/vba/api/excel.axis.minortickmark).
        */
        XlTickMark GetMinorTickMark();

        /**
        Sets the type of minor tick mark for the specified axis. Read/write XlTickMark.

        [MSDN documentation for Axis.MinorTickMark](https://learn.microsoft.com/en-us/office/vba/api/excel.axis.minortickmark).
        */
        void SetMinorTickMark(XlTickMark minorTickMark);

        /**
        Returns the minor units on the value axis. Read/write Double.

        [MSDN documentation for Axis.MinorUnit](https://learn.microsoft.com/en-us/office/vba/api/excel.axis.minorunit).
        */
        double GetMinorUnit();

        /**
        Sets the minor units on the value axis. Read/write Double.

        [MSDN documentation for Axis.MinorUnit](https://learn.microsoft.com/en-us/office/vba/api/excel.axis.minorunit).
        */
        void SetMinorUnit(double minorUnit);

        /**
        True if Microsoft Excel calculates minor units for the value axis.

        [MSDN documentation for Axis.MinorUnitIsAuto](https://learn.microsoft.com/en-us/office/vba/api/excel.axis.minorunitisauto).
        */
        bool GetMinorUnitIsAuto();

        /**
        True if Microsoft Excel calculates minor units for the value axis.

        [MSDN documentation for Axis.MinorUnitIsAuto](https://learn.microsoft.com/en-us/office/vba/api/excel.axis.minorunitisauto).
        */
        void SetMinorUnitIsAuto(bool minorUnitIsAuto);

        /**
        Returns the minor unit scale value for the category axis when the CategoryType property is set to xlTimeScale. Read/write XlTimeUnit.

        [MSDN documentation for Axis.MinorUnitScale](https://learn.microsoft.com/en-us/office/vba/api/excel.axis.minorunitscale).
        */
        XlTimeUnit GetMinorUnitScale();

        /**
        Sets the minor unit scale value for the category axis when the CategoryType property is set to xlTimeScale. Read/write XlTimeUnit.

        [MSDN documentation for Axis.MinorUnitScale](https://learn.microsoft.com/en-us/office/vba/api/excel.axis.minorunitscale).
        */
        void SetMinorUnitScale(XlTimeUnit minorUnitScale);

        /**
        True if Microsoft Excel plots data points from last to first.

        [MSDN documentation for Axis.ReversePlotOrder](https://learn.microsoft.com/en-us/office/vba/api/excel.axis.reverseplotorder).
        */
        bool GetReversePlotOrder();

        /**
        True if Microsoft Excel plots data points from last to first.

        [MSDN documentation for Axis.ReversePlotOrder](https://learn.microsoft.com/en-us/office/vba/api/excel.axis.reverseplotorder).
        */
        void SetReversePlotOrder(bool reversePlotOrder);

        /**
        Returns the value axis scale type. Read/write XlScaleType.

        [MSDN documentation for Axis.ScaleType](https://learn.microsoft.com/en-us/office/vba/api/excel.axis.scaletype).
        */
        XlScaleType GetScaleType();

        /**
        Sets the value axis scale type. Read/write XlScaleType.

        [MSDN documentation for Axis.ScaleType](https://learn.microsoft.com/en-us/office/vba/api/excel.axis.scaletype).
        */
        void SetScaleType(XlScaleType scaleType);

        /**
        Describes the position of tick-mark labels on the specified axis. Read/write XlTickLabelPosition.

        [MSDN documentation for Axis.TickLabelPosition](https://learn.microsoft.com/en-us/office/vba/api/excel.axis.ticklabelposition).
        */
        XlTickLabelPosition GetTickLabelPosition();

        /**
        Describes the position of tick-mark labels on the specified axis. Read/write XlTickLabelPosition.

        [MSDN documentation for Axis.TickLabelPosition](https://learn.microsoft.com/en-us/office/vba/api/excel.axis.ticklabelposition).
        */
        void SetTickLabelPosition(XlTickLabelPosition tickLabelPosition);

        /**
        Returns a TickLabels object that represents the tick-mark labels for the specified axis.

        [MSDN documentation for Axis.TickLabels](https://learn.microsoft.com/en-us/office/vba/api/excel.axis.ticklabels).
        */
        wxExcelTickLabels GetTickLabels();

        /**
        Returns the number of categories or series between tick-mark labels. Applies only to category and series axes. Can be a value from 1 through 31999.

        [MSDN documentation for Axis.TickLabelSpacing](https://learn.microsoft.com/en-us/office/vba/api/excel.axis.ticklabelspacing).
        */
        long GetTickLabelSpacing();

        /**
        Sets the number of categories or series between tick-mark labels. Applies only to category and series axes. Can be a value from 1 through 31999.

        [MSDN documentation for Axis.TickLabelSpacing](https://learn.microsoft.com/en-us/office/vba/api/excel.axis.ticklabelspacing).
        */
        void SetTickLabelSpacing(long tickLabelSpacing);

        /**
        Returns whether or not the tick label spacing is automatic.  Since Excel 2007.

        [MSDN documentation for Axis.TickLabelSpacingIsAuto](https://learn.microsoft.com/en-us/office/vba/api/excel.axis.ticklabelspacingisauto).
        */
        bool GetTickLabelSpacingIsAuto();

        /**
        Sets whether or not the tick label spacing is automatic.  Since Excel 2007.

        [MSDN documentation for Axis.TickLabelSpacingIsAuto](https://learn.microsoft.com/en-us/office/vba/api/excel.axis.ticklabelspacingisauto).
        */
        void SetTickLabelSpacingIsAuto(bool tickLabelSpacingIsAuto);

        /**
        Returns the number of categories or series between tick marks. Applies only to category and series axes. Can be a value from 1 through 31999.

        [MSDN documentation for Axis.TickMarkSpacing](https://learn.microsoft.com/en-us/office/vba/api/excel.axis.tickmarkspacing).
        */
        long GetTickMarkSpacing();

        /**
        Sets the number of categories or series between tick marks. Applies only to category and series axes. Can be a value from 1 through 31999.

        [MSDN documentation for Axis.TickMarkSpacing](https://learn.microsoft.com/en-us/office/vba/api/excel.axis.tickmarkspacing).
        */
        void SetTickMarkSpacing(long tickMarkSpacing);

        /**
        Returns a Double value that represents the distance, in points, from the top edge of the object to the top of row 1 (on a worksheet) or the top of the chart area (on a chart).

        [MSDN documentation for Axis.Top](https://learn.microsoft.com/en-us/office/vba/api/excel.axis.top).
        */
        double GetTop();

        /**
        Returns an XlAxisType value that represents the Axis type.

        [MSDN documentation for Axis.Type](https://learn.microsoft.com/en-us/office/vba/api/excel.axis.type).
        */
        XlAxisType GetType();

        /**
        Returns a Double value that represents the width, in points, of the object.

        [MSDN documentation for Axis.Width](https://learn.microsoft.com/en-us/office/vba/api/excel.axis.width).
        */
        double GetWidth();

        /**
        Returns "Axis".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("Axis"); }
    };

    /**
    @brief Represents Microsoft Excel Axes collection.
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelAxes: public wxExcelObject
    {
    public:
        // ***** METHODS *****

        //@{
        /**
        Returns a single Axis object from an Axes collection.

        [MSDN documentation for Axes.Item](https://learn.microsoft.com/en-us/office/vba/api/excel.axes.item).
        */
        wxExcelAxis Item(long index);
        wxExcelAxis operator[](long index);
        //@}

        // ***** PROPERTIES *****

        /**
        Returns a Long value that represents the number of objects in the collection.

        [MSDN documentation for Axes.Count](https://learn.microsoft.com/en-us/office/vba/api/excel.axes.count).
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
