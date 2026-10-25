/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// License:     MIT license
/////////////////////////////////////////////////////////////////////////////


#ifndef _WXAUTOEXCEL_SERIES_H
#define _WXAUTOEXCEL_SERIES_H

#include "wx/wxAutoExcel_defs.h"

#if WXAUTOEXCEL_USE_CHARTS

#include "wx/wxAutoExcel_object.h"
#include "wx/wxAutoExcel_enums.h"

namespace wxAutoExcel {

    /**
    @brief Represents Microsoft Excel Series object.
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelSeries : public wxExcelObject
    {
    public:
        // ***** METHODS *****

        /**
        Applies data labels to a series.

        [MSDN documentation for Series.ApplyDataLabels](https://learn.microsoft.com/en-us/office/vba/api/excel.series.applydatalabels).
        */
        void ApplyDataLabels(XlDataLabelsType* type = NULL, wxXlTribool legendKey = wxDefaultXlTribool,
                             wxXlTribool autoText = wxDefaultXlTribool, wxXlTribool hasLeaderLines = wxDefaultXlTribool,
                             wxXlTribool showSeriesName = wxDefaultXlTribool, wxXlTribool showCategoryName = wxDefaultXlTribool,
                             wxXlTribool showValue = wxDefaultXlTribool, wxXlTribool showPercentage = wxDefaultXlTribool,
                             wxXlTribool showBubbleSize = wxDefaultXlTribool, const wxString& separator = wxEmptyString);

        /**
        Clears the formatting of the object.

        [MSDN documentation for Series.ClearFormats](https://learn.microsoft.com/en-us/office/vba/api/excel.series.clearformats).
        */
        bool ClearFormats();

        /**
        If the series has a picture fill, then this method copies the picture to the Clipboard.

        [MSDN documentation for Series.Copy](https://learn.microsoft.com/en-us/office/vba/api/excel.series.copy).
        */
        bool Copy();

        /**
        Returns an object that represents either a single data label (a DataLabel object) or a collection of all the data labels for the series (a DataLabels collection).

        [MSDN documentation for Series.DataLabels](https://learn.microsoft.com/en-us/office/vba/api/excel.series.datalabels).
        */
        wxExcelDataLabels DataLabels();


        /**
        Deletes the object.

        [MSDN documentation for Series.Delete](https://learn.microsoft.com/en-us/office/vba/api/excel.series.delete).
        */
        bool Delete();

        //@{
        /**
        Applies error bars to the series. Returns true on success.

        The error amounts are needed for some types only, e.g., not for xlErrorBarTypeStError.
        For xlErrorBarTypeCustom, Excel requires both @a amount and @a minusValues,
        which can then also be a Range, an array of numbers, or a string with an array constant
        or a reference, such as "={1,2,3}" or "=Sheet1!$D$2:$D$5". Use the overload taking a wxVariant
        for an array or a string, and the one taking a Range for ranges. A Range passed to the
        overload taking a wxVariant is converted to its values, so the error bars do not
        change with the cells then.

        [MSDN documentation for Series.ErrorBar](https://learn.microsoft.com/en-us/office/vba/api/excel.series.errorbar).
        */
        bool ErrorBar(XlErrorBarDirection direction, XlErrorBarInclude include, XlErrorBarType type,
                      double* amount = NULL, double* minusValues = NULL);
        bool ErrorBar(XlErrorBarDirection direction, XlErrorBarInclude include, XlErrorBarType type,
                      wxExcelRange amount, wxExcelRange minusValues);
        bool ErrorBar(XlErrorBarDirection direction, XlErrorBarInclude include, XlErrorBarType type,
                      const wxVariant& amount, const wxVariant& minusValues);
        //@}

        /**
        Pastes a picture from the Clipboard as the marker on the selected series.

        [MSDN documentation for Series.Paste](https://learn.microsoft.com/en-us/office/vba/api/excel.series.paste).
        */
        bool Paste();

        //@{
        /**
        Returns an object that represents a single point (a Point object) or a collection of all the points (a Points collection) in the series. Read-only

        [MSDN documentation for Series.Points](https://learn.microsoft.com/en-us/office/vba/api/excel.series.points).
        */
        wxExcelPoints Points();
        wxExcelPoint Points(long index);
        //@}

        /**
        Selects the object.

        [MSDN documentation for Series.Select](https://learn.microsoft.com/en-us/office/vba/api/excel.series.select).
        */
        bool Select();

        /**
        Returns an object that represents a single trendline (a Trendline object) or a collection of all the trendlines (a Trendlines collection) for the series.

        [MSDN documentation for Series.Trendlines](https://learn.microsoft.com/en-us/office/vba/api/excel.series.trendlines).
        */
        wxExcelTrendlines Trendlines();

        // ***** PROPERTIES *****

        /**
        True if a picture is applied to the end of the point or all points in the series.

        [MSDN documentation for Series.ApplyPictToEnd](https://learn.microsoft.com/en-us/office/vba/api/excel.series.applypicttoend).
        */
        bool GetApplyPictToEnd();

        /**
        True if a picture is applied to the end of the point or all points in the series.

        [MSDN documentation for Series.ApplyPictToEnd](https://learn.microsoft.com/en-us/office/vba/api/excel.series.applypicttoend).
        */
        void SetApplyPictToEnd(bool applyPictToEnd);

        /**
        True if a picture is applied to the front of the point or all points in the series.

        [MSDN documentation for Series.ApplyPictToFront](https://learn.microsoft.com/en-us/office/vba/api/excel.series.applypicttofront).
        */
        bool GetApplyPictToFront();

        /**
        True if a picture is applied to the front of the point or all points in the series.

        [MSDN documentation for Series.ApplyPictToFront](https://learn.microsoft.com/en-us/office/vba/api/excel.series.applypicttofront).
        */
        void SetApplyPictToFront(bool applyPictToFront);

        /**
        True if a picture is applied to the sides of the point or all points in the series.

        [MSDN documentation for Series.ApplyPictToSides](https://learn.microsoft.com/en-us/office/vba/api/excel.series.applypicttosides).
        */
        bool GetApplyPictToSides();

        /**
        True if a picture is applied to the sides of the point or all points in the series.

        [MSDN documentation for Series.ApplyPictToSides](https://learn.microsoft.com/en-us/office/vba/api/excel.series.applypicttosides).
        */
        void SetApplyPictToSides(bool applyPictToSides);

        /**
        Returns an XlAxisGroup value that represents the type of axis group.

        [MSDN documentation for Series.AxisGroup](https://learn.microsoft.com/en-us/office/vba/api/excel.series.axisgroup).
        */
        XlAxisGroup GetAxisGroup();

        /**
        Sets an XlAxisGroup value that represents the type of axis group.

        [MSDN documentation for Series.AxisGroup](https://learn.microsoft.com/en-us/office/vba/api/excel.series.axisgroup).
        */
        void SetAxisGroup(XlAxisGroup group);

        /**
        Returns the shape used with the 3-D bar or column chart. Read/write XlBarShape.

        [MSDN documentation for Series.BarShape](https://learn.microsoft.com/en-us/office/vba/api/excel.series.barshape).
        */
        XlBarShape GetBarShape();

        /**
        Sets the shape used with the 3-D bar or column chart. Read/write XlBarShape.

        [MSDN documentation for Series.BarShape](https://learn.microsoft.com/en-us/office/vba/api/excel.series.barshape).
        */
        void SetBarShape(XlBarShape barShape);

        /**
        Returns a string that refers to the worksheet cells containing the x-value, y-value and size data for the bubble chart. When you return the cell reference, it will return a string describing the cells in A1-style notation. To set the size data for the bubble chart, you must use R1C1-style notation. Applies only to bubble charts. Read/write Variant.

        [MSDN documentation for Series.BubbleSizes](https://learn.microsoft.com/en-us/office/vba/api/excel.series.bubblesizes).
        */
        wxString GetBubbleSizes();

        /**
        Sets a string that refers to the worksheet cells containing the x-value, y-value and size data for the bubble chart. When you return the cell reference, it will return a string describing the cells in A1-style notation. To set the size data for the bubble chart, you must use R1C1-style notation. Applies only to bubble charts. Read/write Variant.

        [MSDN documentation for Series.BubbleSizes](https://learn.microsoft.com/en-us/office/vba/api/excel.series.bubblesizes).
        */
        void SetBubbleSizes(const wxString& bubbleSizes);

        /**
        Returns the chart type. Read/write XlChartType.

        [MSDN documentation for Series.ChartType](https://learn.microsoft.com/en-us/office/vba/api/excel.series.charttype).
        */
        XlChartType GetChartType();

        /**
        Sets the chart type. Read/write XlChartType.

        [MSDN documentation for Series.ChartType](https://learn.microsoft.com/en-us/office/vba/api/excel.series.charttype).
        */
        void SetChartType(XlChartType chartType);

        /**
        Returns an ErrorBars object that represents the error bars for the series.

        [MSDN documentation for Series.ErrorBars](https://learn.microsoft.com/en-us/office/vba/api/excel.series.errorbars).
        */
        wxExcelErrorBars GetErrorBars();

        /**
        Returns the explosion value for a pie-chart or doughnut-chart slice. Returns 0 (zero) if there's no explosion (the tip of the slice is in the center of the pie).

        [MSDN documentation for Series.Explosion](https://learn.microsoft.com/en-us/office/vba/api/excel.series.explosion).
        */
        long GetExplosion();

        /**
        Sets the explosion value for a pie-chart or doughnut-chart slice. Returns 0 (zero) if there's no explosion (the tip of the slice is in the center of the pie).

        [MSDN documentation for Series.Explosion](https://learn.microsoft.com/en-us/office/vba/api/excel.series.explosion).
        */
        void SetExplosion(long explosion);

        /**
        Returns the ChartFormat object. Since Excel 2007.

        [MSDN documentation for Series.Format](https://learn.microsoft.com/en-us/office/vba/api/excel.series.format).
        */
        wxExcelChartFormat GetFormat();

        /**
        Returns a String value that represents the object's formula in A1-style notation and in the language of the macro.

        [MSDN documentation for Series.Formula](https://learn.microsoft.com/en-us/office/vba/api/excel.series.formula).
        */
        wxString GetFormula();

        /**
        Sets a String value that represents the object's formula in A1-style notation and in the language of the macro.

        [MSDN documentation for Series.Formula](https://learn.microsoft.com/en-us/office/vba/api/excel.series.formula).
        */
        void SetFormula(const wxString& formula);

        /**
        Returns the formula for the object, using A1-style references in the language of the user.

        [MSDN documentation for Series.FormulaLocal](https://learn.microsoft.com/en-us/office/vba/api/excel.series.formulalocal).
        */
        wxString GetFormulaLocal();

        /**
        Sets the formula for the object, using A1-style references in the language of the user.

        [MSDN documentation for Series.FormulaLocal](https://learn.microsoft.com/en-us/office/vba/api/excel.series.formulalocal).
        */
        void SetFormulaLocal(const wxString& formulaLocal);

        /**
        Returns the formula for the object, using R1C1-style notation in the language of the macro.

        [MSDN documentation for Series.FormulaR1C1](https://learn.microsoft.com/en-us/office/vba/api/excel.series.formular1c1).
        */
        wxString GetFormulaR1C1();

        /**
        Sets the formula for the object, using R1C1-style notation in the language of the macro.

        [MSDN documentation for Series.FormulaR1C1](https://learn.microsoft.com/en-us/office/vba/api/excel.series.formular1c1).
        */
        void SetFormulaR1C1(const wxString& formulaR1C1);

        /**
        Returns the formula for the object, using R1C1-style notation in the language of the user.

        [MSDN documentation for Series.FormulaR1C1Local](https://learn.microsoft.com/en-us/office/vba/api/excel.series.formular1c1local).
        */
        wxString GetFormulaR1C1Local();

        /**
        Sets the formula for the object, using R1C1-style notation in the language of the user.

        [MSDN documentation for Series.FormulaR1C1Local](https://learn.microsoft.com/en-us/office/vba/api/excel.series.formular1c1local).
        */
        void SetFormulaR1C1Local(const wxString& formulaR1C1Local);

        /**
        Specifies the geography projection type for the specified series within the chart group.

        [Excel VBA documentation for Series.GeoProjectionType](https://learn.microsoft.com/en-us/office/vba/api/excel.series.geoprojectiontype)
        */
        XlGeoProjectionType GetGeoProjectionType();

        /**
        Specifies the geography projection type for the specified series within the chart group.

        [Excel VBA documentation for Series.GeoProjectionType](https://learn.microsoft.com/en-us/office/vba/api/excel.series.geoprojectiontype)
        */
        void SetGeoProjectionType(XlGeoProjectionType type);

        /**
        True if the series has a three-dimensional appearance.

        [MSDN documentation for Series.Has3DEffect](https://learn.microsoft.com/en-us/office/vba/api/excel.series.has3deffect).
        */
        bool GetHas3DEffect();

        /**
        True if the series has a three-dimensional appearance.

        [MSDN documentation for Series.Has3DEffect](https://learn.microsoft.com/en-us/office/vba/api/excel.series.has3deffect).
        */
        void SetHas3DEffect(bool has3DEffect);

        /**
        True if the series has data labels.

        [MSDN documentation for Series.HasDataLabels](https://learn.microsoft.com/en-us/office/vba/api/excel.series.hasdatalabels).
        */
        bool GetHasDataLabels();

        /**
        True if the series has data labels.

        [MSDN documentation for Series.HasDataLabels](https://learn.microsoft.com/en-us/office/vba/api/excel.series.hasdatalabels).
        */
        void SetHasDataLabels(bool hasDataLabels);

        /**
        True if the series has error bars. This property isn’t available for 3-D charts.

        [MSDN documentation for Series.HasErrorBars](https://learn.microsoft.com/en-us/office/vba/api/excel.series.haserrorbars).
        */
        bool GetHasErrorBars();

        /**
        True if the series has error bars. This property isn’t available for 3-D charts.

        [MSDN documentation for Series.HasErrorBars](https://learn.microsoft.com/en-us/office/vba/api/excel.series.haserrorbars).
        */
        void SetHasErrorBars(bool hasErrorBars);

        /**
        True if the series has leader lines.

        [MSDN documentation for Series.HasLeaderLines](https://learn.microsoft.com/en-us/office/vba/api/excel.series.hasleaderlines).
        */
        bool GetHasLeaderLines();

        /**
        True if the series has leader lines.

        [MSDN documentation for Series.HasLeaderLines](https://learn.microsoft.com/en-us/office/vba/api/excel.series.hasleaderlines).
        */
        void SetHasLeaderLines(bool hasLeaderLines);

        /**
        True if Microsoft Excel inverts the pattern in the item when it corresponds to a negative number.

        [MSDN documentation for Series.InvertIfNegative](https://learn.microsoft.com/en-us/office/vba/api/excel.series.invertifnegative).
        */
        bool GetInvertIfNegative();

        /**
        True if Microsoft Excel inverts the pattern in the item when it corresponds to a negative number.

        [MSDN documentation for Series.InvertIfNegative](https://learn.microsoft.com/en-us/office/vba/api/excel.series.invertifnegative).
        */
        void SetInvertIfNegative(bool invertIfNegative);
        
        /**
        Whether to filter the series out from the chart.

        [Excel VBA documentation for Series.IsFiltered](https://learn.microsoft.com/en-us/office/vba/api/excel.series.isfiltered)
        */
        bool GetIsFiltered();

        /**
        Whether to filter the series out from the chart.

        [Excel VBA documentation for Series.IsFiltered](https://learn.microsoft.com/en-us/office/vba/api/excel.series.isfiltered)
        */
        void SetIsFiltered(bool isFiltered);

        /**
        Returns a LeaderLines object that represents the leader lines for the series.

        [MSDN documentation for Series.LeaderLines](https://learn.microsoft.com/en-us/office/vba/api/excel.series.leaderlines).
        */
        wxExcelLeaderLines GetLeaderLines();

        /**
        Sets the marker background color as an RGB value or returns the corresponding color index value. Applies only to line, scatter, and radar charts.

        [MSDN documentation for Series.MarkerBackgroundColor](https://learn.microsoft.com/en-us/office/vba/api/excel.series.markerbackgroundcolor).
        */
        wxColour GetMarkerBackgroundColor();

        /**
        Sets the marker background color as an RGB value or returns the corresponding color index value. Applies only to line, scatter, and radar charts.

        [MSDN documentation for Series.MarkerBackgroundColor](https://learn.microsoft.com/en-us/office/vba/api/excel.series.markerbackgroundcolor).
        */
        void SetMarkerBackgroundColor(const wxColour& markerBackgroundColor);

        /**
        Returns the marker background color as an index into the current color palette, or as one of the following XlColorIndex constants: xlColorIndexAutomatic or xlColorIndexNone. Applies only to line, scatter, and radar charts.

        [MSDN documentation for Series.MarkerBackgroundColorIndex](https://learn.microsoft.com/en-us/office/vba/api/excel.series.markerbackgroundcolorindex).
        */
        long GetMarkerBackgroundColorIndex();

        /**
        Sets the marker background color as an index into the current color palette, or as one of the following XlColorIndex constants: xlColorIndexAutomatic or xlColorIndexNone. Applies only to line, scatter, and radar charts.

        [MSDN documentation for Series.MarkerBackgroundColorIndex](https://learn.microsoft.com/en-us/office/vba/api/excel.series.markerbackgroundcolorindex).
        */
        void SetMarkerBackgroundColorIndex(long markerBackgroundColorIndex);

        /**
        Sets the marker background color as an RGB value or returns the corresponding color index value. Applies only to line, scatter, and radar charts.

        [MSDN documentation for Series.MarkerForegroundColor](https://learn.microsoft.com/en-us/office/vba/api/excel.series.markerforegroundcolor).
        */
        wxColour GetMarkerForegroundColor();

        /**
        Sets the marker background color as an RGB value or returns the corresponding color index value. Applies only to line, scatter, and radar charts.

        [MSDN documentation for Series.MarkerForegroundColor](https://learn.microsoft.com/en-us/office/vba/api/excel.series.markerforegroundcolor).
        */
        void SetMarkerForegroundColor(const wxColour& markerForegroundColor);

        /**
        Returns the marker foreground color as an index into the current color palette, or as one of the following XlColorIndex constants: xlColorIndexAutomatic or xlColorIndexNone. Applies only to line, scatter, and radar charts.

        [MSDN documentation for Series.MarkerForegroundColorIndex](https://learn.microsoft.com/en-us/office/vba/api/excel.series.markerforegroundcolorindex).
        */
        long GetMarkerForegroundColorIndex();

        /**
        Sets the marker foreground color as an index into the current color palette, or as one of the following XlColorIndex constants: xlColorIndexAutomatic or xlColorIndexNone. Applies only to line, scatter, and radar charts.

        [MSDN documentation for Series.MarkerForegroundColorIndex](https://learn.microsoft.com/en-us/office/vba/api/excel.series.markerforegroundcolorindex).
        */
        void SetMarkerForegroundColorIndex(long markerForegroundColorIndex);

        /**
        Returns the data-marker size, in points. Can be a value from 2 through 72.

        [MSDN documentation for Series.MarkerSize](https://learn.microsoft.com/en-us/office/vba/api/excel.series.markersize).
        */
        long GetMarkerSize();

        /**
        Sets the data-marker size, in points. Can be a value from 2 through 72.

        [MSDN documentation for Series.MarkerSize](https://learn.microsoft.com/en-us/office/vba/api/excel.series.markersize).
        */
        void SetMarkerSize(long markerSize);

        /**
        Returns the marker style for a point or series in a line chart, scatter chart, or radar chart. Read/write XlMarkerStyle.

        [MSDN documentation for Series.MarkerStyle](https://learn.microsoft.com/en-us/office/vba/api/excel.series.markerstyle).
        */
        XlMarkerStyle GetMarkerStyle();

        /**
        Sets the marker style for a point or series in a line chart, scatter chart, or radar chart. Read/write XlMarkerStyle.

        [MSDN documentation for Series.MarkerStyle](https://learn.microsoft.com/en-us/office/vba/api/excel.series.markerstyle).
        */
        void SetMarkerStyle(XlMarkerStyle markerStyle);

        /**
        Returns a String value representing the name of the object.

        [MSDN documentation for Series.Name](https://learn.microsoft.com/en-us/office/vba/api/excel.series.name).
        */
        wxString GetName();

        /**
        Sets a String value representing the name of the object.

        [MSDN documentation for Series.Name](https://learn.microsoft.com/en-us/office/vba/api/excel.series.name).
        */
        void SetName(const wxString& name);

        /**
        Specifies the parent data label option (banner, overlapping, or none) for the specified series within the chart group. Since Excel 2016.

        [MSDN documentation for Series.ParentDataLabelOption](https://learn.microsoft.com/en-us/office/vba/api/excel.series.parentdatalabeloption).
        */
        XlParentDataLabelOptions GetParentDataLabelOption();

        /**
        Specifies the parent data label option (banner, overlapping, or none) for the specified series within the chart group. Since Excel 2016.

        [MSDN documentation for Series.ParentDataLabelOption](https://learn.microsoft.com/en-us/office/vba/api/excel.series.parentdatalabeloption).
        */
        void SetParentDataLabelOption(XlParentDataLabelOptions option);

        /**
        Returns a XlChartPictureType value that represents the way pictures are displayed on a column or bar picture chart.

        [MSDN documentation for Series.PictureType](https://learn.microsoft.com/en-us/office/vba/api/excel.series.picturetype).
        */
        XlChartPictureType GetPictureType();

        /**
        Sets a XlChartPictureType value that represents the way pictures are displayed on a column or bar picture chart.

        [MSDN documentation for Series.PictureType](https://learn.microsoft.com/en-us/office/vba/api/excel.series.picturetype).
        */
        void SetPictureType(XlChartPictureType pictureType);

        /**
        Returns the unit for each picture on the chart if the PictureType property is set to xlStackScale (if not, this property is ignored).

        [MSDN documentation for Series.PictureUnit]().
        */
        long GetPictureUnit();

        /**
        Sets the unit for each picture on the chart if the PictureType property is set to xlStackScale (if not, this property is ignored).

        [MSDN documentation for Series.PictureUnit]().
        */
        void SetPictureUnit(long pictureUnit);

        /**
        Returns the unit for each picture on the chart if the PictureType property is set to xlStackScale (if not, this property is ignored). Since Excel 2007.

        [MSDN documentation for Series.PictureUnit2](https://learn.microsoft.com/en-us/office/vba/api/excel.series.pictureunit2).
        */
        double GetPictureUnit2();

        /**
        Sets the unit for each picture on the chart if the PictureType property is set to xlStackScale (if not, this property is ignored). Since Excel 2007.

        [MSDN documentation for Series.PictureUnit2](https://learn.microsoft.com/en-us/office/vba/api/excel.series.pictureunit2).
        */
        void SetPictureUnit2(double pictureUnit2);

        /**
        Returns the plot order for the selected series within the chart group.

        [MSDN documentation for Series.PlotOrder](https://learn.microsoft.com/en-us/office/vba/api/excel.series.plotorder).
        */
        long GetPlotOrder();

        /**
        Sets the plot order for the selected series within the chart group.

        [MSDN documentation for Series.PlotOrder](https://learn.microsoft.com/en-us/office/vba/api/excel.series.plotorder).
        */
        void SetPlotOrder(long plotOrder);

        /**
        True if the series uses an inclusive median quartile calculation method. Since Excel 2016.

        [MSDN documentation for Series.QuartileCalculationInclusiveMedian](https://learn.microsoft.com/en-us/office/vba/api/excel.series.quartilecalculationinclusivemedian).
        */
        bool GetQuartileCalculationInclusiveMedian();

        /**
        True if the series uses an inclusive median quartile calculation method. Since Excel 2016.

        [MSDN documentation for Series.QuartileCalculationInclusiveMedian](https://learn.microsoft.com/en-us/office/vba/api/excel.series.quartilecalculationinclusivemedian).
        */
        void SetQuartileCalculationInclusiveMedian(bool value);

        /**
        Specifies the region labelling behavior for the specified series within the chart group.

        [Excel VBA documentation for Series.RegionLabelOptions](https://learn.microsoft.com/en-us/office/vba/api/excel.series.regionlabeloptions)
        */
        XlRegionLabelOptions GetRegionLabelOptions();

        /**
        Specifies the region labelling behavior for the specified series within the chart group.

        [Excel VBA documentation for Series.RegionLabelOptions](https://learn.microsoft.com/en-us/office/vba/api/excel.series.regionlabeloptions)
        */
        void SetRegionLabelOptions(XlRegionLabelOptions options);

        /**
        Returns a Boolean value that determines if the object has a shadow.

        [MSDN documentation for Series.Shadow](https://learn.microsoft.com/en-us/office/vba/api/excel.series.shadow).
        */
        bool GetShadow();

        /**
        Sets a Boolean value that determines if the object has a shadow.

        [MSDN documentation for Series.Shadow](https://learn.microsoft.com/en-us/office/vba/api/excel.series.shadow).
        */
        void SetShadow(bool shadow);

        /**
        True if curve smoothing is turned on for the line chart or scatter chart. Applies only to line and scatter charts.

        [MSDN documentation for Series.Smooth](https://learn.microsoft.com/en-us/office/vba/api/excel.series.smooth).
        */
        bool GetSmooth();

        /**
        True if curve smoothing is turned on for the line chart or scatter chart. Applies only to line and scatter charts.

        [MSDN documentation for Series.Smooth](https://learn.microsoft.com/en-us/office/vba/api/excel.series.smooth).
        */
        void SetSmooth(bool smooth);

        /**
        Returns a Long value that represents the series type.

        [MSDN documentation for Series.Type](https://learn.microsoft.com/en-us/office/vba/api/excel.series.type).
        */
        long GetType();

        /**
        Sets a Long value that represents the series type.

        [MSDN documentation for Series.Type](https://learn.microsoft.com/en-us/office/vba/api/excel.series.type).
        */
        void SetType(long type);

        /**
        After calling GetValues(), first check if the result is true.
        If it is then if rangeValues.IsOk_() is true, the values are stored in that Range, else values were copied into variantValues.

        [MSDN documentation for Series.Values](https://learn.microsoft.com/en-us/office/vba/api/excel.series.values).
        */
        bool GetValues(wxExcelRange& rangeValues, wxVariant& variantValues);

        //@{
        /**
        Sets a Variant value that represents a collection of all the values in the series.

        [MSDN documentation for Series.Values](https://learn.microsoft.com/en-us/office/vba/api/excel.series.values).
        */
        void SetValues(wxExcelRange values);
        void SetValues(const wxVariant& values);
        //@}

        /**
        After calling GetXValues(), first check if the result is true.
        If it is then if rangeValues.IsOk_() is true, the values are stored in that Range, else values were copied into variantValues.

        [MSDN documentation for Series.XValues](https://learn.microsoft.com/en-us/office/vba/api/excel.series.xvalues).
        */
        bool GetXValues(wxExcelRange& rangeValues, wxVariant& variantValues);

        //@{
        /**
        Sets an array of x values for a chart series. The XValues property can be set to a range on a worksheet or to an array of values, but it cannot be a combination of both. Read/write Variant.

        [MSDN documentation for Series.XValues](https://learn.microsoft.com/en-us/office/vba/api/excel.series.xvalues).
        */
        void SetXValues(wxExcelRange values);
        void SetXValues(const wxVariant& values);
        //@}

        /**
        Returns "Series".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("Series"); }
    };


} // namespace wxAutoExcel

#endif // #if WXAUTOEXCEL_USE_CHARTS

#endif // _WXAUTOEXCEL_SERIES_H
