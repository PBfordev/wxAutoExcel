/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Modified by:
// RCS-ID:      $Id: $
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////


#ifndef _WXAUTOEXCEL_SERIES_H
#define _WXAUTOEXCEL_SERIES_H

#include "wx/wxAutoExcel_defs.h"

#if WXAUTOEXCEL_USE_CHARTS

#include "wx/wxAutoExcelObject.h"
#include "wx/wxAutoExcel_enums.h"

namespace wxAutoExcel {

    /**
    Represents Microsoft Excel Series object.
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelSeries : public wxExcelObject
    {
    public:
        // ***** METHODS *****

        /**
        Applies data labels to a series.

        [MSDN documentation for Series.ApplyDataLabels](http://msdn.microsoft.com/en-us/library/bb178905).
        */
        void ApplyDataLabels(XlDataLabelsType* type = NULL, wxXlTribool legendKey = wxDefaultXlTribool, 
                             wxXlTribool autoText = wxDefaultXlTribool, wxXlTribool hasLeaderLines = wxDefaultXlTribool, 
                             wxXlTribool showSeriesName = wxDefaultXlTribool, wxXlTribool showCategoryName = wxDefaultXlTribool, 
                             wxXlTribool showValue = wxDefaultXlTribool, wxXlTribool showPercentage = wxDefaultXlTribool, 
                             wxXlTribool showBubbleSize = wxDefaultXlTribool, const wxString& separator = wxEmptyString);

        /**
        Clears the formatting of the object.

        [MSDN documentation for Series.ClearFormats](http://msdn.microsoft.com/en-us/library/bb178909).
        */
        bool ClearFormats();

        /**
        If the series has a picture fill, then this method copies the picture to the Clipboard.

        [MSDN documentation for Series.Copy](http://msdn.microsoft.com/en-us/library/bb178925).
        */
        bool Copy();

        /**
        Returns an object that represents either a single data label (a DataLabel object) or a collection of all the data labels for the series (a DataLabels collection).

        [MSDN documentation for Series.DataLabels](http://msdn.microsoft.com/en-us/library/bb223310).
        */
        wxExcelDataLabels DataLabels();        
        

        /**
        Deletes the object.

        [MSDN documentation for Series.Delete](http://msdn.microsoft.com/en-us/library/bb178930).
        */
        bool Delete();

        /**
        Applies error bars to the series. Variant.

        [MSDN documentation for Series.ErrorBar](http://msdn.microsoft.com/en-us/library/bb209818).
        */
        bool ErrorBar(XlErrorBarDirection direction, XlErrorBarInclude include, XlErrorBarType type,
                      double* amount = NULL, double* minusValues = NULL);

        /**
        Pastes a picture from the Clipboard as the marker on the selected series.

        [MSDN documentation for Series.Paste](http://msdn.microsoft.com/en-us/library/bb178937).
        */
        bool Paste();

        //@{
        /**
        Returns an object that represents a single point (a Point object) or a collection of all the points (a Points collection) in the series. Read-only

        [MSDN documentation for Series.Points](http://msdn.microsoft.com/en-us/library/bb242027).
        */
        wxExcelPoints Points();
        wxExcelPoint Points(long index);
        //@}

        /**
        Selects the object.

        [MSDN documentation for Series.Select](http://msdn.microsoft.com/en-us/library/bb238260).
        */
        bool Select();
        
        /**
        Returns an object that represents a single trendline (a Trendline object) or a collection of all the trendlines (a Trendlines collection) for the series.

        [MSDN documentation for Series.Trendlines](http://msdn.microsoft.com/en-us/library/bb178165).
        */
        wxExcelTrendlines Trendlines();
        
        // ***** PROPERTIES *****

        /**
        True if a picture is applied to the end of the point or all points in the series.

        [MSDN documentation for Series.ApplyPictToEnd](http://msdn.microsoft.com/en-us/library/bb237539).
        */
        bool GetApplyPictToEnd();

        /**
        True if a picture is applied to the end of the point or all points in the series.

        [MSDN documentation for Series.ApplyPictToEnd](http://msdn.microsoft.com/en-us/library/bb237539).
        */
        void SetApplyPictToEnd(bool applyPictToEnd);

        /**
        True if a picture is applied to the front of the point or all points in the series.

        [MSDN documentation for Series.ApplyPictToFront](http://msdn.microsoft.com/en-us/library/bb237545).
        */
        bool GetApplyPictToFront();

        /**
        True if a picture is applied to the front of the point or all points in the series.

        [MSDN documentation for Series.ApplyPictToFront](http://msdn.microsoft.com/en-us/library/bb237545).
        */
        void SetApplyPictToFront(bool applyPictToFront);

        /**
        True if a picture is applied to the sides of the point or all points in the series.

        [MSDN documentation for Series.ApplyPictToSides](http://msdn.microsoft.com/en-us/library/bb237549).
        */
        bool GetApplyPictToSides();

        /**
        True if a picture is applied to the sides of the point or all points in the series.

        [MSDN documentation for Series.ApplyPictToSides](http://msdn.microsoft.com/en-us/library/bb237549).
        */
        void SetApplyPictToSides(bool applyPictToSides);

        /**
        Returns an XlAxisGroup value that represents the type of axis group.

        [MSDN documentation for Series.AxisGroup](http://msdn.microsoft.com/en-us/library/dd787725).
        */
        XlAxisGroup GetAxisGroup();

        /**
        Returns the shape used with the 3-D bar or column chart. Read/write XlBarShape.

        [MSDN documentation for Series.BarShape](http://msdn.microsoft.com/en-us/library/bb237550).
        */
        XlBarShape GetBarShape();

        /**
        Sets the shape used with the 3-D bar or column chart. Read/write XlBarShape.

        [MSDN documentation for Series.BarShape](http://msdn.microsoft.com/en-us/library/bb237550).
        */
        void SetBarShape(XlBarShape barShape);

        /**
        Returns a string that refers to the worksheet cells containing the x-value, y-value and size data for the bubble chart. When you return the cell reference, it will return a string describing the cells in A1-style notation. To set the size data for the bubble chart, you must use R1C1-style notation. Applies only to bubble charts. Read/write Variant.

        [MSDN documentation for Series.BubbleSizes](http://msdn.microsoft.com/en-us/library/bb220894).
        */
        wxString GetBubbleSizes();

        /**
        Sets a string that refers to the worksheet cells containing the x-value, y-value and size data for the bubble chart. When you return the cell reference, it will return a string describing the cells in A1-style notation. To set the size data for the bubble chart, you must use R1C1-style notation. Applies only to bubble charts. Read/write Variant.

        [MSDN documentation for Series.BubbleSizes](http://msdn.microsoft.com/en-us/library/bb220894).
        */
        void SetBubbleSizes(const wxString& bubbleSizes);

        /**
        Returns the chart type. Read/write XlChartType.

        [MSDN documentation for Series.ChartType](http://msdn.microsoft.com/en-us/library/bb237556).
        */
        XlChartType GetChartType();

        /**
        Sets the chart type. Read/write XlChartType.

        [MSDN documentation for Series.ChartType](http://msdn.microsoft.com/en-us/library/bb237556).
        */
        void SetChartType(XlChartType chartType);

        /**
        Returns an ErrorBars object that represents the error bars for the series.

        [MSDN documentation for Series.ErrorBars](http://msdn.microsoft.com/en-us/library/bb208471).
        */
        wxExcelErrorBars GetErrorBars();

        /**
        Returns the explosion value for a pie-chart or doughnut-chart slice. Returns 0 (zero) if there's no explosion (the tip of the slice is in the center of the pie).

        [MSDN documentation for Series.Explosion](http://msdn.microsoft.com/en-us/library/bb237571).
        */
        long GetExplosion();

        /**
        Sets the explosion value for a pie-chart or doughnut-chart slice. Returns 0 (zero) if there's no explosion (the tip of the slice is in the center of the pie).

        [MSDN documentation for Series.Explosion](http://msdn.microsoft.com/en-us/library/bb237571).
        */
        void SetExplosion(long explosion);

        /**
        Returns the ChartFormat object. Since Excel 2007.

        [MSDN documentation for Series.Format](http://msdn.microsoft.com/en-us/library/bb242538).
        */
        wxExcelChartFormat GetFormat();

        /**
        Returns a String value that represents the object's formula in A1-style notation and in the language of the macro.

        [MSDN documentation for Series.Formula](http://msdn.microsoft.com/en-us/library/bb237572).
        */
        wxString GetFormula();

        /**
        Sets a String value that represents the object's formula in A1-style notation and in the language of the macro.

        [MSDN documentation for Series.Formula](http://msdn.microsoft.com/en-us/library/bb237572).
        */
        void SetFormula(const wxString& formula);

        /**
        Returns the formula for the object, using A1-style references in the language of the user.

        [MSDN documentation for Series.FormulaLocal](http://msdn.microsoft.com/en-us/library/bb237576).
        */
        wxString GetFormulaLocal();

        /**
        Sets the formula for the object, using A1-style references in the language of the user.

        [MSDN documentation for Series.FormulaLocal](http://msdn.microsoft.com/en-us/library/bb237576).
        */
        void SetFormulaLocal(const wxString& formulaLocal);

        /**
        Returns the formula for the object, using R1C1-style notation in the language of the macro.

        [MSDN documentation for Series.FormulaR1C1](http://msdn.microsoft.com/en-us/library/bb237582).
        */
        wxString GetFormulaR1C1();

        /**
        Sets the formula for the object, using R1C1-style notation in the language of the macro.

        [MSDN documentation for Series.FormulaR1C1](http://msdn.microsoft.com/en-us/library/bb237582).
        */
        void SetFormulaR1C1(const wxString& formulaR1C1);

        /**
        Returns the formula for the object, using R1C1-style notation in the language of the user.

        [MSDN documentation for Series.FormulaR1C1Local](http://msdn.microsoft.com/en-us/library/bb237587).
        */
        wxString GetFormulaR1C1Local();

        /**
        Sets the formula for the object, using R1C1-style notation in the language of the user.

        [MSDN documentation for Series.FormulaR1C1Local](http://msdn.microsoft.com/en-us/library/bb237587).
        */
        void SetFormulaR1C1Local(const wxString& formulaR1C1Local);

        /**
        True if the series has a three-dimensional appearance. 

        [MSDN documentation for Series.Has3DEffect](http://msdn.microsoft.com/en-us/library/bb237590).
        */
        bool GetHas3DEffect();

        /**
        True if the series has a three-dimensional appearance. 

        [MSDN documentation for Series.Has3DEffect](http://msdn.microsoft.com/en-us/library/bb237590).
        */
        void SetHas3DEffect(bool has3DEffect);

        /**
        True if the series has data labels.

        [MSDN documentation for Series.HasDataLabels](http://msdn.microsoft.com/en-us/library/bb208608).
        */
        bool GetHasDataLabels();

        /**
        True if the series has data labels.

        [MSDN documentation for Series.HasDataLabels](http://msdn.microsoft.com/en-us/library/bb208608).
        */
        void SetHasDataLabels(bool hasDataLabels);

        /**
        True if the series has error bars. This property isn’t available for 3-D charts.

        [MSDN documentation for Series.HasErrorBars](http://msdn.microsoft.com/en-us/library/bb208622).
        */
        bool GetHasErrorBars();

        /**
        True if the series has error bars. This property isn’t available for 3-D charts.

        [MSDN documentation for Series.HasErrorBars](http://msdn.microsoft.com/en-us/library/bb208622).
        */
        void SetHasErrorBars(bool hasErrorBars);

        /**
        True if the series has leader lines.

        [MSDN documentation for Series.HasLeaderLines](http://msdn.microsoft.com/en-us/library/bb208633).
        */
        bool GetHasLeaderLines();

        /**
        True if the series has leader lines.

        [MSDN documentation for Series.HasLeaderLines](http://msdn.microsoft.com/en-us/library/bb208633).
        */
        void SetHasLeaderLines(bool hasLeaderLines);

        /**
        True if Microsoft Excel inverts the pattern in the item when it corresponds to a negative number.

        [MSDN documentation for Series.InvertIfNegative](http://msdn.microsoft.com/en-us/library/bb237591).
        */
        bool GetInvertIfNegative();

        /**
        True if Microsoft Excel inverts the pattern in the item when it corresponds to a negative number.

        [MSDN documentation for Series.InvertIfNegative](http://msdn.microsoft.com/en-us/library/bb237591).
        */
        void SetInvertIfNegative(bool invertIfNegative);

        /**
        Returns a LeaderLines object that represents the leader lines for the series.

        [MSDN documentation for Series.LeaderLines](http://msdn.microsoft.com/en-us/library/bb177844).
        */
        wxExcelLeaderLines GetLeaderLines();

        /**
        Sets the marker background color as an RGB value or returns the corresponding color index value. Applies only to line, scatter, and radar charts.

        [MSDN documentation for Series.MarkerBackgroundColor](http://msdn.microsoft.com/en-us/library/bb237597).
        */
        wxColour GetMarkerBackgroundColor();

        /**
        Sets the marker background color as an RGB value or returns the corresponding color index value. Applies only to line, scatter, and radar charts.

        [MSDN documentation for Series.MarkerBackgroundColor](http://msdn.microsoft.com/en-us/library/bb237597).
        */
        void SetMarkerBackgroundColor(const wxColour& markerBackgroundColor);

        /**
        Returns the marker background color as an index into the current color palette, or as one of the following XlColorIndex constants: xlColorIndexAutomatic or xlColorIndexNone. Applies only to line, scatter, and radar charts.

        [MSDN documentation for Series.MarkerBackgroundColorIndex](http://msdn.microsoft.com/en-us/library/bb237600).
        */
        long GetMarkerBackgroundColorIndex();

        /**
        Sets the marker background color as an index into the current color palette, or as one of the following XlColorIndex constants: xlColorIndexAutomatic or xlColorIndexNone. Applies only to line, scatter, and radar charts.

        [MSDN documentation for Series.MarkerBackgroundColorIndex](http://msdn.microsoft.com/en-us/library/bb237600).
        */
        void SetMarkerBackgroundColorIndex(long markerBackgroundColorIndex);

        /**
        Sets the marker background color as an RGB value or returns the corresponding color index value. Applies only to line, scatter, and radar charts.

        [MSDN documentation for Series.MarkerForegroundColor](http://msdn.microsoft.com/en-us/library/bb237602).
        */
        wxColour GetMarkerForegroundColor();

        /**
        Sets the marker background color as an RGB value or returns the corresponding color index value. Applies only to line, scatter, and radar charts.

        [MSDN documentation for Series.MarkerForegroundColor](http://msdn.microsoft.com/en-us/library/bb237602).
        */
        void SetMarkerForegroundColor(const wxColour& markerForegroundColor);

        /**
        Returns the marker foreground color as an index into the current color palette, or as one of the following XlColorIndex constants: xlColorIndexAutomatic or xlColorIndexNone. Applies only to line, scatter, and radar charts.

        [MSDN documentation for Series.MarkerForegroundColorIndex](http://msdn.microsoft.com/en-us/library/bb237607).
        */
        long GetMarkerForegroundColorIndex();

        /**
        Sets the marker foreground color as an index into the current color palette, or as one of the following XlColorIndex constants: xlColorIndexAutomatic or xlColorIndexNone. Applies only to line, scatter, and radar charts.

        [MSDN documentation for Series.MarkerForegroundColorIndex](http://msdn.microsoft.com/en-us/library/bb237607).
        */
        void SetMarkerForegroundColorIndex(long markerForegroundColorIndex);

        /**
        Returns the data-marker size, in points. Can be a value from 2 through 72.

        [MSDN documentation for Series.MarkerSize](http://msdn.microsoft.com/en-us/library/bb237613).
        */
        long GetMarkerSize();

        /**
        Sets the data-marker size, in points. Can be a value from 2 through 72.

        [MSDN documentation for Series.MarkerSize](http://msdn.microsoft.com/en-us/library/bb237613).
        */
        void SetMarkerSize(long markerSize);

        /**
        Returns the marker style for a point or series in a line chart, scatter chart, or radar chart. Read/write XlMarkerStyle.

        [MSDN documentation for Series.MarkerStyle](http://msdn.microsoft.com/en-us/library/bb237615).
        */
        XlMarkerStyle GetMarkerStyle();

        /**
        Sets the marker style for a point or series in a line chart, scatter chart, or radar chart. Read/write XlMarkerStyle.

        [MSDN documentation for Series.MarkerStyle](http://msdn.microsoft.com/en-us/library/bb237615).
        */
        void SetMarkerStyle(XlMarkerStyle markerStyle);

        /**
        Returns a String value representing the name of the object.

        [MSDN documentation for Series.Name](http://msdn.microsoft.com/en-us/library/bb237616).
        */
        wxString GetName();

        /**
        Sets a String value representing the name of the object.

        [MSDN documentation for Series.Name](http://msdn.microsoft.com/en-us/library/bb237616).
        */
        void SetName(const wxString& name);

        /**
        Returns a XlChartPictureType value that represents the way pictures are displayed on a column or bar picture chart.

        [MSDN documentation for Series.PictureType](http://msdn.microsoft.com/en-us/library/bb237621).
        */
        XlChartPictureType GetPictureType();

        /**
        Sets a XlChartPictureType value that represents the way pictures are displayed on a column or bar picture chart.

        [MSDN documentation for Series.PictureType](http://msdn.microsoft.com/en-us/library/bb237621).
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

        [MSDN documentation for Series.PictureUnit2](http://msdn.microsoft.com/en-us/library/bb240685).
        */
        double GetPictureUnit2();

        /**
        Sets the unit for each picture on the chart if the PictureType property is set to xlStackScale (if not, this property is ignored). Since Excel 2007.

        [MSDN documentation for Series.PictureUnit2](http://msdn.microsoft.com/en-us/library/bb240685).
        */
        void SetPictureUnit2(double pictureUnit2);

        /**
        Returns the plot order for the selected series within the chart group.

        [MSDN documentation for Series.PlotOrder](http://msdn.microsoft.com/en-us/library/bb221433).
        */
        long GetPlotOrder();

        /**
        Sets the plot order for the selected series within the chart group.

        [MSDN documentation for Series.PlotOrder](http://msdn.microsoft.com/en-us/library/bb221433).
        */
        void SetPlotOrder(long plotOrder);

        /**
        Returns a Boolean value that determines if the object has a shadow.

        [MSDN documentation for Series.Shadow](http://msdn.microsoft.com/en-us/library/bb238626).
        */
        bool GetShadow();

        /**
        Sets a Boolean value that determines if the object has a shadow.

        [MSDN documentation for Series.Shadow](http://msdn.microsoft.com/en-us/library/bb238626).
        */
        void SetShadow(bool shadow);

        /**
        True if curve smoothing is turned on for the line chart or scatter chart. Applies only to line and scatter charts.

        [MSDN documentation for Series.Smooth](http://msdn.microsoft.com/en-us/library/bb238629).
        */
        bool GetSmooth();

        /**
        True if curve smoothing is turned on for the line chart or scatter chart. Applies only to line and scatter charts.

        [MSDN documentation for Series.Smooth](http://msdn.microsoft.com/en-us/library/bb238629).
        */
        void SetSmooth(bool smooth);

        /**
        Returns a Long value that represents the series type.

        [MSDN documentation for Series.Type](http://msdn.microsoft.com/en-us/library/bb238633).
        */
        long GetType();

        /**
        Sets a Long value that represents the series type.

        [MSDN documentation for Series.Type](http://msdn.microsoft.com/en-us/library/bb238633).
        */
        void SetType(long type);

        /**
        After calling GetValues(), first check if the result is true. 
        If it is then if rangeValues.IsOk_() is true, the values are stored in that Range, else values were copied into variantValues.

        [MSDN documentation for Series.Values](http://msdn.microsoft.com/en-us/library/bb238636).
        */
        bool GetValues(wxExcelRange& rangeValues, wxVariant& variantValues);

        //@{
        /**
        Sets a Variant value that represents a collection of all the values in the series.

        [MSDN documentation for Series.Values](http://msdn.microsoft.com/en-us/library/bb238636).
        */
        void SetValues(wxExcelRange values);
        void SetValues(const wxVariant& values);
        //@}

        /**
        After calling GetXValues(), first check if the result is true. 
        If it is then if rangeValues.IsOk_() is true, the values are stored in that Range, else values were copied into variantValues.

        [MSDN documentation for Series.XValues](http://msdn.microsoft.com/en-us/library/bb209521).
        */
        bool GetXValues(wxExcelRange& rangeValues, wxVariant& variantValues);

        //@{
        /**
        Sets an array of x values for a chart series. The XValues property can be set to a range on a worksheet or to an array of values, but it cannot be a combination of both. Read/write Variant.

        [MSDN documentation for Series.XValues](http://msdn.microsoft.com/en-us/library/bb209521).
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
