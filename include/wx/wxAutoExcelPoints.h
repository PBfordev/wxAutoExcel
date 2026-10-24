/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// License:     MIT license
/////////////////////////////////////////////////////////////////////////////


#ifndef _WXAUTOEXCEL_POINTS_H
#define _WXAUTOEXCEL_POINTS_H

#include "wx/wxAutoExcel_defs.h"

#if WXAUTOEXCEL_USE_CHARTS

#include "wx/wxAutoExcel_object.h"
#include "wx/wxAutoExcel_enums.h"

class WXDLLIMPEXP_FWD_CORE wxColour;

namespace wxAutoExcel {

    /**
    @brief Represents Microsoft Excel Point object.
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelPoint : public wxExcelObject
    {
    public:
        // ***** METHODS *****

        /**
        Applies data labels to a point.

        [MSDN documentation for Point.ApplyDataLabels](https://learn.microsoft.com/en-us/office/vba/api/excel.point.applydatalabels).
        */
        void ApplyDataLabels(XlDataLabelsType* type = NULL, wxXlTribool legendKey = wxDefaultXlTribool,
                             wxXlTribool autoText = wxDefaultXlTribool, wxXlTribool hasLeaderLines = wxDefaultXlTribool,
                             wxXlTribool showSeriesName = wxDefaultXlTribool, wxXlTribool showCategoryName = wxDefaultXlTribool,
                             wxXlTribool showValue = wxDefaultXlTribool, wxXlTribool showPercentage = wxDefaultXlTribool,
                             wxXlTribool showBubbleSize = wxDefaultXlTribool, const wxString& separator = wxEmptyString);

        /**
        Clears the formatting of the object.

        [MSDN documentation for Point.ClearFormats](https://learn.microsoft.com/en-us/office/vba/api/excel.point.clearformats).
        */
        bool ClearFormats();

        /**
        If the point has a picture fill, then this method copies the picture to the Clipboard.

        [MSDN documentation for Point.Copy](https://learn.microsoft.com/en-us/office/vba/api/excel.point.copy).
        */
        bool Copy();

        /**
        Deletes the series the point belongs to.

        [MSDN documentation for Point.Delete](https://learn.microsoft.com/en-us/office/vba/api/excel.point.delete).
        */
        bool Delete();

        /**
        Pastes a picture from the Clipboard as the marker on the selected point.

        [MSDN documentation for Point.Paste](https://learn.microsoft.com/en-us/office/vba/api/excel.point.paste).
        */
        bool Paste();

        /**
        Selects the object.

        [MSDN documentation for Point.Select](https://learn.microsoft.com/en-us/office/vba/api/excel.point.select).
        */
        bool Select();

        // ***** PROPERTIES *****

        /**
        True if a picture is applied to the end of the point or all points in the series.

        [MSDN documentation for Point.ApplyPictToEnd](https://learn.microsoft.com/en-us/office/vba/api/excel.point.applypicttoend).
        */
        bool GetApplyPictToEnd();

        /**
        True if a picture is applied to the end of the point or all points in the series.

        [MSDN documentation for Point.ApplyPictToEnd](https://learn.microsoft.com/en-us/office/vba/api/excel.point.applypicttoend).
        */
        void SetApplyPictToEnd(bool applyPictToEnd);

        /**
        True if a picture is applied to the front of the point or all points in the series.

        [MSDN documentation for Point.ApplyPictToFront](https://learn.microsoft.com/en-us/office/vba/api/excel.point.applypicttofront).
        */
        bool GetApplyPictToFront();

        /**
        True if a picture is applied to the front of the point or all points in the series.

        [MSDN documentation for Point.ApplyPictToFront](https://learn.microsoft.com/en-us/office/vba/api/excel.point.applypicttofront).
        */
        void SetApplyPictToFront(bool applyPictToFront);

        /**
        True if a picture is applied to the sides of the point or all points in the series.

        [MSDN documentation for Point.ApplyPictToSides](https://learn.microsoft.com/en-us/office/vba/api/excel.point.applypicttosides).
        */
        bool GetApplyPictToSides();

        /**
        True if a picture is applied to the sides of the point or all points in the series.

        [MSDN documentation for Point.ApplyPictToSides](https://learn.microsoft.com/en-us/office/vba/api/excel.point.applypicttosides).
        */
        void SetApplyPictToSides(bool applyPictToSides);

        /**
        Returns a DataLabel object that represents the data label associated with the point.

        [MSDN documentation for Point.DataLabel](https://learn.microsoft.com/en-us/office/vba/api/excel.point.datalabel).
        */
        wxExcelDataLabel GetDataLabel();

        /**
        Returns the explosion value for a pie-chart or doughnut-chart slice. Returns 0 (zero) if there's no explosion (the tip of the slice is in the center of the pie).

        [MSDN documentation for Point.Explosion](https://learn.microsoft.com/en-us/office/vba/api/excel.point.explosion).
        */
        long GetExplosion();

        /**
        Sets the explosion value for a pie-chart or doughnut-chart slice. Returns 0 (zero) if there's no explosion (the tip of the slice is in the center of the pie).

        [MSDN documentation for Point.Explosion](https://learn.microsoft.com/en-us/office/vba/api/excel.point.explosion).
        */
        void SetExplosion(long explosion);

        /**
        Returns the ChartFormat object. Since Excel 2007.

        [MSDN documentation for Point.Format](https://learn.microsoft.com/en-us/office/vba/api/excel.point.format).
        */
        wxExcelChartFormat GetFormat();

        /**
        True if a point has a three-dimensional appearance.  Since Excel 2007.

        [MSDN documentation for Point.Has3DEffect](https://learn.microsoft.com/en-us/office/vba/api/excel.point.has3deffect).
        */
        bool GetHas3DEffect();

        /**
        True if a point has a three-dimensional appearance.  Since Excel 2007.

        [MSDN documentation for Point.Has3DEffect](https://learn.microsoft.com/en-us/office/vba/api/excel.point.has3deffect).
        */
        void SetHas3DEffect(bool has3DEffect);

        /**
        True if the point has a data label.

        [MSDN documentation for Point.HasDataLabel](https://learn.microsoft.com/en-us/office/vba/api/excel.point.hasdatalabel).
        */
        bool GetHasDataLabel();

        /**
        True if the point has a data label.

        [MSDN documentation for Point.HasDataLabel](https://learn.microsoft.com/en-us/office/vba/api/excel.point.hasdatalabel).
        */
        void SetHasDataLabel(bool hasDataLabel);

        /**
        True if Microsoft Excel inverts the pattern in the item when it corresponds to a negative number.

        [MSDN documentation for Point.InvertIfNegative](https://learn.microsoft.com/en-us/office/vba/api/excel.point.invertifnegative).
        */
        bool GetInvertIfNegative();

        /**
        True if Microsoft Excel inverts the pattern in the item when it corresponds to a negative number.

        [MSDN documentation for Point.InvertIfNegative](https://learn.microsoft.com/en-us/office/vba/api/excel.point.invertifnegative).
        */
        void SetInvertIfNegative(bool invertIfNegative);

        /**True if the point represents a total. Since Excel 2016.

        [MSDN documentation for Point.IsTotal](https://learn.microsoft.com/en-us/office/vba/api/excel.point.istotal).
        */
        bool GetIsTotal();

        /**True if the point represents a total. Since Excel 2016.

        [MSDN documentation for Point.IsTotal](https://learn.microsoft.com/en-us/office/vba/api/excel.point.istotal).
        */
        void SetIsTotal(bool isTotal);

        /**
        Sets the marker background color as an RGB value or returns the corresponding color index value. Applies only to line, scatter, and radar charts.

        [MSDN documentation for Point.MarkerBackgroundColor](https://learn.microsoft.com/en-us/office/vba/api/excel.point.markerbackgroundcolor).
        */
        wxColour GetMarkerBackgroundColor();

        /**
        Sets the marker background color as an RGB value or returns the corresponding color index value. Applies only to line, scatter, and radar charts.

        [MSDN documentation for Point.MarkerBackgroundColor](https://learn.microsoft.com/en-us/office/vba/api/excel.point.markerbackgroundcolor).
        */
        void SetMarkerBackgroundColor(const wxColour& markerBackgroundColor);

        /**
        Returns the marker background color as an index into the current color palette, or as one of the following XlColorIndex constants: xlColorIndexAutomatic or xlColorIndexNone. Applies only to line, scatter, and radar charts.

        [MSDN documentation for Point.MarkerBackgroundColorIndex](https://learn.microsoft.com/en-us/office/vba/api/excel.point.markerbackgroundcolorindex).
        */
        long GetMarkerBackgroundColorIndex();

        /**
        Sets the marker background color as an index into the current color palette, or as one of the following XlColorIndex constants: xlColorIndexAutomatic or xlColorIndexNone. Applies only to line, scatter, and radar charts.

        [MSDN documentation for Point.MarkerBackgroundColorIndex](https://learn.microsoft.com/en-us/office/vba/api/excel.point.markerbackgroundcolorindex).
        */
        void SetMarkerBackgroundColorIndex(long markerBackgroundColorIndex);

        /**
        Sets the marker background color as an RGB value or returns the corresponding color index value. Applies only to line, scatter, and radar charts.

        [MSDN documentation for Point.MarkerForegroundColor](https://learn.microsoft.com/en-us/office/vba/api/excel.point.markerforegroundcolor).
        */
        wxColour GetMarkerForegroundColor();

        /**
        Sets the marker background color as an RGB value or returns the corresponding color index value. Applies only to line, scatter, and radar charts.

        [MSDN documentation for Point.MarkerForegroundColor](https://learn.microsoft.com/en-us/office/vba/api/excel.point.markerforegroundcolor).
        */
        void SetMarkerForegroundColor(const wxColour& markerForegroundColor);

        /**
        Returns the marker foreground color as an index into the current color palette, or as one of the following XlColorIndex constants: xlColorIndexAutomatic or xlColorIndexNone. Applies only to line, scatter, and radar charts.

        [MSDN documentation for Point.MarkerForegroundColorIndex](https://learn.microsoft.com/en-us/office/vba/api/excel.point.markerforegroundcolorindex).
        */
        long GetMarkerForegroundColorIndex();

        /**
        Sets the marker foreground color as an index into the current color palette, or as one of the following XlColorIndex constants: xlColorIndexAutomatic or xlColorIndexNone. Applies only to line, scatter, and radar charts.

        [MSDN documentation for Point.MarkerForegroundColorIndex](https://learn.microsoft.com/en-us/office/vba/api/excel.point.markerforegroundcolorindex).
        */
        void SetMarkerForegroundColorIndex(long markerForegroundColorIndex);

        /**
        Returns the data-marker size, in points. Can be a value from 2 through 72.

        [MSDN documentation for Point.MarkerSize](https://learn.microsoft.com/en-us/office/vba/api/excel.point.markersize).
        */
        long GetMarkerSize();

        /**
        Sets the data-marker size, in points. Can be a value from 2 through 72.

        [MSDN documentation for Point.MarkerSize](https://learn.microsoft.com/en-us/office/vba/api/excel.point.markersize).
        */
        void SetMarkerSize(long markerSize);

        /**
        Returns the marker style for a point or series in a line chart, scatter chart, or radar chart. Read/write XlMarkerStyle.

        [MSDN documentation for Point.MarkerStyle](https://learn.microsoft.com/en-us/office/vba/api/excel.point.markerstyle).
        */
        XlMarkerStyle GetMarkerStyle();

        /**
        Sets the marker style for a point or series in a line chart, scatter chart, or radar chart. Read/write XlMarkerStyle.

        [MSDN documentation for Point.MarkerStyle](https://learn.microsoft.com/en-us/office/vba/api/excel.point.markerstyle).
        */
        void SetMarkerStyle(XlMarkerStyle markerStyle);

        /**
        Returns a XlChartPictureType value that represents the way pictures are displayed on a column or bar picture chart.

        [MSDN documentation for Point.PictureType](https://learn.microsoft.com/en-us/office/vba/api/excel.point.picturetype).
        */
        XlChartPictureType GetPictureType();

        /**
        Sets a XlChartPictureType value that represents the way pictures are displayed on a column or bar picture chart.

        [MSDN documentation for Point.PictureType](https://learn.microsoft.com/en-us/office/vba/api/excel.point.picturetype).
        */
        void SetPictureType(XlChartPictureType pictureType);

        /**
        Returns the unit for each picture on the chart if the PictureType property is set to xlStackScale (if not, this property is ignored).

        [MSDN documentation for Point.PictureUnit]().
        */
        long GetPictureUnit();

        /**
        Sets the unit for each picture on the chart if the PictureType property is set to xlStackScale (if not, this property is ignored).

        [MSDN documentation for Point.PictureUnit]().
        */
        void SetPictureUnit(long pictureUnit);

        /**
        Read/write Since Excel 2007.

        [MSDN documentation for Point.PictureUnit2](https://learn.microsoft.com/en-us/office/vba/api/excel.point.pictureunit2).
        */
        double GetPictureUnit2();

        /**
        Read/write Since Excel 2007.

        [MSDN documentation for Point.PictureUnit2](https://learn.microsoft.com/en-us/office/vba/api/excel.point.pictureunit2).
        */
        void SetPictureUnit2(double pictureUnit2);

        /**
        True if the point is in the secondary section of either a pie of pie chart or a bar of pie chart. Applies only to points on pie of pie charts or bar of pie charts.

        [MSDN documentation for Point.SecondaryPlot](https://learn.microsoft.com/en-us/office/vba/api/excel.point.secondaryplot).
        */
        bool GetSecondaryPlot();

        /**
        True if the point is in the secondary section of either a pie of pie chart or a bar of pie chart. Applies only to points on pie of pie charts or bar of pie charts.

        [MSDN documentation for Point.SecondaryPlot](https://learn.microsoft.com/en-us/office/vba/api/excel.point.secondaryplot).
        */
        void SetSecondaryPlot(bool secondaryPlot);

        /**
        Returns a Boolean value that determines if the object has a shadow.

        [MSDN documentation for Point.Shadow](https://learn.microsoft.com/en-us/office/vba/api/excel.point.shadow).
        */
        bool GetShadow();

        /**
        Sets a Boolean value that determines if the object has a shadow.

        [MSDN documentation for Point.Shadow](https://learn.microsoft.com/en-us/office/vba/api/excel.point.shadow).
        */
        void SetShadow(bool shadow);
        /**
        Returns "Point".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("Point"); }
    };

    /**
    @brief Represents Microsoft Excel Points collection.
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelPoints : public wxExcelObject
    {
    public:

        //@{
        /**
        Returns a single object from a collection.

        [MSDN documentation for Points.Item](https://learn.microsoft.com/en-us/office/vba/api/excel.points.item).
        */
        wxExcelPoint Item(long index);
        wxExcelPoint operator[](long index);
        //@}

        // ***** PROPERTIES *****

        /**
        Returns a Long value that represents the number of objects in the collection.

        [MSDN documentation for Points.Count](https://learn.microsoft.com/en-us/office/vba/api/excel.points.count).
        */
        long GetCount();

        /**
        The iterator over the items, see wxExcelCollectionIterator.
        */
        typedef wxExcelCollectionIterator<wxExcelPoints, wxExcelPoint> iterator;

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
        Returns "Points".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("Points"); }
    };


} // namespace wxAutoExcel

#endif // #if WXAUTOEXCEL_USE_CHARTS

#endif // _WXAUTOEXCEL_POINTS_H
