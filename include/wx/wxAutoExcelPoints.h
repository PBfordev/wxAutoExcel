/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Modified by:
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
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
    Represents Microsoft Excel Point object.
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelPoint : public wxExcelObject
    {
    public:
        // ***** METHODS *****

        /**
        Applies data labels to a point.

        [MSDN documentation for Point.ApplyDataLabels](http://msdn.microsoft.com/en-us/library/bb178774).
        */
        void ApplyDataLabels(XlDataLabelsType* type = NULL, wxXlTribool legendKey = wxDefaultXlTribool,
                             wxXlTribool autoText = wxDefaultXlTribool, wxXlTribool hasLeaderLines = wxDefaultXlTribool,
                             wxXlTribool showSeriesName = wxDefaultXlTribool, wxXlTribool showCategoryName = wxDefaultXlTribool,
                             wxXlTribool showValue = wxDefaultXlTribool, wxXlTribool showPercentage = wxDefaultXlTribool,
                             wxXlTribool showBubbleSize = wxDefaultXlTribool, const wxString& separator = wxEmptyString);

        /**
        Clears the formatting of the object.

        [MSDN documentation for Point.ClearFormats](http://msdn.microsoft.com/en-us/library/bb178776).
        */
        bool ClearFormats();

        /**
        If the point has a picture fill, then this method copies the picture to the Clipboard.

        [MSDN documentation for Point.Copy](http://msdn.microsoft.com/en-us/library/bb178781).
        */
        bool Copy();

        /**
        Deletes the series the point belongs to.

        [MSDN documentation for Point.Delete](http://msdn.microsoft.com/en-us/library/bb178784).
        */
        bool Delete();

        /**
        Pastes a picture from the Clipboard as the marker on the selected point.

        [MSDN documentation for Point.Paste](http://msdn.microsoft.com/en-us/library/bb178789).
        */
        bool Paste();

        /**
        Selects the object.

        [MSDN documentation for Point.Select](http://msdn.microsoft.com/en-us/library/bb238179).
        */
        bool Select();

        // ***** PROPERTIES *****

        /**
        True if a picture is applied to the end of the point or all points in the series.

        [MSDN documentation for Point.ApplyPictToEnd](http://msdn.microsoft.com/en-us/library/bb237454).
        */
        bool GetApplyPictToEnd();

        /**
        True if a picture is applied to the end of the point or all points in the series.

        [MSDN documentation for Point.ApplyPictToEnd](http://msdn.microsoft.com/en-us/library/bb237454).
        */
        void SetApplyPictToEnd(bool applyPictToEnd);

        /**
        True if a picture is applied to the front of the point or all points in the series.

        [MSDN documentation for Point.ApplyPictToFront](http://msdn.microsoft.com/en-us/library/bb237458).
        */
        bool GetApplyPictToFront();

        /**
        True if a picture is applied to the front of the point or all points in the series.

        [MSDN documentation for Point.ApplyPictToFront](http://msdn.microsoft.com/en-us/library/bb237458).
        */
        void SetApplyPictToFront(bool applyPictToFront);

        /**
        True if a picture is applied to the sides of the point or all points in the series.

        [MSDN documentation for Point.ApplyPictToSides](http://msdn.microsoft.com/en-us/library/bb237461).
        */
        bool GetApplyPictToSides();

        /**
        True if a picture is applied to the sides of the point or all points in the series.

        [MSDN documentation for Point.ApplyPictToSides](http://msdn.microsoft.com/en-us/library/bb237461).
        */
        void SetApplyPictToSides(bool applyPictToSides);

        /**
        Returns a DataLabel object that represents the data label associated with the point.

        [MSDN documentation for Point.DataLabel](http://msdn.microsoft.com/en-us/library/bb237464).
        */
        wxExcelDataLabel GetDataLabel();

        /**
        Returns the explosion value for a pie-chart or doughnut-chart slice. Returns 0 (zero) if there's no explosion (the tip of the slice is in the center of the pie).

        [MSDN documentation for Point.Explosion](http://msdn.microsoft.com/en-us/library/bb237467).
        */
        long GetExplosion();

        /**
        Sets the explosion value for a pie-chart or doughnut-chart slice. Returns 0 (zero) if there's no explosion (the tip of the slice is in the center of the pie).

        [MSDN documentation for Point.Explosion](http://msdn.microsoft.com/en-us/library/bb237467).
        */
        void SetExplosion(long explosion);

        /**
        Returns the ChartFormat object. Since Excel 2007.

        [MSDN documentation for Point.Format](http://msdn.microsoft.com/en-us/library/bb242537).
        */
        wxExcelChartFormat GetFormat();

        /**
        True if a point has a three-dimensional appearance.  Since Excel 2007.

        [MSDN documentation for Point.Has3DEffect](http://msdn.microsoft.com/en-us/library/bb237471).
        */
        bool GetHas3DEffect();

        /**
        True if a point has a three-dimensional appearance.  Since Excel 2007.

        [MSDN documentation for Point.Has3DEffect](http://msdn.microsoft.com/en-us/library/bb237471).
        */
        void SetHas3DEffect(bool has3DEffect);

        /**
        True if the point has a data label.

        [MSDN documentation for Point.HasDataLabel](http://msdn.microsoft.com/en-us/library/bb208606).
        */
        bool GetHasDataLabel();

        /**
        True if the point has a data label.

        [MSDN documentation for Point.HasDataLabel](http://msdn.microsoft.com/en-us/library/bb208606).
        */
        void SetHasDataLabel(bool hasDataLabel);

        /**
        True if Microsoft Excel inverts the pattern in the item when it corresponds to a negative number.

        [MSDN documentation for Point.InvertIfNegative](http://msdn.microsoft.com/en-us/library/bb237474).
        */
        bool GetInvertIfNegative();

        /**
        True if Microsoft Excel inverts the pattern in the item when it corresponds to a negative number.

        [MSDN documentation for Point.InvertIfNegative](http://msdn.microsoft.com/en-us/library/bb237474).
        */
        void SetInvertIfNegative(bool invertIfNegative);

        /**True if the point represents a total. Since Excel 2016.

        [MSDN documentation for Point.IsTotal](https://msdn.microsoft.com/VBA/Excel-VBA/articles/point-istotal-property-excel).
        */
        bool GetIsTotal();

        /**True if the point represents a total. Since Excel 2016.

        [MSDN documentation for Point.IsTotal](https://msdn.microsoft.com/VBA/Excel-VBA/articles/point-istotal-property-excel).
        */
        void SetIsTotal(bool isTotal);

        /**
        Sets the marker background color as an RGB value or returns the corresponding color index value. Applies only to line, scatter, and radar charts.

        [MSDN documentation for Point.MarkerBackgroundColor](http://msdn.microsoft.com/en-us/library/bb237478).
        */
        wxColour GetMarkerBackgroundColor();

        /**
        Sets the marker background color as an RGB value or returns the corresponding color index value. Applies only to line, scatter, and radar charts.

        [MSDN documentation for Point.MarkerBackgroundColor](http://msdn.microsoft.com/en-us/library/bb237478).
        */
        void SetMarkerBackgroundColor(const wxColour& markerBackgroundColor);

        /**
        Returns the marker background color as an index into the current color palette, or as one of the following XlColorIndex constants: xlColorIndexAutomatic or xlColorIndexNone. Applies only to line, scatter, and radar charts.

        [MSDN documentation for Point.MarkerBackgroundColorIndex](http://msdn.microsoft.com/en-us/library/bb237481).
        */
        long GetMarkerBackgroundColorIndex();

        /**
        Sets the marker background color as an index into the current color palette, or as one of the following XlColorIndex constants: xlColorIndexAutomatic or xlColorIndexNone. Applies only to line, scatter, and radar charts.

        [MSDN documentation for Point.MarkerBackgroundColorIndex](http://msdn.microsoft.com/en-us/library/bb237481).
        */
        void SetMarkerBackgroundColorIndex(long markerBackgroundColorIndex);

        /**
        Sets the marker background color as an RGB value or returns the corresponding color index value. Applies only to line, scatter, and radar charts.

        [MSDN documentation for Point.MarkerForegroundColor](http://msdn.microsoft.com/en-us/library/bb237483).
        */
        wxColour GetMarkerForegroundColor();

        /**
        Sets the marker background color as an RGB value or returns the corresponding color index value. Applies only to line, scatter, and radar charts.

        [MSDN documentation for Point.MarkerForegroundColor](http://msdn.microsoft.com/en-us/library/bb237483).
        */
        void SetMarkerForegroundColor(const wxColour& markerForegroundColor);

        /**
        Returns the marker foreground color as an index into the current color palette, or as one of the following XlColorIndex constants: xlColorIndexAutomatic or xlColorIndexNone. Applies only to line, scatter, and radar charts.

        [MSDN documentation for Point.MarkerForegroundColorIndex](http://msdn.microsoft.com/en-us/library/bb237487).
        */
        long GetMarkerForegroundColorIndex();

        /**
        Sets the marker foreground color as an index into the current color palette, or as one of the following XlColorIndex constants: xlColorIndexAutomatic or xlColorIndexNone. Applies only to line, scatter, and radar charts.

        [MSDN documentation for Point.MarkerForegroundColorIndex](http://msdn.microsoft.com/en-us/library/bb237487).
        */
        void SetMarkerForegroundColorIndex(long markerForegroundColorIndex);

        /**
        Returns the data-marker size, in points. Can be a value from 2 through 72.

        [MSDN documentation for Point.MarkerSize](http://msdn.microsoft.com/en-us/library/bb237489).
        */
        long GetMarkerSize();

        /**
        Sets the data-marker size, in points. Can be a value from 2 through 72.

        [MSDN documentation for Point.MarkerSize](http://msdn.microsoft.com/en-us/library/bb237489).
        */
        void SetMarkerSize(long markerSize);

        /**
        Returns the marker style for a point or series in a line chart, scatter chart, or radar chart. Read/write XlMarkerStyle.

        [MSDN documentation for Point.MarkerStyle](http://msdn.microsoft.com/en-us/library/bb213463).
        */
        XlMarkerStyle GetMarkerStyle();

        /**
        Sets the marker style for a point or series in a line chart, scatter chart, or radar chart. Read/write XlMarkerStyle.

        [MSDN documentation for Point.MarkerStyle](http://msdn.microsoft.com/en-us/library/bb213463).
        */
        void SetMarkerStyle(XlMarkerStyle markerStyle);

        /**
        Returns a XlChartPictureType value that represents the way pictures are displayed on a column or bar picture chart.

        [MSDN documentation for Point.PictureType](http://msdn.microsoft.com/en-us/library/bb213465).
        */
        XlChartPictureType GetPictureType();

        /**
        Sets a XlChartPictureType value that represents the way pictures are displayed on a column or bar picture chart.

        [MSDN documentation for Point.PictureType](http://msdn.microsoft.com/en-us/library/bb213465).
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

        [MSDN documentation for Point.PictureUnit2](http://msdn.microsoft.com/en-us/library/bb240597).
        */
        double GetPictureUnit2();

        /**
        Read/write Since Excel 2007.

        [MSDN documentation for Point.PictureUnit2](http://msdn.microsoft.com/en-us/library/bb240597).
        */
        void SetPictureUnit2(double pictureUnit2);

        /**
        True if the point is in the secondary section of either a pie of pie chart or a bar of pie chart. Applies only to points on pie of pie charts or bar of pie charts.

        [MSDN documentation for Point.SecondaryPlot](http://msdn.microsoft.com/en-us/library/bb221624).
        */
        bool GetSecondaryPlot();

        /**
        True if the point is in the secondary section of either a pie of pie chart or a bar of pie chart. Applies only to points on pie of pie charts or bar of pie charts.

        [MSDN documentation for Point.SecondaryPlot](http://msdn.microsoft.com/en-us/library/bb221624).
        */
        void SetSecondaryPlot(bool secondaryPlot);

        /**
        Returns a Boolean value that determines if the object has a shadow.

        [MSDN documentation for Point.Shadow](http://msdn.microsoft.com/en-us/library/bb238563).
        */
        bool GetShadow();

        /**
        Sets a Boolean value that determines if the object has a shadow.

        [MSDN documentation for Point.Shadow](http://msdn.microsoft.com/en-us/library/bb238563).
        */
        void SetShadow(bool shadow);
        /**
        Returns "Point".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("Point"); }
    };

    /**
    Represents Microsoft Excel Points collection.
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelPoints : public wxExcelObject
    {
    public:

        //@{
        /**
        Returns a single object from a collection.

        [MSDN documentation for Points.Item](http://msdn.microsoft.com/en-us/library/bb178790).
        */
        wxExcelPoint Item(long index);
        wxExcelPoint operator[](long index);
        //@}

        // ***** PROPERTIES *****

        /**
        Returns a Long value that represents the number of objects in the collection.

        [MSDN documentation for Points.Count](http://msdn.microsoft.com/en-us/library/bb213469).
        */
        long GetCount();

        /**
        Returns "Points".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("Points"); }
    };


} // namespace wxAutoExcel

#endif // #if WXAUTOEXCEL_USE_CHARTS

#endif // _WXAUTOEXCEL_POINTS_H
