/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// License:     MIT license
/////////////////////////////////////////////////////////////////////////////


#ifndef _WXAUTOEXCEL_LEGENDKEY_H
#define _WXAUTOEXCEL_LEGENDKEY_H

#include "wx/wxAutoExcel_defs.h"

#if WXAUTOEXCEL_USE_CHARTS

#include "wx/wxAutoExcel_object.h"
#include "wx/wxAutoExcel_enums.h"

namespace wxAutoExcel {

    /**
    @brief Represents Microsoft Excel LegendKey object.
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelLegendKey : public wxExcelObject
    {
    public:
        // ***** METHODS *****

        /**
        Clears the formatting of the object.

        [MSDN documentation for LegendKey.ClearFormats](https://learn.microsoft.com/en-us/office/vba/api/excel.legendkey.clearformats).
        */
        bool ClearFormats();

        /**
        Deletes the object.

        [MSDN documentation for LegendKey.Delete](https://learn.microsoft.com/en-us/office/vba/api/excel.legendkey.delete).
        */
        bool Delete();

        // ***** PROPERTIES *****

        /**
        Returns the ChartFormat object. Since Excel 2007.

        [MSDN documentation for LegendKey.Format](https://learn.microsoft.com/en-us/office/vba/api/excel.legendkey.format).
        */
        wxExcelChartFormat GetFormat();

        /**
        Returns a Double value that represents the height, in points, of the object.

        [MSDN documentation for LegendKey.Height](https://learn.microsoft.com/en-us/office/vba/api/excel.legendkey.height).
        */
        double GetHeight();

        /**
        True if Microsoft Excel inverts the pattern in the item when it corresponds to a negative number.

        [MSDN documentation for LegendKey.InvertIfNegative](https://learn.microsoft.com/en-us/office/vba/api/excel.legendkey.invertifnegative).
        */
        bool GetInvertIfNegative();

        /**
        True if Microsoft Excel inverts the pattern in the item when it corresponds to a negative number.

        [MSDN documentation for LegendKey.InvertIfNegative](https://learn.microsoft.com/en-us/office/vba/api/excel.legendkey.invertifnegative).
        */
        void SetInvertIfNegative(bool invertIfNegative);

        /**
        Returns a Double value that represents the distance, in points, from the left edge of the object to the left edge of the chart area.

        [MSDN documentation for LegendKey.Left](https://learn.microsoft.com/en-us/office/vba/api/excel.legendkey.left).
        */
        double GetLeft();

        /**
        Sets the marker background color as an RGB value or returns the corresponding color index value. Applies only to line, scatter, and radar charts.

        [MSDN documentation for LegendKey.MarkerBackgroundColor](https://learn.microsoft.com/en-us/office/vba/api/excel.legendkey.markerbackgroundcolor).
        */
        wxColour GetMarkerBackgroundColor();

        /**
        Sets the marker background color as an RGB value or returns the corresponding color index value. Applies only to line, scatter, and radar charts.

        [MSDN documentation for LegendKey.MarkerBackgroundColor](https://learn.microsoft.com/en-us/office/vba/api/excel.legendkey.markerbackgroundcolor).
        */
        void SetMarkerBackgroundColor(const wxColour& markerBackgroundColor);

        /**
        Returns the marker background color as an index into the current color palette, or as one of the following XlColorIndex constants: xlColorIndexAutomatic or xlColorIndexNone. Applies only to line, scatter, and radar charts.

        [MSDN documentation for LegendKey.MarkerBackgroundColorIndex](https://learn.microsoft.com/en-us/office/vba/api/excel.legendkey.markerbackgroundcolorindex).
        */
        long GetMarkerBackgroundColorIndex();

        /**
        Sets the marker background color as an index into the current color palette, or as one of the following XlColorIndex constants: xlColorIndexAutomatic or xlColorIndexNone. Applies only to line, scatter, and radar charts.

        [MSDN documentation for LegendKey.MarkerBackgroundColorIndex](https://learn.microsoft.com/en-us/office/vba/api/excel.legendkey.markerbackgroundcolorindex).
        */
        void SetMarkerBackgroundColorIndex(long markerBackgroundColorIndex);

        /**
        Sets the marker background color as an RGB value or returns the corresponding color index value. Applies only to line, scatter, and radar charts.

        [MSDN documentation for LegendKey.MarkerForegroundColor](https://learn.microsoft.com/en-us/office/vba/api/excel.legendkey.markerforegroundcolor).
        */
        wxColour GetMarkerForegroundColor();

        /**
        Sets the marker background color as an RGB value or returns the corresponding color index value. Applies only to line, scatter, and radar charts.

        [MSDN documentation for LegendKey.MarkerForegroundColor](https://learn.microsoft.com/en-us/office/vba/api/excel.legendkey.markerforegroundcolor).
        */
        void SetMarkerForegroundColor(const wxColour& markerForegroundColor);

        /**
        Returns the marker foreground color as an index into the current color palette, or as one of the following XlColorIndex constants: xlColorIndexAutomatic or xlColorIndexNone. Applies only to line, scatter, and radar charts.

        [MSDN documentation for LegendKey.MarkerForegroundColorIndex](https://learn.microsoft.com/en-us/office/vba/api/excel.legendkey.markerforegroundcolorindex).
        */
        long GetMarkerForegroundColorIndex();

        /**
        Sets the marker foreground color as an index into the current color palette, or as one of the following XlColorIndex constants: xlColorIndexAutomatic or xlColorIndexNone. Applies only to line, scatter, and radar charts.

        [MSDN documentation for LegendKey.MarkerForegroundColorIndex](https://learn.microsoft.com/en-us/office/vba/api/excel.legendkey.markerforegroundcolorindex).
        */
        void SetMarkerForegroundColorIndex(long markerForegroundColorIndex);

        /**
        Returns the data-marker size, in points. Can be a value from 2 through 72.

        [MSDN documentation for LegendKey.MarkerSize](https://learn.microsoft.com/en-us/office/vba/api/excel.legendkey.markersize).
        */
        long GetMarkerSize();

        /**
        Sets the data-marker size, in points. Can be a value from 2 through 72.

        [MSDN documentation for LegendKey.MarkerSize](https://learn.microsoft.com/en-us/office/vba/api/excel.legendkey.markersize).
        */
        void SetMarkerSize(long markerSize);

        /**
        Returns the marker style for a point or series in a line chart, scatter chart, or radar chart. Read/write XlMarkerStyle.

        [MSDN documentation for LegendKey.MarkerStyle](https://learn.microsoft.com/en-us/office/vba/api/excel.legendkey.markerstyle).
        */
        XlMarkerStyle GetMarkerStyle();

        /**
        Sets the marker style for a point or series in a line chart, scatter chart, or radar chart. Read/write XlMarkerStyle.

        [MSDN documentation for LegendKey.MarkerStyle](https://learn.microsoft.com/en-us/office/vba/api/excel.legendkey.markerstyle).
        */
        void SetMarkerStyle(XlMarkerStyle markerStyle);

        /**
        Returns a XlChartPictureType value that represents the way pictures are displayed on a legend key.

        [MSDN documentation for LegendKey.PictureType](https://learn.microsoft.com/en-us/office/vba/api/excel.legendkey.picturetype).
        */
        XlChartPictureType GetPictureType();

        /**
        Sets a XlChartPictureType value that represents the way pictures are displayed on a legend key.

        [MSDN documentation for LegendKey.PictureType](https://learn.microsoft.com/en-us/office/vba/api/excel.legendkey.picturetype).
        */
        void SetPictureType(XlChartPictureType pictureType);

        /**
        Returns the unit for each picture on the chart if the PictureType property is set to xlStackScale (if not, this property is ignored).

        [MSDN documentation for LegendKey.PictureUnit](https://learn.microsoft.com/en-us/previous-versions/office/developer/office-2007/bb148579%28v=office.12%29).
        */
        long GetPictureUnit();

        /**
        Sets the unit for each picture on the chart if the PictureType property is set to xlStackScale (if not, this property is ignored).

        [MSDN documentation for LegendKey.PictureUnit](https://learn.microsoft.com/en-us/previous-versions/office/developer/office-2007/bb148579%28v=office.12%29).
        */
        void SetPictureUnit(long pictureUnit);

        /**
        Returns the unit for each picture on the chart if the PictureType property is set to xlStackScale (if not, this property is ignored). Since Excel 2007.

        [MSDN documentation for LegendKey.PictureUnit2](https://learn.microsoft.com/en-us/office/vba/api/excel.legendkey.pictureunit2).
        */
        double GetPictureUnit2();

        /**
        Sets the unit for each picture on the chart if the PictureType property is set to xlStackScale (if not, this property is ignored). Since Excel 2007.

        [MSDN documentation for LegendKey.PictureUnit2](https://learn.microsoft.com/en-us/office/vba/api/excel.legendkey.pictureunit2).
        */
        void SetPictureUnit2(double pictureUnit2);

        /**
        Returns a Boolean value that determines if the object has a shadow.

        [MSDN documentation for LegendKey.Shadow](https://learn.microsoft.com/en-us/office/vba/api/excel.legendkey.shadow).
        */
        bool GetShadow();

        /**
        Sets a Boolean value that determines if the object has a shadow.

        [MSDN documentation for LegendKey.Shadow](https://learn.microsoft.com/en-us/office/vba/api/excel.legendkey.shadow).
        */
        void SetShadow(bool shadow);

        /**
        True if curve smoothing is turned on for the legend key.

        [MSDN documentation for LegendKey.Smooth](https://learn.microsoft.com/en-us/office/vba/api/excel.legendkey.smooth).
        */
        bool GetSmooth();

        /**
        True if curve smoothing is turned on for the legend key.

        [MSDN documentation for LegendKey.Smooth](https://learn.microsoft.com/en-us/office/vba/api/excel.legendkey.smooth).
        */
        void SetSmooth(bool smooth);

        /**
        Returns a Double value that represents the distance, in points, from the top edge of the object to the top of row 1 (on a worksheet) or the top of the chart area (on a chart).

        [MSDN documentation for LegendKey.Top](https://learn.microsoft.com/en-us/office/vba/api/excel.legendkey.top).
        */
        double GetTop();

        /**
        Returns a Double value that represents the width, in points, of the object.

        [MSDN documentation for LegendKey.Width](https://learn.microsoft.com/en-us/office/vba/api/excel.legendkey.width).
        */
        double GetWidth();

        /**
        Returns "LegendKey".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("LegendKey"); }
    };


} // namespace wxAutoExcel

#endif // #if WXAUTOEXCEL_USE_CHARTS

#endif // _WXAUTOEXCEL_LEGENDKEY_H
