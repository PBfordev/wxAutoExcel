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

        [MSDN documentation for LegendKey.ClearFormats](http://msdn.microsoft.com/en-us/library/bb211841).
        */
        bool ClearFormats();

        /**
        Deletes the object.

        [MSDN documentation for LegendKey.Delete](http://msdn.microsoft.com/en-us/library/bb211845).
        */
        bool Delete();

        // ***** PROPERTIES *****

        /**
        Returns the ChartFormat object. Since Excel 2007.

        [MSDN documentation for LegendKey.Format](http://msdn.microsoft.com/en-us/library/bb242535).
        */
        wxExcelChartFormat GetFormat();

        /**
        Returns a Double value that represents the height, in points, of the object.

        [MSDN documentation for LegendKey.Height](http://msdn.microsoft.com/en-us/library/bb148551).
        */
        double GetHeight();

        /**
        True if Microsoft Excel inverts the pattern in the item when it corresponds to a negative number.

        [MSDN documentation for LegendKey.InvertIfNegative](http://msdn.microsoft.com/en-us/library/bb148553).
        */
        bool GetInvertIfNegative();

        /**
        True if Microsoft Excel inverts the pattern in the item when it corresponds to a negative number.

        [MSDN documentation for LegendKey.InvertIfNegative](http://msdn.microsoft.com/en-us/library/bb148553).
        */
        void SetInvertIfNegative(bool invertIfNegative);

        /**
        Returns a Double value that represents the distance, in points, from the left edge of the object to the left edge of the chart area.

        [MSDN documentation for LegendKey.Left](http://msdn.microsoft.com/en-us/library/bb148554).
        */
        double GetLeft();

        /**
        Sets the marker background color as an RGB value or returns the corresponding color index value. Applies only to line, scatter, and radar charts.

        [MSDN documentation for LegendKey.MarkerBackgroundColor](http://msdn.microsoft.com/en-us/library/bb148558).
        */
        wxColour GetMarkerBackgroundColor();

        /**
        Sets the marker background color as an RGB value or returns the corresponding color index value. Applies only to line, scatter, and radar charts.

        [MSDN documentation for LegendKey.MarkerBackgroundColor](http://msdn.microsoft.com/en-us/library/bb148558).
        */
        void SetMarkerBackgroundColor(const wxColour& markerBackgroundColor);

        /**
        Returns the marker background color as an index into the current color palette, or as one of the following XlColorIndex constants: xlColorIndexAutomatic or xlColorIndexNone. Applies only to line, scatter, and radar charts.

        [MSDN documentation for LegendKey.MarkerBackgroundColorIndex](http://msdn.microsoft.com/en-us/library/bb148562).
        */
        long GetMarkerBackgroundColorIndex();

        /**
        Sets the marker background color as an index into the current color palette, or as one of the following XlColorIndex constants: xlColorIndexAutomatic or xlColorIndexNone. Applies only to line, scatter, and radar charts.

        [MSDN documentation for LegendKey.MarkerBackgroundColorIndex](http://msdn.microsoft.com/en-us/library/bb148562).
        */
        void SetMarkerBackgroundColorIndex(long markerBackgroundColorIndex);

        /**
        Sets the marker background color as an RGB value or returns the corresponding color index value. Applies only to line, scatter, and radar charts.

        [MSDN documentation for LegendKey.MarkerForegroundColor](http://msdn.microsoft.com/en-us/library/bb148564).
        */
        wxColour GetMarkerForegroundColor();

        /**
        Sets the marker background color as an RGB value or returns the corresponding color index value. Applies only to line, scatter, and radar charts.

        [MSDN documentation for LegendKey.MarkerForegroundColor](http://msdn.microsoft.com/en-us/library/bb148564).
        */
        void SetMarkerForegroundColor(const wxColour& markerForegroundColor);

        /**
        Returns the marker foreground color as an index into the current color palette, or as one of the following XlColorIndex constants: xlColorIndexAutomatic or xlColorIndexNone. Applies only to line, scatter, and radar charts.

        [MSDN documentation for LegendKey.MarkerForegroundColorIndex](http://msdn.microsoft.com/en-us/library/bb148567).
        */
        long GetMarkerForegroundColorIndex();

        /**
        Sets the marker foreground color as an index into the current color palette, or as one of the following XlColorIndex constants: xlColorIndexAutomatic or xlColorIndexNone. Applies only to line, scatter, and radar charts.

        [MSDN documentation for LegendKey.MarkerForegroundColorIndex](http://msdn.microsoft.com/en-us/library/bb148567).
        */
        void SetMarkerForegroundColorIndex(long markerForegroundColorIndex);

        /**
        Returns the data-marker size, in points. Can be a value from 2 through 72.

        [MSDN documentation for LegendKey.MarkerSize](http://msdn.microsoft.com/en-us/library/bb148570).
        */
        long GetMarkerSize();

        /**
        Sets the data-marker size, in points. Can be a value from 2 through 72.

        [MSDN documentation for LegendKey.MarkerSize](http://msdn.microsoft.com/en-us/library/bb148570).
        */
        void SetMarkerSize(long markerSize);

        /**
        Returns the marker style for a point or series in a line chart, scatter chart, or radar chart. Read/write XlMarkerStyle.

        [MSDN documentation for LegendKey.MarkerStyle](http://msdn.microsoft.com/en-us/library/bb148575).
        */
        XlMarkerStyle GetMarkerStyle();

        /**
        Sets the marker style for a point or series in a line chart, scatter chart, or radar chart. Read/write XlMarkerStyle.

        [MSDN documentation for LegendKey.MarkerStyle](http://msdn.microsoft.com/en-us/library/bb148575).
        */
        void SetMarkerStyle(XlMarkerStyle markerStyle);

        /**
        Returns a XlChartPictureType value that represents the way pictures are displayed on a legend key.

        [MSDN documentation for LegendKey.PictureType](http://msdn.microsoft.com/en-us/library/bb148577).
        */
        XlChartPictureType GetPictureType();

        /**
        Sets a XlChartPictureType value that represents the way pictures are displayed on a legend key.

        [MSDN documentation for LegendKey.PictureType](http://msdn.microsoft.com/en-us/library/bb148577).
        */
        void SetPictureType(XlChartPictureType pictureType);

        /**
        Returns the unit for each picture on the chart if the PictureType property is set to xlStackScale (if not, this property is ignored).

        [MSDN documentation for LegendKey.PictureUnit](http://msdn.microsoft.com/en-us/library/bb148579).
        */
        long GetPictureUnit();

        /**
        Sets the unit for each picture on the chart if the PictureType property is set to xlStackScale (if not, this property is ignored).

        [MSDN documentation for LegendKey.PictureUnit](http://msdn.microsoft.com/en-us/library/bb148579).
        */
        void SetPictureUnit(long pictureUnit);

        /**
        Returns the unit for each picture on the chart if the PictureType property is set to xlStackScale (if not, this property is ignored). Since Excel 2007.

        [MSDN documentation for LegendKey.PictureUnit2](http://msdn.microsoft.com/en-us/library/bb240159).
        */
        double GetPictureUnit2();

        /**
        Sets the unit for each picture on the chart if the PictureType property is set to xlStackScale (if not, this property is ignored). Since Excel 2007.

        [MSDN documentation for LegendKey.PictureUnit2](http://msdn.microsoft.com/en-us/library/bb240159).
        */
        void SetPictureUnit2(double pictureUnit2);

        /**
        Returns a Boolean value that determines if the object has a shadow.

        [MSDN documentation for LegendKey.Shadow](http://msdn.microsoft.com/en-us/library/bb214641).
        */
        bool GetShadow();

        /**
        Sets a Boolean value that determines if the object has a shadow.

        [MSDN documentation for LegendKey.Shadow](http://msdn.microsoft.com/en-us/library/bb214641).
        */
        void SetShadow(bool shadow);

        /**
        True if curve smoothing is turned on for the legend key.

        [MSDN documentation for LegendKey.Smooth](http://msdn.microsoft.com/en-us/library/bb214643).
        */
        bool GetSmooth();

        /**
        True if curve smoothing is turned on for the legend key.

        [MSDN documentation for LegendKey.Smooth](http://msdn.microsoft.com/en-us/library/bb214643).
        */
        void SetSmooth(bool smooth);

        /**
        Returns a Double value that represents the distance, in points, from the top edge of the object to the top of row 1 (on a worksheet) or the top of the chart area (on a chart).

        [MSDN documentation for LegendKey.Top](http://msdn.microsoft.com/en-us/library/bb214644).
        */
        double GetTop();

        /**
        Returns a Double value that represents the width, in points, of the object.

        [MSDN documentation for LegendKey.Width](http://msdn.microsoft.com/en-us/library/bb214646).
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
