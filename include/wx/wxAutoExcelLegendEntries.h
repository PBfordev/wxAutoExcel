/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// License:     MIT license
/////////////////////////////////////////////////////////////////////////////


#ifndef _WXAUTOEXCEL_LEGENDENTRIES_H
#define _WXAUTOEXCEL_LEGENDENTRIES_H

#include "wx/wxAutoExcel_defs.h"

#if WXAUTOEXCEL_USE_CHARTS

#include "wx/wxAutoExcel_object.h"

namespace wxAutoExcel {

    /**
    @brief Represents Microsoft Excel LegendEntry object.
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelLegendEntry : public wxExcelObject
    {
    public:

        // ***** METHODS *****

        /**
        Deletes the object.

        [MSDN documentation for LegendEntry.Delete](https://learn.microsoft.com/en-us/office/vba/api/excel.legendentry.delete).
        */
        bool Delete();

        /**
        Selects the object.

        [MSDN documentation for LegendEntry.Select](https://learn.microsoft.com/en-us/office/vba/api/excel.legendentry.select).
        */
        bool Select();

        // ***** PROPERTIES *****


        /**
        True if the text in the object changes font size when the object size changes. The default value is True. Read/write Variant.

        [MSDN documentation for LegendEntry.AutoScaleFont](https://learn.microsoft.com/en-us/previous-versions/office/developer/office-2007/bb148538%28v=office.12%29).
        */
        bool GetAutoScaleFont();

        /**
        True if the text in the object changes font size when the object size changes. The default value is True. Read/write Variant.

        [MSDN documentation for LegendEntry.AutoScaleFont](https://learn.microsoft.com/en-us/previous-versions/office/developer/office-2007/bb148538%28v=office.12%29).
        */
        void SetAutoScaleFont(bool autoScaleFont);

        /**
        Returns a Font object that represents the font of the specified object.

        [MSDN documentation for LegendEntry.Font](https://learn.microsoft.com/en-us/office/vba/api/excel.legendentry.font).
        */
        wxExcelFont GetFont();

        /**
        Returns the ChartFormat object. Since Excel 2007.

        [MSDN documentation for LegendEntry.Format](https://learn.microsoft.com/en-us/office/vba/api/excel.legendentry.format).
        */
        wxExcelChartFormat GetFormat();

        /**
        Returns a Double value that represents the height, in points, of the object.

        [MSDN documentation for LegendEntry.Height](https://learn.microsoft.com/en-us/office/vba/api/excel.legendentry.height).
        */
        double GetHeight();

        /**
        Returns a Long value that represents the index number of the object within the collection of similar objects.

        [MSDN documentation for LegendEntry.Index](https://learn.microsoft.com/en-us/office/vba/api/excel.legendentry.index).
        */
        long GetIndex();

        /**
        Returns a Double value that represents the distance, in points, from the left edge of the object to the left edge of the chart area.

        [MSDN documentation for LegendEntry.Left](https://learn.microsoft.com/en-us/office/vba/api/excel.legendentry.left).
        */
        double GetLeft();

        /**
        Returns a LegendKey object that represents the legend key associated with the entry.

        [MSDN documentation for LegendEntry.LegendKey](https://learn.microsoft.com/en-us/office/vba/api/excel.legendentry.legendkey).
        */
        wxExcelLegendKey GetLegendKey();

        /**
        Returns a Double value that represents the distance, in points, from the top edge of the object to the top of row 1 (on a worksheet) or the top of the chart area (on a chart).

        [MSDN documentation for LegendEntry.Top](https://learn.microsoft.com/en-us/office/vba/api/excel.legendentry.top).
        */
        double GetTop();

        /**
        Returns a Double value that represents the width, in points, of the object.

        [MSDN documentation for LegendEntry.Width](https://learn.microsoft.com/en-us/office/vba/api/excel.legendentry.width).
        */
        double GetWidth();

        /**
        Returns "LegendEntry".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("LegendEntry"); }
    };

    /**
    @brief Represents Microsoft Excel LegendEntries collection.
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelLegendEntries : public wxExcelObject
    {
    public:
        // ***** METHODS *****

        //@{
        /**
        Returns a single object from a collection.

        [MSDN documentation for LegendEntries.Item](https://learn.microsoft.com/en-us/office/vba/api/excel.legendentries.item).
        */
        wxExcelLegendEntry Item(long index);
        wxExcelLegendEntry operator[](long index);
        //@}

        // ***** PROPERTIES *****


        /**
        Returns a Long value that represents the number of objects in the collection.

        [MSDN documentation for LegendEntries.Count](https://learn.microsoft.com/en-us/office/vba/api/excel.legendentries.count).
        */
        long GetCount();

        /**
        Returns "LegendEntries".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("LegendEntries"); }
    };


} // namespace wxAutoExcel

#endif // #if WXAUTOEXCEL_USE_CHARTS

#endif // _WXAUTOEXCEL_LEGENDENTRIES_H
