/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Modified by:
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////


#ifndef _WXAUTOEXCEL_LEGENDENTRIES_H
#define _WXAUTOEXCEL_LEGENDENTRIES_H

#include "wx/wxAutoExcel_defs.h"

#if WXAUTOEXCEL_USE_CHARTS

#include "wx/wxAutoExcel_object.h"

namespace wxAutoExcel {

    /**
    Represents Microsoft Excel LegendEntry object.
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelLegendEntry : public wxExcelObject
    {
    public:

        // ***** METHODS *****

        /**
        Deletes the object.

        [MSDN documentation for LegendEntry.Delete](http://msdn.microsoft.com/en-us/library/bb211838).
        */
        bool Delete();

        /**
        Selects the object.

        [MSDN documentation for LegendEntry.Select](http://msdn.microsoft.com/en-us/library/bb237971).
        */
        bool Select();

        // ***** PROPERTIES *****


        /**
        True if the text in the object changes font size when the object size changes. The default value is True. Read/write Variant.

        [MSDN documentation for LegendEntry.AutoScaleFont](http://msdn.microsoft.com/en-us/library/bb148538).
        */
        bool GetAutoScaleFont();

        /**
        True if the text in the object changes font size when the object size changes. The default value is True. Read/write Variant.

        [MSDN documentation for LegendEntry.AutoScaleFont](http://msdn.microsoft.com/en-us/library/bb148538).
        */
        void SetAutoScaleFont(bool autoScaleFont);

        /**
        Returns a Font object that represents the font of the specified object.

        [MSDN documentation for LegendEntry.Font](http://msdn.microsoft.com/en-us/library/bb148540).
        */
        wxExcelFont GetFont();

        /**
        Returns the ChartFormat object. Since Excel 2007.

        [MSDN documentation for LegendEntry.Format](http://msdn.microsoft.com/en-us/library/bb242533).
        */
        wxExcelChartFormat GetFormat();

        /**
        Returns a Double value that represents the height, in points, of the object.

        [MSDN documentation for LegendEntry.Height](http://msdn.microsoft.com/en-us/library/bb148541).
        */
        double GetHeight();

        /**
        Returns a Long value that represents the index number of the object within the collection of similar objects.

        [MSDN documentation for LegendEntry.Index](http://msdn.microsoft.com/en-us/library/bb148542).
        */
        long GetIndex();

        /**
        Returns a Double value that represents the distance, in points, from the left edge of the object to the left edge of the chart area.

        [MSDN documentation for LegendEntry.Left](http://msdn.microsoft.com/en-us/library/bb148544).
        */
        double GetLeft();

        /**
        Returns a LegendKey object that represents the legend key associated with the entry.

        [MSDN documentation for LegendEntry.LegendKey](http://msdn.microsoft.com/en-us/library/bb177909).
        */
        wxExcelLegendKey GetLegendKey();

        /**
        Returns a Double value that represents the distance, in points, from the top edge of the object to the top of row 1 (on a worksheet) or the top of the chart area (on a chart).

        [MSDN documentation for LegendEntry.Top](http://msdn.microsoft.com/en-us/library/bb214638).
        */
        double GetTop();

        /**
        Returns a Double value that represents the width, in points, of the object.

        [MSDN documentation for LegendEntry.Width](http://msdn.microsoft.com/en-us/library/bb214639).
        */
        double GetWidth();

        /**
        Returns "LegendEntry".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("LegendEntry"); }
    };

    /**
    Represents Microsoft Excel LegendEntries collection.
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelLegendEntries : public wxExcelObject
    {
    public:
        // ***** METHODS *****

        //@{
        /**
        Returns a single object from a collection.

        [MSDN documentation for LegendEntries.Item](http://msdn.microsoft.com/en-us/library/bb211835).
        */
        wxExcelLegendEntry Item(long index);
        wxExcelLegendEntry operator[](long index);
        //@}

        // ***** PROPERTIES *****


        /**
        Returns a Long value that represents the number of objects in the collection.

        [MSDN documentation for LegendEntries.Count](http://msdn.microsoft.com/en-us/library/bb148536).
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
