/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// License:     MIT license
/////////////////////////////////////////////////////////////////////////////


#ifndef _WXAUTOEXCEL_SERIESCOLLECTION_H
#define _WXAUTOEXCEL_SERIESCOLLECTION_H

#include "wx/wxAutoExcel_defs.h"

#if WXAUTOEXCEL_USE_CHARTS

#include "wx/wxAutoExcel_object.h"
#include "wx/wxAutoExcel_enums.h"

namespace wxAutoExcel {

    /**
    @brief Represents Microsoft Excel SeriesCollection collection.
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelSeriesCollection : public wxExcelObject
    {
    public:
        // ***** METHODS *****

        /**
        Adds one or more new series to the SeriesCollection collection.

        [MSDN documentation for SeriesCollection.Add](https://learn.microsoft.com/en-us/office/vba/api/excel.seriescollection.add).
        */
        wxExcelSeries Add(wxExcelRange source, XlRowCol* rowCol = NULL,
                          wxXlTribool seriesLabels = wxDefaultXlTribool, wxXlTribool categoryLabels = wxDefaultXlTribool,
                          wxXlTribool replace = wxDefaultXlTribool);

        /**
        Adds new data points to an existing series collection.

        [MSDN documentation for SeriesCollection.Extend](https://learn.microsoft.com/en-us/office/vba/api/excel.seriescollection.extend).
        */
        bool Extend(wxExcelRange source, XlRowCol* rowCol = NULL, wxXlTribool categoryLabels = wxDefaultXlTribool);

        //@{
        /**
        Returns a single object from a collection.

        [MSDN documentation for SeriesCollection.Item](https://learn.microsoft.com/en-us/office/vba/api/excel.seriescollection.item).
        */
        wxExcelSeries Item(long index);
        wxExcelSeries operator[](long index);
        //@}

        /**
        Creates a new series. Returns a Series object that represents the new series.

        [MSDN documentation for SeriesCollection.NewSeries](https://learn.microsoft.com/en-us/office/vba/api/excel.seriescollection.newseries).
        */
        wxExcelSeries NewSeries();

        /**
        Pastes data from the Clipboard into the specified series collection.

        [MSDN documentation for SeriesCollection.Paste](https://learn.microsoft.com/en-us/office/vba/api/excel.seriescollection.paste).
        */
        bool Paste(XlRowCol* rowcol = NULL,
                   wxXlTribool seriesLabels = wxDefaultXlTribool, wxXlTribool categoryLabels = wxDefaultXlTribool,
                   wxXlTribool replace = wxDefaultXlTribool, wxXlTribool newSeries = wxDefaultXlTribool);

        // ***** PROPERTIES *****

        /**
        Returns a Long value that represents the number of objects in the collection.

        [MSDN documentation for SeriesCollection.Count](https://learn.microsoft.com/en-us/office/vba/api/excel.seriescollection.count).
        */
        long GetCount();

        /**
        Returns "SeriesCollection".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("SeriesCollection"); }
    };

   /**
   Represents the full set of Series objects in the chart.
   */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelFullSeriesCollection : public wxExcelObject
   {
   public:
       // ***** METHODS *****

       //@{
       /**
       Returns a single object from a collection.

       [MSDN documentation for FullSeriesCollection.Item](https://learn.microsoft.com/en-us/office/vba/api/excel.fullseriescollection.item).
       */
       wxExcelSeries Item(long index);
       wxExcelSeries operator[](long index);
       //@}

       // ***** PROPERTIES *****

       /**
       Returns a Long value that represents the number of objects in the collection.

       [MSDN documentation for FullSeriesCollection.Count](https://learn.microsoft.com/en-us/office/vba/api/excel.fullseriescollection.count).
       */
       long GetCount();

       /**
       Returns "FullSeriesCollection".
       */
       virtual wxString GetAutoExcelObjectName_() const { return wxS("FullSeriesCollection"); }
   };


} // namespace wxAutoExcel

#endif // #if WXAUTOEXCEL_USE_CHARTS

#endif // _WXAUTOEXCEL_SERIESCOLLECTION_H
