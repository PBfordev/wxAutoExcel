/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Modified by:
// RCS-ID:      $Id: $
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////


#ifndef _WXAUTOEXCEL_SERIESCOLLECTION_H
#define _WXAUTOEXCEL_SERIESCOLLECTION_H

#include "wx/wxAutoExcel_defs.h"

#if WXAUTOEXCEL_USE_CHARTS

#include "wx/wxAutoExcelObject.h"
#include "wx/wxAutoExcel_enums.h"

namespace wxAutoExcel {

    /**
    Represents Microsoft Excel SeriesCollection collection.
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelSeriesCollection : public wxExcelObject
    {
    public:
        // ***** METHODS *****

        /**
        Adds one or more new series to the SeriesCollection collection.

        [MSDN documentation for SeriesCollection.Add](http://msdn.microsoft.com/en-us/library/bb178914).
        */
        wxExcelSeries Add(wxExcelRange source, XlRowCol* rowCol = NULL, 
                          wxXlTribool seriesLabels = wxDefaultXlTribool, wxXlTribool categoryLabels = wxDefaultXlTribool, 
                          wxXlTribool replace = wxDefaultXlTribool);

        /**
        Adds new data points to an existing series collection.

        [MSDN documentation for SeriesCollection.Extend](http://msdn.microsoft.com/en-us/library/bb209834).
        */
        bool Extend(wxExcelRange source, XlRowCol* rowCol = NULL, wxXlTribool categoryLabels = wxDefaultXlTribool);

        //@{
        /**
        Returns a single object from a collection.

        [MSDN documentation for SeriesCollection.Item](http://msdn.microsoft.com/en-us/library/bb178917).
        */
        wxExcelSeries Item(long index);
        wxExcelSeries operator[](long index);
        //@}

        /**
        Creates a new series. Returns a Series object that represents the new series.

        [MSDN documentation for SeriesCollection.NewSeries](http://msdn.microsoft.com/en-us/library/bb223487).
        */
        wxExcelSeries NewSeries();

        /**
        Pastes data from the Clipboard into the specified series collection.

        [MSDN documentation for SeriesCollection.Paste](http://msdn.microsoft.com/en-us/library/bb178922).
        */
        bool Paste(XlRowCol* rowcol = NULL, 
                   wxXlTribool seriesLabels = wxDefaultXlTribool, wxXlTribool categoryLabels = wxDefaultXlTribool, 
                   wxXlTribool replace = wxDefaultXlTribool, wxXlTribool newSeries = wxDefaultXlTribool);

        // ***** PROPERTIES *****

        /**
        Returns a Long value that represents the number of objects in the collection.

        [MSDN documentation for SeriesCollection.Count](http://msdn.microsoft.com/en-us/library/bb237562).
        */
        long GetCount();

        /**
        Returns "SeriesCollection".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("SeriesCollection"); }
    };


} // namespace wxAutoExcel

#endif // #if WXAUTOEXCEL_USE_CHARTS

#endif // _WXAUTOEXCEL_SERIESCOLLECTION_H
