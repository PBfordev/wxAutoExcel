/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Modified by:
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////


#ifndef _WXAUTOEXCEL_CHARTS_H
#define _WXAUTOEXCEL_CHARTS_H

#include "wx/wxAutoExcel_defs.h"

#if WXAUTOEXCEL_USE_CHARTS

#include "wx/wxAutoExcel_object.h"

namespace wxAutoExcel {

    /**
    @brief Represents Microsoft Excel Charts collection.
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelCharts : public wxExcelObject
    {
    public:
        // ***** METHODS *****

        //@{
        /**
        Creates a new chart sheet and returns a Chart object.

        [MSDN documentation for Charts.Add](http://msdn.microsoft.com/en-us/library/bb211680).
        */
        wxExcelChart Add();
        wxExcelChart AddAfterOrBefore(wxExcelSheet sheetAfterOrBefore, bool after);
        //@}

        /**
        Inserts a chart directly onto the grid.

        [Excel VBA documentation for Charts.Add2](https://docs.microsoft.com/en-us/office/vba/api/excel.charts.add2)
        */
        wxExcelChart Add2(wxExcelSheet* before = NULL, wxExcelSheet* after = NULL, long* count = NULL, 
                          wxXlTribool newLayout = wxDefaultXlTribool);

        // ***** PROPERTIES *****


        /**
        Returns a Long value that represents the number of objects in the collection.

        [MSDN documentation for Charts.Count](http://msdn.microsoft.com/en-us/library/bb179491).
        */
        long GetCount();

        //@{
        /**
        Returns a single object from a collection.

        [MSDN documentation for Charts.Item](http://msdn.microsoft.com/en-us/library/bb179496).
        */
        wxExcelChart GetItem(long index);
        wxExcelChart operator[](long index);
        //@}

        /**
        Returns "Charts".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("Charts"); }

private:
        wxExcelChart DoAdd(wxExcelSheet* sheetAfterOrBefore, bool after);
    };


} // namespace wxAutoExcel

#endif // #if WXAUTOEXCEL_USE_CHARTS

#endif // _WXAUTOEXCEL_CHART_H
