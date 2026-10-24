/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// License:     MIT license
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

        [MSDN documentation for Charts.Add](https://learn.microsoft.com/en-us/previous-versions/office/developer/office-2007/bb211680%28v=office.12%29).
        */
        wxExcelChart Add();
        wxExcelChart AddAfterOrBefore(wxExcelSheet sheetAfterOrBefore, bool after);
        //@}

        /**
        Inserts a chart directly onto the grid.

        [Excel VBA documentation for Charts.Add2](https://learn.microsoft.com/en-us/office/vba/api/excel.charts.add2)
        */
        wxExcelChart Add2(wxExcelSheet* before = NULL, wxExcelSheet* after = NULL, long* count = NULL, 
                          wxXlTribool newLayout = wxDefaultXlTribool);

        // ***** PROPERTIES *****


        /**
        Returns a Long value that represents the number of objects in the collection.

        [MSDN documentation for Charts.Count](https://learn.microsoft.com/en-us/office/vba/api/excel.charts.count).
        */
        long GetCount();

        //@{
        /**
        Returns a single object from a collection.

        [MSDN documentation for Charts.Item](https://learn.microsoft.com/en-us/office/vba/api/excel.charts.item).
        */
        wxExcelChart GetItem(long index);
        wxExcelChart operator[](long index);
        //@}

        /**
        The iterator over the items, see wxExcelCollectionIterator.
        */
        typedef wxExcelCollectionIterator<wxExcelCharts, wxExcelChart> iterator;

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
        Returns "Charts".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("Charts"); }

private:
        wxExcelChart DoAdd(wxExcelSheet* sheetAfterOrBefore, bool after);
    };


} // namespace wxAutoExcel

#endif // #if WXAUTOEXCEL_USE_CHARTS

#endif // _WXAUTOEXCEL_CHART_H
