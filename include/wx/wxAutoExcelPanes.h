/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// License:     MIT license
/////////////////////////////////////////////////////////////////////////////


#ifndef _WXAUTOEXCEL_PANES_H
#define _WXAUTOEXCEL_PANES_H

#include "wx/wxAutoExcel_defs.h"

#include "wx/wxAutoExcel_object.h"

namespace wxAutoExcel {

    /**
    @brief Represents Microsoft Excel Pane object.
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelPane : public wxExcelObject
    {
    public:
        // ***** METHODS *****

        /**
        Activates the pane.

        [MSDN documentation for Pane.Activate](https://learn.microsoft.com/en-us/office/vba/api/excel.pane.activate).
        */
        bool Activate();

        /**
        Scrolls the contents of the window by pages.

        [MSDN documentation for Pane.LargeScroll](https://learn.microsoft.com/en-us/office/vba/api/excel.pane.largescroll).
        */
        bool LargeScroll(long* down = NULL, long* up = NULL, long* toRight = NULL, long* toLeft = NULL);

        /**
        Returns or sets a pixel point on the screen.

        [MSDN documentation for Pane.PointsToScreenPixelsX](https://learn.microsoft.com/en-us/office/vba/api/excel.pane.pointstoscreenpixelsx).
        */
        long PointsToScreenPixelsX(double points);

        /**
        Returns or sets the location of the pixel on the screen.

        [MSDN documentation for Pane.PointsToScreenPixelsY](https://learn.microsoft.com/en-us/office/vba/api/excel.pane.pointstoscreenpixelsy).
        */
        long PointsToScreenPixelsY(double points);

        /**
        Scrolls the document window so that the contents of a specified rectangular area are displayed in either the upper-left or lower-right corner of the documentpane (depending on the value of the Start argument).

        [MSDN documentation for Pane.ScrollIntoView](https://learn.microsoft.com/en-us/office/vba/api/excel.pane.scrollintoview).
        */
        void ScrollIntoView(long left, long top, long width, long height, wxXlTribool start = wxDefaultXlTribool);

        /**
        Scrolls the contents of the window by rows or columns.

        [MSDN documentation for Pane.SmallScroll](https://learn.microsoft.com/en-us/office/vba/api/excel.pane.smallscroll).
        */
        bool SmallScroll(long* down = NULL, long* up = NULL, long* toRight = NULL, long* toLeft = NULL);

        // ***** PROPERTIES *****

        /**
        Returns the index number of the window in the Panescollection.

        [MSDN documentation for Pane.Index](https://learn.microsoft.com/en-us/office/vba/api/excel.pane.index).
        */
        long GetIndex();

        /**
        Returns the number of the leftmost column in the pane.

        [MSDN documentation for Pane.ScrollColumn](https://learn.microsoft.com/en-us/office/vba/api/excel.pane.scrollcolumn).
        */
        long GetScrollColumn();

        /**
        Sets the number of the leftmost column in the pane.

        [MSDN documentation for Pane.ScrollColumn](https://learn.microsoft.com/en-us/office/vba/api/excel.pane.scrollcolumn).
        */
        void SetScrollColumn(long scrollColumn);

        /**
        Returns the number of the row that appears at the top of the pane.

        [MSDN documentation for Pane.ScrollRow](https://learn.microsoft.com/en-us/office/vba/api/excel.pane.scrollrow).
        */
        long GetScrollRow();

        /**
        Sets the number of the row that appears at the top of the pane.

        [MSDN documentation for Pane.ScrollRow](https://learn.microsoft.com/en-us/office/vba/api/excel.pane.scrollrow).
        */
        void SetScrollRow(long scrollRow);

        /**
        Returns a Range Represents the range of cells that are visible in the pane. If a column or row is partially visible, it's included in the range.

        [MSDN documentation for Pane.VisibleRange](https://learn.microsoft.com/en-us/office/vba/api/excel.pane.visiblerange).
        */
        wxExcelRange GetVisibleRange();
        /**
        Returns "Pane".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("Pane"); }
    };

    /**
    @brief Represents Microsoft Excel Panes collection.
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelPanes : public wxExcelObject
    {
    public:
        // ***** PROPERTIES *****

        /**
        Returns the number of panes in the collection.

        [MSDN documentation for Panes.Count](https://learn.microsoft.com/en-us/office/vba/api/excel.panes.count).
        */
        long GetCount();

        //@{
        /**
        Returns the Pane.

        [MSDN documentation for Panes.Item](https://learn.microsoft.com/en-us/office/vba/api/excel.panes.item).
        */
        wxExcelPane GetItem(long index);
        wxExcelPane operator[](long index);
        //@}

        /**
        The iterator over the items, see wxExcelCollectionIterator.
        */
        typedef wxExcelCollectionIterator<wxExcelPanes, wxExcelPane> iterator;

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
        Returns "Panes".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("Panes"); }
    };


} // namespace wxAutoExcel

#endif //_WXAUTOEXCEL_PANES_H
