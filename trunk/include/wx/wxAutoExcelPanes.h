/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Modified by:
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
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

        [MSDN documentation for Pane.Activate](http://msdn.microsoft.com/en-us/library/bb211945).
        */
        bool Activate();

        /**
        Scrolls the contents of the window by pages.

        [MSDN documentation for Pane.LargeScroll](http://msdn.microsoft.com/en-us/library/bb211949).
        */
        bool LargeScroll(long* down = NULL, long* up = NULL, long* toRight = NULL, long* toLeft = NULL);

        /**
        Returns or sets a pixel point on the screen.

        [MSDN documentation for Pane.PointsToScreenPixelsX](http://msdn.microsoft.com/en-us/library/bb242081).
        */
        long PointsToScreenPixelsX(double points);

        /**
        Returns or sets the location of the pixel on the screen.

        [MSDN documentation for Pane.PointsToScreenPixelsY](http://msdn.microsoft.com/en-us/library/bb242084).
        */
        long PointsToScreenPixelsY(double points);

        /**
        Scrolls the document window so that the contents of a specified rectangular area are displayed in either the upper-left or lower-right corner of the documentpane (depending on the value of the Start argument).

        [MSDN documentation for Pane.ScrollIntoView](http://msdn.microsoft.com/en-us/library/bb238051).
        */
        void ScrollIntoView(long left, long top, long width, long height, wxXlTribool start = wxDefaultXlTribool);

        /**
        Scrolls the contents of the window by rows or columns.

        [MSDN documentation for Pane.SmallScroll](http://msdn.microsoft.com/en-us/library/bb238054).
        */
        bool SmallScroll(long* down = NULL, long* up = NULL, long* toRight = NULL, long* toLeft = NULL);

        // ***** PROPERTIES *****

        /**
        Returns the index number of the window in the Panescollection.

        [MSDN documentation for Pane.Index](http://msdn.microsoft.com/en-us/library/bb213223).
        */
        long GetIndex();

        /**
        Returns the number of the leftmost column in the pane.

        [MSDN documentation for Pane.ScrollColumn](http://msdn.microsoft.com/en-us/library/bb214930).
        */
        long GetScrollColumn();

        /**
        Sets the number of the leftmost column in the pane.

        [MSDN documentation for Pane.ScrollColumn](http://msdn.microsoft.com/en-us/library/bb214930).
        */
        void SetScrollColumn(long scrollColumn);

        /**
        Returns the number of the row that appears at the top of the pane.

        [MSDN documentation for Pane.ScrollRow](http://msdn.microsoft.com/en-us/library/bb214932).
        */
        long GetScrollRow();

        /**
        Sets the number of the row that appears at the top of the pane.

        [MSDN documentation for Pane.ScrollRow](http://msdn.microsoft.com/en-us/library/bb214932).
        */
        void SetScrollRow(long scrollRow);

        /**
        Returns a Range Represents the range of cells that are visible in the pane. If a column or row is partially visible, it's included in the range.

        [MSDN documentation for Pane.VisibleRange](http://msdn.microsoft.com/en-us/library/bb214934).
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

        [MSDN documentation for Panes.Count](http://msdn.microsoft.com/en-us/library/bb213225).
        */
        long GetCount();

        //@{
        /**
        Returns the Pane.

        [MSDN documentation for Panes.Item](http://msdn.microsoft.com/en-us/library/bb213228).
        */
        wxExcelPane GetItem(long index);
        wxExcelPane operator[](long index);
        //@}

        /**
        Returns "Panes".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("Panes"); }
    };


} // namespace wxAutoExcel

#endif //_WXAUTOEXCEL_PANES_H
