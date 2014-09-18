/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Modified by:
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////


#ifndef _WXAUTOEXCEL_WINDOWS_H
#define _WXAUTOEXCEL_WINDOWS_H

#include "wx/wxAutoExcel_defs.h"

#include "wx/wxAutoExcelObject.h"
#include "wx/wxAutoExcel_enums.h"

#include "wx/wxAutoExcelSheetViews.h"

namespace wxAutoExcel {

    /**
    @brief Represents Microsoft Excel Window object.
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelWindow : public wxExcelObject
    {
    public:
        // ***** METHODS *****

        /**
        Brings the window to the front of the z-order.

        [MSDN documentation for Window.Activate](http://msdn.microsoft.com/en-us/library/bb179123).
        */
        bool Activate();

        /**
        Activates the specified window and then sends it to the back of the window z-order.

        [MSDN documentation for Window.ActivateNext](http://msdn.microsoft.com/en-us/library/bb209536).
        */
        bool ActivateNext();

        /**
        Activates the specified window and then activates the window at the back of the window z-order.

        [MSDN documentation for Window.ActivatePrevious](http://msdn.microsoft.com/en-us/library/bb209540).
        */
        bool ActivatePrevious();

        /**
        Closes the object.

        [MSDN documentation for Window.Close](http://msdn.microsoft.com/en-us/library/bb179128).
        */
        bool Close();

        /**
        Scrolls the contents of the window by pages.

        [MSDN documentation for Window.LargeScroll](http://msdn.microsoft.com/en-us/library/bb179130).
        */
        bool LargeScroll(long* down = NULL, long* up = NULL, long* toRight = NULL, long* toLeft = NULL);

        /**
        Creates a new window or a copy of the specified window.

        [MSDN documentation for Window.NewWindow](http://msdn.microsoft.com/en-us/library/bb179133).
        */
        wxExcelWindow NewWindow();

        /**
        Converts a horizontal measurement from points (document coordinates) to screen pixels (screen coordinates). Returns the converted measurement as a Long value.

        [MSDN documentation for Window.PointsToScreenPixelsX](http://msdn.microsoft.com/en-us/library/bb223538).
        */
        long PointsToScreenPixelsX(double points);

        /**
        Converts a vertical measurement from points (document coordinates) to screen pixels (screen coordinates). Returns the converted measurement as a Long value.

        [MSDN documentation for Window.PointsToScreenPixelsY](http://msdn.microsoft.com/en-us/library/bb223542).
        */
        long PointsToScreenPixelsY(double points);

        //@{
        /**
        Prints the object.

        [MSDN documentation for Window.PrintOut](http://msdn.microsoft.com/en-us/library/bb179137).
        */
        bool PrintOut(long* from = NULL, long* to = NULL, long* copies = NULL, wxXlTribool preview = wxDefaultXlTribool,
            const wxString& activePrinter = wxEmptyString, wxXlTribool printToFile = wxDefaultXlTribool,
            wxXlTribool collate = wxDefaultXlTribool, const wxString& prToFileName= wxEmptyString);
        bool PrintOut(const wxVariantVector& args);
        //@}

        /**
        Shows a preview of the object as it would look when printed.

        [MSDN documentation for Window.PrintPreview](http://msdn.microsoft.com/en-us/library/bb179141).
        */
        bool PrintPreview(wxXlTribool enableChanges = wxDefaultXlTribool);

#if WXAUTOEXCEL_USE_SHAPES
        /**
        Returns Range or Shape that is positioned at the specified pair of screen coordinates, if there's any.

        [MSDN documentation for Window.RangeFromPoint](http://msdn.microsoft.com/en-us/library/bb223563).
        */
        bool RangeFromPoint(long x, long y, wxExcelRange& range, wxExcelShape& shape);
#endif // #if WXAUTOEXCEL_USE_SHAPES

        /**
        Scrolls the document window so that the contents of a specified rectangular area are displayed in either the upper-left or lower-right corner of the document window or pane (depending on the value of the Start argument).

        [MSDN documentation for Window.ScrollIntoView](http://msdn.microsoft.com/en-us/library/bb214109).
        */
        void ScrollIntoView(long left, long top, long width, long height, wxXlTribool start = wxDefaultXlTribool);

        /**
        Scrolls through the workbook tabs at the bottom of the window. Doesn't affect the active sheet in the workbook.

        [MSDN documentation for Window.ScrollWorkbookTabs](http://msdn.microsoft.com/en-us/library/bb178015).
        */
        bool ScrollWorkbookTabs(long* sheets = NULL, long* position = NULL);

        /**
        Scrolls the contents of the window by rows or columns.

        [MSDN documentation for Window.SmallScroll](http://msdn.microsoft.com/en-us/library/bb214119).
        */
        bool SmallScroll(long* down = NULL, long* up = NULL, long* toRight = NULL, long* toLeft = NULL);


        // ***** PROPERTIES *****

        /**
        Returns a Range Represents the active cell in the active window (the window on top) or in the specified window. If the window isn't displaying a worksheet, this property fails.

        [MSDN documentation for Window.ActiveCell](http://msdn.microsoft.com/en-us/library/bb148698).
        */
        wxExcelRange GetActiveCell();

        /**
        Returns a Pane Represents the active pane in the window.

        [MSDN documentation for Window.ActivePane](http://msdn.microsoft.com/en-us/library/bb220817).
        */
        wxExcelPane GetActivePane();

        /**
        Returns an Represents the active sheet (the sheet on top) in the active workbook or in the specified window or workbook. Returns Nothing if no sheet is active.

        [MSDN documentation for Window.ActiveSheet](http://msdn.microsoft.com/en-us/library/bb148714).
        */
        wxExcelSheet GetActiveSheet();

        /**
        True if the auto filter for date grouping is currently displayed in the specified window. Since Excel 2007.

        [MSDN documentation for Window.AutoFilterDateGrouping](http://msdn.microsoft.com/en-us/library/bb242471).
        */
        bool GetAutoFilterDateGrouping();

        /**
        True if the auto filter for date grouping is currently displayed in the specified window.

        [MSDN documentation for Window.AutoFilterDateGrouping](http://msdn.microsoft.com/en-us/library/bb242471).
        */
        void SetAutoFilterDateGrouping(bool autoFilterDateGrouping);

        /**
        Returns the name that appears in the title bar of the document window.

        [MSDN documentation for Window.Caption](http://msdn.microsoft.com/en-us/library/bb148721).
        */
        wxString GetCaption();

        /**
        Sets the name that appears in the title bar of the document window.

        [MSDN documentation for Window.Caption](http://msdn.microsoft.com/en-us/library/bb148721).
        */
        void SetCaption(const wxString& caption);

        /**
        True if the window is displaying formulas; False if the window is displaying values.

        [MSDN documentation for Window.DisplayFormulas](http://msdn.microsoft.com/en-us/library/bb216035).
        */
        bool GetDisplayFormulas();

        /**
        True if the window is displaying formulas; False if the window is displaying values.

        [MSDN documentation for Window.DisplayFormulas](http://msdn.microsoft.com/en-us/library/bb216035).
        */
        void SetDisplayFormulas(bool displayFormulas);

        /**
        True if gridlines are displayed.

        [MSDN documentation for Window.DisplayGridlines](http://msdn.microsoft.com/en-us/library/bb216038).
        */
        bool GetDisplayGridlines();

        /**
        True if gridlines are displayed.

        [MSDN documentation for Window.DisplayGridlines](http://msdn.microsoft.com/en-us/library/bb216038).
        */
        void SetDisplayGridlines(bool displayGridlines);

        /**
        True if both row and column headings are displayed; False if no headings are displayed.

        [MSDN documentation for Window.DisplayHeadings](http://msdn.microsoft.com/en-us/library/bb216039).
        */
        bool GetDisplayHeadings();

        /**
        True if both row and column headings are displayed; False if no headings are displayed.

        [MSDN documentation for Window.DisplayHeadings](http://msdn.microsoft.com/en-us/library/bb216039).
        */
        void SetDisplayHeadings(bool displayHeadings);

        /**
        True if the horizontal scroll bar is displayed.

        [MSDN documentation for Window.DisplayHorizontalScrollBar](http://msdn.microsoft.com/en-us/library/bb177508).
        */
        bool GetDisplayHorizontalScrollBar();

        /**
        True if the horizontal scroll bar is displayed.

        [MSDN documentation for Window.DisplayHorizontalScrollBar](http://msdn.microsoft.com/en-us/library/bb177508).
        */
        void SetDisplayHorizontalScrollBar(bool displayHorizontalScrollBar);

        /**
        True if outline symbols are displayed.

        [MSDN documentation for Window.DisplayOutline](http://msdn.microsoft.com/en-us/library/bb216040).
        */
        bool GetDisplayOutline();

        /**
        True if outline symbols are displayed.

        [MSDN documentation for Window.DisplayOutline](http://msdn.microsoft.com/en-us/library/bb216040).
        */
        void SetDisplayOutline(bool displayOutline);

        /**
        True if the specified window is displayed from right to left instead of from left to right. False if the object is displayed from left to right.

        [MSDN documentation for Window.DisplayRightToLeft](http://msdn.microsoft.com/en-us/library/bb148724).
        */
        bool GetDisplayRightToLeft();

        /**
        True if the specified window is displayed from right to left instead of from left to right. False if the object is displayed from left to right.

        [MSDN documentation for Window.DisplayRightToLeft](http://msdn.microsoft.com/en-us/library/bb148724).
        */
        void SetDisplayRightToLeft(bool displayRightToLeft);

        /**
        True if a ruler is displayed for the specified window. Since Excel 2007.

        [MSDN documentation for Window.DisplayRuler](http://msdn.microsoft.com/en-us/library/bb216041).
        */
        bool GetDisplayRuler();

        /**
        True if a ruler is displayed for the specified window.

        [MSDN documentation for Window.DisplayRuler](http://msdn.microsoft.com/en-us/library/bb216041).
        */
        void SetDisplayRuler(bool displayRuler);

        /**
        True if the vertical scroll bar is displayed.

        [MSDN documentation for Window.DisplayVerticalScrollBar](http://msdn.microsoft.com/en-us/library/bb221016).
        */
        bool GetDisplayVerticalScrollBar();

        /**
        True if the vertical scroll bar is displayed.

        [MSDN documentation for Window.DisplayVerticalScrollBar](http://msdn.microsoft.com/en-us/library/bb221016).
        */
        void SetDisplayVerticalScrollBar(bool displayVerticalScrollBar);

        /**
        True if whitespace is displayed.  Since Excel 2007.

        [MSDN documentation for Window.DisplayWhitespace](http://msdn.microsoft.com/en-us/library/bb242644).
        */
        bool GetDisplayWhitespace();

        /**
        True if whitespace is displayed. 

        [MSDN documentation for Window.DisplayWhitespace](http://msdn.microsoft.com/en-us/library/bb242644).
        */
        void SetDisplayWhitespace(bool displayWhitespace);

        /**
        True if the workbook tabs are displayed.

        [MSDN documentation for Window.DisplayWorkbookTabs](http://msdn.microsoft.com/en-us/library/bb221019).
        */
        bool GetDisplayWorkbookTabs();

        /**
        True if the workbook tabs are displayed.

        [MSDN documentation for Window.DisplayWorkbookTabs](http://msdn.microsoft.com/en-us/library/bb221019).
        */
        void SetDisplayWorkbookTabs(bool displayWorkbookTabs);

        /**
        True if zero values are displayed. 

        [MSDN documentation for Window.DisplayZeros](http://msdn.microsoft.com/en-us/library/bb216042).
        */
        bool GetDisplayZeros();

        /**
        True if zero values are displayed. 

        [MSDN documentation for Window.DisplayZeros](http://msdn.microsoft.com/en-us/library/bb216042).
        */
        void SetDisplayZeros(bool displayZeros);

        /**
        True if the window can be resized.

        [MSDN documentation for Window.EnableResize](http://msdn.microsoft.com/en-us/library/bb221166).
        */
        bool GetEnableResize();

        /**
        True if the window can be resized.

        [MSDN documentation for Window.EnableResize](http://msdn.microsoft.com/en-us/library/bb221166).
        */
        void SetEnableResize(bool enableResize);

        /**
        True if split panes are frozen.

        [MSDN documentation for Window.FreezePanes](http://msdn.microsoft.com/en-us/library/bb208531).
        */
        bool GetFreezePanes();

        /**
        True if split panes are frozen.

        [MSDN documentation for Window.FreezePanes](http://msdn.microsoft.com/en-us/library/bb208531).
        */
        void SetFreezePanes(bool freezePanes);

        /**
        Returns the gridline color as an RGB value.

        [MSDN documentation for Window.GridlineColor](http://msdn.microsoft.com/en-us/library/bb208583).
        */
        wxColour GetGridlineColor();

        /**
        Sets the gridline color as an RGB value.

        [MSDN documentation for Window.GridlineColor](http://msdn.microsoft.com/en-us/library/bb208583).
        */
        void SetGridlineColor(const wxColour& gridlineColor);

        /**
        Returns the gridline color as an index into the current color palette or as the following XlColorIndex constant.

        [MSDN documentation for Window.GridlineColorIndex](http://msdn.microsoft.com/en-us/library/bb208585).
        */
        long GetGridlineColorIndex();

        /**
        Sets the gridline color as an index into the current color palette or as the following XlColorIndex constant.

        [MSDN documentation for Window.GridlineColorIndex](http://msdn.microsoft.com/en-us/library/bb208585).
        */
        void SetGridlineColorIndex(long gridlineColorIndex);

        /**
        Returns a Double value that represents tThe height, in points, of the window.

        [MSDN documentation for Window.Height](http://msdn.microsoft.com/en-us/library/bb148730).
        */
        double GetHeight();

        /**
        Sets a Double value that represents tThe height, in points, of the window.

        [MSDN documentation for Window.Height](http://msdn.microsoft.com/en-us/library/bb148730).
        */
        void SetHeight(double height);

        /**
        Returns a Long value that represents the index number of the object within the collection of similar objects.

        [MSDN documentation for Window.Index](http://msdn.microsoft.com/en-us/library/bb148733).
        */
        long GetIndex();

        /**
        Returns a Double value that represents the distance, in points, from the left edge of the client area to the left edge of the window.

        [MSDN documentation for Window.Left](http://msdn.microsoft.com/en-us/library/bb148736).
        */
        double GetLeft();

        /**
        Sets a Double value that represents the distance, in points, from the left edge of the client area to the left edge of the window.

        [MSDN documentation for Window.Left](http://msdn.microsoft.com/en-us/library/bb148736).
        */
        void SetLeft(double left);

        /**
        Returns a Panes collection that represents all the panes in the specified window.

        [MSDN documentation for Window.Panes](http://msdn.microsoft.com/en-us/library/bb208940).
        */
        wxExcelPanes GetPanes();

        /**
        Returns a Range Represents the selected cells on the worksheet in the specified window even if a graphic object is active or selected on the worksheet.

        [MSDN documentation for Window.RangeSelection](http://msdn.microsoft.com/en-us/library/bb209056).
        */
        wxExcelRange GetRangeSelection();


        /**
        Returns the number of the leftmost column in the pane or window.

        [MSDN documentation for Window.ScrollColumn](http://msdn.microsoft.com/en-us/library/bb215194).
        */
        long GetScrollColumn();

        /**
        Sets the number of the leftmost column in the pane or window.

        [MSDN documentation for Window.ScrollColumn](http://msdn.microsoft.com/en-us/library/bb215194).
        */
        void SetScrollColumn(long scrollColumn);

        /**
        Returns the number of the row that appears at the top of the pane or window.

        [MSDN documentation for Window.ScrollRow](http://msdn.microsoft.com/en-us/library/bb215197).
        */
        long GetScrollRow();

        /**
        Sets the number of the row that appears at the top of the pane or window.

        [MSDN documentation for Window.ScrollRow](http://msdn.microsoft.com/en-us/library/bb215197).
        */
        void SetScrollRow(long scrollRow);

        /**
        Returns a Sheets collection that represents all the selected sheets in the specified window.

        [MSDN documentation for Window.SelectedSheets](http://msdn.microsoft.com/en-us/library/bb221639).
        */
        wxExcelSheets GetSelectedSheets();

        /**
        Returns the SheetViews object for the specified window. Since Excel 2007.

        [MSDN documentation for Window.SheetViews](http://msdn.microsoft.com/en-us/library/bb216049).
        */
        wxExcelSheetViews GetSheetViews();
        

        /**
        True if the window is split.

        [MSDN documentation for Window.Split](http://msdn.microsoft.com/en-us/library/bb209269).
        */
        bool GetSplit();

        /**
        True if the window is split.

        [MSDN documentation for Window.Split](http://msdn.microsoft.com/en-us/library/bb209269).
        */
        void SetSplit(bool split);

        /**
        Returns the column number where the window is split into panes (the number of columns to the left of the split line).

        [MSDN documentation for Window.SplitColumn](http://msdn.microsoft.com/en-us/library/bb209271).
        */
        long GetSplitColumn();

        /**
        Sets the column number where the window is split into panes (the number of columns to the left of the split line).

        [MSDN documentation for Window.SplitColumn](http://msdn.microsoft.com/en-us/library/bb209271).
        */
        void SetSplitColumn(long splitColumn);

        /**
        Returns the location of the horizontal window split, in points. Read/write Double.

        [MSDN documentation for Window.SplitHorizontal](http://msdn.microsoft.com/en-us/library/bb209273).
        */
        double GetSplitHorizontal();

        /**
        Sets the location of the horizontal window split, in points. Read/write Double.

        [MSDN documentation for Window.SplitHorizontal](http://msdn.microsoft.com/en-us/library/bb209273).
        */
        void SetSplitHorizontal(double splitHorizontal);

        /**
        Returns the row number where the window is split into panes (the number of rows above the split).

        [MSDN documentation for Window.SplitRow](http://msdn.microsoft.com/en-us/library/bb209275).
        */
        long GetSplitRow();

        /**
        Sets the row number where the window is split into panes (the number of rows above the split).

        [MSDN documentation for Window.SplitRow](http://msdn.microsoft.com/en-us/library/bb209275).
        */
        void SetSplitRow(long splitRow);

        /**
        Returns the location of the vertical window split, in points. Read/write Double.

        [MSDN documentation for Window.SplitVertical](http://msdn.microsoft.com/en-us/library/bb209283).
        */
        double GetSplitVertical();

        /**
        Sets the location of the vertical window split, in points. Read/write Double.

        [MSDN documentation for Window.SplitVertical](http://msdn.microsoft.com/en-us/library/bb209283).
        */
        void SetSplitVertical(double splitVertical);

        /**
        Returns the ratio of the width of the workbook's tab area to the width of the window's horizontal scroll bar (as a number between 0 (zero) and 1; the default value is 0.6). Read/write Double.

        [MSDN documentation for Window.TabRatio](http://msdn.microsoft.com/en-us/library/bb209330).
        */
        double GetTabRatio();

        /**
        Sets the ratio of the width of the workbook's tab area to the width of the window's horizontal scroll bar (as a number between 0 (zero) and 1; the default value is 0.6). Read/write Double.

        [MSDN documentation for Window.TabRatio](http://msdn.microsoft.com/en-us/library/bb209330).
        */
        void SetTabRatio(double tabRatio);

        /**
        Returns a Double value that represents the distance, in points, from the top edge of the window to the top edge of the usable area (below the menus, any toolbars docked at the top, and the formula bar).

        [MSDN documentation for Window.Top](http://msdn.microsoft.com/en-us/library/bb215205).
        */
        double GetTop();

        /**
        Sets a Double value that represents the distance, in points, from the top edge of the window to the top edge of the usable area (below the menus, any toolbars docked at the top, and the formula bar).

        [MSDN documentation for Window.Top](http://msdn.microsoft.com/en-us/library/bb215205).
        */
        void SetTop(double top);

        /**
        Returns a XlWindowType value that represents the window type.

        [MSDN documentation for Window.Type](http://msdn.microsoft.com/en-us/library/bb215209).
        */
        XlWindowType GetType();

        /**
        Sets a XlWindowType value that represents the window type.

        [MSDN documentation for Window.Type](http://msdn.microsoft.com/en-us/library/bb215209).
        */
        void SetType(XlWindowType type);

        /**
        Returns the maximum height of the space that a window can occupy in the application window area, in points. Read-only Double.

        [MSDN documentation for Window.UsableHeight](http://msdn.microsoft.com/en-us/library/bb215213).
        */
        double GetUsableHeight();

        /**
        Returns the maximum width of the space that a window can occupy in the application window area, in points. Read-only Double.

        [MSDN documentation for Window.UsableWidth](http://msdn.microsoft.com/en-us/library/bb215219).
        */
        double GetUsableWidth();

        /**
        Returns the view showing in the window. Read/write XlWindowView.

        [MSDN documentation for Window.View](http://msdn.microsoft.com/en-us/library/bb223019).
        */
        XlWindowView GetView();

        /**
        Sets the view showing in the window. Read/write XlWindowView.

        [MSDN documentation for Window.View](http://msdn.microsoft.com/en-us/library/bb223019).
        */
        void SetView(XlWindowView view);

        /**
        Returns a Boolean value that determines whether the object is visible.

        [MSDN documentation for Window.Visible](http://msdn.microsoft.com/en-us/library/bb215223).
        */
        bool GetVisible();

        /**
        Sets a Boolean value that determines whether the object is visible.

        [MSDN documentation for Window.Visible](http://msdn.microsoft.com/en-us/library/bb215223).
        */
        void SetVisible(bool visible);

        /**
        Returns a Range Represents the range of cells that are visible in the window or pane. If a column or row is partially visible, it's included in the range.

        [MSDN documentation for Window.VisibleRange](http://msdn.microsoft.com/en-us/library/bb215227).
        */
        wxExcelRange GetVisibleRange();

        /**
        Returns a Double value that represents the width, in points, of the window.

        [MSDN documentation for Window.Width](http://msdn.microsoft.com/en-us/library/bb215230).
        */
        double GetWidth();

        /**
        Sets a Double value that represents the width, in points, of the window.

        [MSDN documentation for Window.Width](http://msdn.microsoft.com/en-us/library/bb215230).
        */
        void SetWidth(double width);

        /**
        Returns the window number. For example, a window named "Book1.xls:2" has 2 as its window number. Most windows have the window number 1.

        [MSDN documentation for Window.WindowNumber](http://msdn.microsoft.com/en-us/library/bb223058).
        */
        long GetWindowNumber();

        /**
        Returns the state of the window. Read/write XlWindowState.

        [MSDN documentation for Window.WindowState](http://msdn.microsoft.com/en-us/library/bb215234).
        */
        XlWindowState GetWindowState();

        /**
        Sets the state of the window. Read/write XlWindowState.

        [MSDN documentation for Window.WindowState](http://msdn.microsoft.com/en-us/library/bb215234).
        */
        void SetWindowState(XlWindowState windowState);

        /**
        Returns the display size of the window, as a percentage (100 equals normal size, 200 equals double size, and so on).

        [MSDN documentation for Window.Zoom](http://msdn.microsoft.com/en-us/library/bb215238).
        */
        long GetZoom();

        /**
        Sets the display size of the window, as a percentage (100 equals normal size, 200 equals double size, and so on).

        [MSDN documentation for Window.Zoom](http://msdn.microsoft.com/en-us/library/bb215238).
        */
        void SetZoom(long zoom);


        /**
        Returns "Window".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("Window"); }
    };

    /**
    @brief Represents Microsoft Excel Window Pane collection.
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelWindows : public wxExcelObject
    {
    public:

        // ***** METHODS *****

        /**
        Arranges the windows on the screen.

        [MSDN documentation for Windows.Arrange](http://msdn.microsoft.com/en-us/library/bb209660).
        */
        bool Arrange(XlArrangeStyle* arrangeStyle = NULL, wxXlTribool activeWorkbook =  wxDefaultXlTribool,
            wxXlTribool syncHorizontal = wxDefaultXlTribool, wxXlTribool syncVertica = wxDefaultXlTribool);

        /**
        Ends side-by-side mode if two windows are in side-by-side mode. Returns a Boolean value that represents whether the method was successful.

        [MSDN documentation for Windows.BreakSideBySide](http://msdn.microsoft.com/en-us/library/bb209721).
        */
        bool BreakSideBySide();

        /**
        Opens two windows in side-by-side mode. Returns a Boolean value.

        [MSDN documentation for Windows.CompareSideBySideWith](http://msdn.microsoft.com/en-us/library/bb223274).
        */
        bool CompareSideBySideWith(const wxString& windowName);

        /**
        Resets the position of two worksheet windows that are being compared side by side.

        [MSDN documentation for Windows.ResetPositionsSideBySide](http://msdn.microsoft.com/en-us/library/bb177968).
        */
        void ResetPositionsSideBySide();


        // ***** PROPERTIES *****
        
        /**
        Returns a Long value that represents the number of objects in the collection.

        [MSDN documentation for Windows.Count](http://msdn.microsoft.com/en-us/library/bb148745).
        */
        long GetCount();

        //@{
        /**
        Returns a single object from a collection.

        [MSDN documentation for Windows.Item](http://msdn.microsoft.com/en-us/library/bb148752).
        */
        wxExcelWindow GetItem(long index);
        wxExcelWindow operator[](long index);
        wxExcelWindow GetItem(const wxString& name);
        wxExcelWindow operator[](const wxString& name);
        //@}

        /**
        True enables scrolling the contents of windows at the same time when documents are being compared side by side. False disables scrolling the windows at the same time.

        [MSDN documentation for Windows.SyncScrollingSideBySide](http://msdn.microsoft.com/en-us/library/bb209326).
        */
        bool GetSyncScrollingSideBySide();

        /**
        True enables scrolling the contents of windows at the same time when documents are being compared side by side. False disables scrolling the windows at the same time.

        [MSDN documentation for Windows.SyncScrollingSideBySide](http://msdn.microsoft.com/en-us/library/bb209326).
        */
        void SetSyncScrollingSideBySide(bool syncScrollingSideBySide);


        /**
        Returns "Windows".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("Windows"); }
    };


} // namespace wxAutoExcel

#endif //_WXAUTOEXCEL_WINDOWS_H
