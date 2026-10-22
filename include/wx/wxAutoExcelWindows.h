/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// License:     MIT license
/////////////////////////////////////////////////////////////////////////////


#ifndef _WXAUTOEXCEL_WINDOWS_H
#define _WXAUTOEXCEL_WINDOWS_H

#include "wx/wxAutoExcel_defs.h"

#include "wx/wxAutoExcel_object.h"
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

        [MSDN documentation for Window.Activate](https://learn.microsoft.com/en-us/office/vba/api/excel.window.activate).
        */
        void Activate();

        /**
        Activates the specified window and then sends it to the back of the window z-order.

        [MSDN documentation for Window.ActivateNext](https://learn.microsoft.com/en-us/office/vba/api/excel.window.activatenext).
        */
        bool ActivateNext();

        /**
        Activates the specified window and then activates the window at the back of the window z-order.

        [MSDN documentation for Window.ActivatePrevious](https://learn.microsoft.com/en-us/office/vba/api/excel.window.activateprevious).
        */
        bool ActivatePrevious();

        /**
        Closes the object.

        [MSDN documentation for Window.Close](https://learn.microsoft.com/en-us/office/vba/api/excel.window.close).
        */
        bool Close();

        /**
        Scrolls the contents of the window by pages.

        [MSDN documentation for Window.LargeScroll](https://learn.microsoft.com/en-us/office/vba/api/excel.window.largescroll).
        */
        bool LargeScroll(long* down = NULL, long* up = NULL, long* toRight = NULL, long* toLeft = NULL);

        /**
        Creates a new window or a copy of the specified window.

        [MSDN documentation for Window.NewWindow](https://learn.microsoft.com/en-us/office/vba/api/excel.window.newwindow).
        */
        wxExcelWindow NewWindow();

        /**
        Converts a horizontal measurement from points (document coordinates) to screen pixels (screen coordinates). Returns the converted measurement as a Long value.

        [MSDN documentation for Window.PointsToScreenPixelsX](https://learn.microsoft.com/en-us/office/vba/api/excel.window.pointstoscreenpixelsx).
        */
        long PointsToScreenPixelsX(double points);

        /**
        Converts a vertical measurement from points (document coordinates) to screen pixels (screen coordinates). Returns the converted measurement as a Long value.

        [MSDN documentation for Window.PointsToScreenPixelsY](https://learn.microsoft.com/en-us/office/vba/api/excel.window.pointstoscreenpixelsy).
        */
        long PointsToScreenPixelsY(double points);

        //@{
        /**
        Prints the object.

        [MSDN documentation for Window.PrintOut](https://learn.microsoft.com/en-us/office/vba/api/excel.window.printout).
        */
        bool PrintOut(long* from = NULL, long* to = NULL, long* copies = NULL, wxXlTribool preview = wxDefaultXlTribool,
            const wxString& activePrinter = wxEmptyString, wxXlTribool printToFile = wxDefaultXlTribool,
            wxXlTribool collate = wxDefaultXlTribool, const wxString& prToFileName= wxEmptyString);
        bool PrintOut(const wxVariantVector& args);
        //@}

        /**
        Shows a preview of the object as it would look when printed.

        [MSDN documentation for Window.PrintPreview](https://learn.microsoft.com/en-us/office/vba/api/excel.window.printpreview).
        */
        bool PrintPreview(wxXlTribool enableChanges = wxDefaultXlTribool);

#if WXAUTOEXCEL_USE_SHAPES
        /**
        Returns Range or Shape that is positioned at the specified pair of screen coordinates, if there's any.

        [MSDN documentation for Window.RangeFromPoint](https://learn.microsoft.com/en-us/office/vba/api/excel.window.rangefrompoint).
        */
        bool RangeFromPoint(long x, long y, wxExcelRange& range, wxExcelShape& shape);
#endif // #if WXAUTOEXCEL_USE_SHAPES

        /**
        Scrolls the document window so that the contents of a specified rectangular area are displayed in either the upper-left or lower-right corner of the document window or pane (depending on the value of the Start argument).

        [MSDN documentation for Window.ScrollIntoView](https://learn.microsoft.com/en-us/office/vba/api/excel.window.scrollintoview).
        */
        void ScrollIntoView(long left, long top, long width, long height, wxXlTribool start = wxDefaultXlTribool);

        /**
        Scrolls through the workbook tabs at the bottom of the window. Doesn't affect the active sheet in the workbook.

        [MSDN documentation for Window.ScrollWorkbookTabs](https://learn.microsoft.com/en-us/office/vba/api/excel.window.scrollworkbooktabs).
        */
        bool ScrollWorkbookTabs(long* sheets = NULL, long* position = NULL);

        /**
        Scrolls the contents of the window by rows or columns.

        [MSDN documentation for Window.SmallScroll](https://learn.microsoft.com/en-us/office/vba/api/excel.window.smallscroll).
        */
        bool SmallScroll(long* down = NULL, long* up = NULL, long* toRight = NULL, long* toLeft = NULL);


        // ***** PROPERTIES *****

        /**
        Returns a Range Represents the active cell in the active window (the window on top) or in the specified window. If the window isn't displaying a worksheet, this property fails.

        [MSDN documentation for Window.ActiveCell](https://learn.microsoft.com/en-us/office/vba/api/excel.window.activecell).
        */
        wxExcelRange GetActiveCell();

        /**
        Returns a Pane Represents the active pane in the window.

        [MSDN documentation for Window.ActivePane](https://learn.microsoft.com/en-us/office/vba/api/excel.window.activepane).
        */
        wxExcelPane GetActivePane();

        /**
        Returns an Represents the active sheet (the sheet on top) in the active workbook or in the specified window or workbook. Returns Nothing if no sheet is active.

        [MSDN documentation for Window.ActiveSheet](https://learn.microsoft.com/en-us/office/vba/api/excel.window.activesheet).
        */
        wxExcelSheet GetActiveSheet();

        /**
        True if the auto filter for date grouping is currently displayed in the specified window. Since Excel 2007.

        [MSDN documentation for Window.AutoFilterDateGrouping](https://learn.microsoft.com/en-us/office/vba/api/excel.window.autofilterdategrouping).
        */
        bool GetAutoFilterDateGrouping();

        /**
        True if the auto filter for date grouping is currently displayed in the specified window.

        [MSDN documentation for Window.AutoFilterDateGrouping](https://learn.microsoft.com/en-us/office/vba/api/excel.window.autofilterdategrouping).
        */
        void SetAutoFilterDateGrouping(bool autoFilterDateGrouping);

        /**
        Returns the name that appears in the title bar of the document window.

        [MSDN documentation for Window.Caption](https://learn.microsoft.com/en-us/office/vba/api/excel.window.caption).
        */
        wxString GetCaption();

        /**
        Sets the name that appears in the title bar of the document window.

        [MSDN documentation for Window.Caption](https://learn.microsoft.com/en-us/office/vba/api/excel.window.caption).
        */
        void SetCaption(const wxString& caption);

        /**
        True if the window is displaying formulas; False if the window is displaying values.

        [MSDN documentation for Window.DisplayFormulas](https://learn.microsoft.com/en-us/office/vba/api/excel.window.displayformulas).
        */
        bool GetDisplayFormulas();

        /**
        True if the window is displaying formulas; False if the window is displaying values.

        [MSDN documentation for Window.DisplayFormulas](https://learn.microsoft.com/en-us/office/vba/api/excel.window.displayformulas).
        */
        void SetDisplayFormulas(bool displayFormulas);

        /**
        True if gridlines are displayed.

        [MSDN documentation for Window.DisplayGridlines](https://learn.microsoft.com/en-us/office/vba/api/excel.window.displaygridlines).
        */
        bool GetDisplayGridlines();

        /**
        True if gridlines are displayed.

        [MSDN documentation for Window.DisplayGridlines](https://learn.microsoft.com/en-us/office/vba/api/excel.window.displaygridlines).
        */
        void SetDisplayGridlines(bool displayGridlines);

        /**
        True if both row and column headings are displayed; False if no headings are displayed.

        [MSDN documentation for Window.DisplayHeadings](https://learn.microsoft.com/en-us/office/vba/api/excel.window.displayheadings).
        */
        bool GetDisplayHeadings();

        /**
        True if both row and column headings are displayed; False if no headings are displayed.

        [MSDN documentation for Window.DisplayHeadings](https://learn.microsoft.com/en-us/office/vba/api/excel.window.displayheadings).
        */
        void SetDisplayHeadings(bool displayHeadings);

        /**
        True if the horizontal scroll bar is displayed.

        [MSDN documentation for Window.DisplayHorizontalScrollBar](https://learn.microsoft.com/en-us/office/vba/api/excel.window.displayhorizontalscrollbar).
        */
        bool GetDisplayHorizontalScrollBar();

        /**
        True if the horizontal scroll bar is displayed.

        [MSDN documentation for Window.DisplayHorizontalScrollBar](https://learn.microsoft.com/en-us/office/vba/api/excel.window.displayhorizontalscrollbar).
        */
        void SetDisplayHorizontalScrollBar(bool displayHorizontalScrollBar);

        /**
        True if outline symbols are displayed.

        [MSDN documentation for Window.DisplayOutline](https://learn.microsoft.com/en-us/office/vba/api/excel.window.displayoutline).
        */
        bool GetDisplayOutline();

        /**
        True if outline symbols are displayed.

        [MSDN documentation for Window.DisplayOutline](https://learn.microsoft.com/en-us/office/vba/api/excel.window.displayoutline).
        */
        void SetDisplayOutline(bool displayOutline);

        /**
        True if the specified window is displayed from right to left instead of from left to right. False if the object is displayed from left to right.

        [MSDN documentation for Window.DisplayRightToLeft](https://learn.microsoft.com/en-us/office/vba/api/excel.window.displayrighttoleft).
        */
        bool GetDisplayRightToLeft();

        /**
        True if the specified window is displayed from right to left instead of from left to right. False if the object is displayed from left to right.

        [MSDN documentation for Window.DisplayRightToLeft](https://learn.microsoft.com/en-us/office/vba/api/excel.window.displayrighttoleft).
        */
        void SetDisplayRightToLeft(bool displayRightToLeft);

        /**
        True if a ruler is displayed for the specified window. Since Excel 2007.

        [MSDN documentation for Window.DisplayRuler](https://learn.microsoft.com/en-us/office/vba/api/excel.window.displayruler).
        */
        bool GetDisplayRuler();

        /**
        True if a ruler is displayed for the specified window.

        [MSDN documentation for Window.DisplayRuler](https://learn.microsoft.com/en-us/office/vba/api/excel.window.displayruler).
        */
        void SetDisplayRuler(bool displayRuler);

        /**
        True if the vertical scroll bar is displayed.

        [MSDN documentation for Window.DisplayVerticalScrollBar](https://learn.microsoft.com/en-us/office/vba/api/excel.window.displayverticalscrollbar).
        */
        bool GetDisplayVerticalScrollBar();

        /**
        True if the vertical scroll bar is displayed.

        [MSDN documentation for Window.DisplayVerticalScrollBar](https://learn.microsoft.com/en-us/office/vba/api/excel.window.displayverticalscrollbar).
        */
        void SetDisplayVerticalScrollBar(bool displayVerticalScrollBar);

        /**
        True if whitespace is displayed.  Since Excel 2007.

        [MSDN documentation for Window.DisplayWhitespace](https://learn.microsoft.com/en-us/office/vba/api/excel.window.displaywhitespace).
        */
        bool GetDisplayWhitespace();

        /**
        True if whitespace is displayed.

        [MSDN documentation for Window.DisplayWhitespace](https://learn.microsoft.com/en-us/office/vba/api/excel.window.displaywhitespace).
        */
        void SetDisplayWhitespace(bool displayWhitespace);

        /**
        True if the workbook tabs are displayed.

        [MSDN documentation for Window.DisplayWorkbookTabs](https://learn.microsoft.com/en-us/office/vba/api/excel.window.displayworkbooktabs).
        */
        bool GetDisplayWorkbookTabs();

        /**
        True if the workbook tabs are displayed.

        [MSDN documentation for Window.DisplayWorkbookTabs](https://learn.microsoft.com/en-us/office/vba/api/excel.window.displayworkbooktabs).
        */
        void SetDisplayWorkbookTabs(bool displayWorkbookTabs);

        /**
        True if zero values are displayed.

        [MSDN documentation for Window.DisplayZeros](https://learn.microsoft.com/en-us/office/vba/api/excel.window.displayzeros).
        */
        bool GetDisplayZeros();

        /**
        True if zero values are displayed.

        [MSDN documentation for Window.DisplayZeros](https://learn.microsoft.com/en-us/office/vba/api/excel.window.displayzeros).
        */
        void SetDisplayZeros(bool displayZeros);

        /**
        True if the window can be resized.

        [MSDN documentation for Window.EnableResize](https://learn.microsoft.com/en-us/office/vba/api/excel.window.enableresize).
        */
        bool GetEnableResize();

        /**
        True if the window can be resized.

        [MSDN documentation for Window.EnableResize](https://learn.microsoft.com/en-us/office/vba/api/excel.window.enableresize).
        */
        void SetEnableResize(bool enableResize);

        /**
        True if split panes are frozen.

        [MSDN documentation for Window.FreezePanes](https://learn.microsoft.com/en-us/office/vba/api/excel.window.freezepanes).
        */
        bool GetFreezePanes();

        /**
        True if split panes are frozen.

        [MSDN documentation for Window.FreezePanes](https://learn.microsoft.com/en-us/office/vba/api/excel.window.freezepanes).
        */
        void SetFreezePanes(bool freezePanes);

        /**
        Returns the gridline color as an RGB value.

        [MSDN documentation for Window.GridlineColor](https://learn.microsoft.com/en-us/office/vba/api/excel.window.gridlinecolor).
        */
        wxColour GetGridlineColor();

        /**
        Sets the gridline color as an RGB value.

        [MSDN documentation for Window.GridlineColor](https://learn.microsoft.com/en-us/office/vba/api/excel.window.gridlinecolor).
        */
        void SetGridlineColor(const wxColour& gridlineColor);

        /**
        Returns the gridline color as an index into the current color palette or as the following XlColorIndex constant.

        [MSDN documentation for Window.GridlineColorIndex](https://learn.microsoft.com/en-us/office/vba/api/excel.window.gridlinecolorindex).
        */
        long GetGridlineColorIndex();

        /**
        Sets the gridline color as an index into the current color palette or as the following XlColorIndex constant.

        [MSDN documentation for Window.GridlineColorIndex](https://learn.microsoft.com/en-us/office/vba/api/excel.window.gridlinecolorindex).
        */
        void SetGridlineColorIndex(long gridlineColorIndex);

        /**
        Returns a Double value that represents tThe height, in points, of the window.

        [MSDN documentation for Window.Height](https://learn.microsoft.com/en-us/office/vba/api/excel.window.height).
        */
        double GetHeight();

        /**
        Sets a Double value that represents tThe height, in points, of the window.

        [MSDN documentation for Window.Height](https://learn.microsoft.com/en-us/office/vba/api/excel.window.height).
        */
        void SetHeight(double height);

        /**
        Returns a Long value that represents the index number of the object within the collection of similar objects.

        [MSDN documentation for Window.Index](https://learn.microsoft.com/en-us/office/vba/api/excel.window.index).
        */
        long GetIndex();

        /**
        Returns a Double value that represents the distance, in points, from the left edge of the client area to the left edge of the window.

        [MSDN documentation for Window.Left](https://learn.microsoft.com/en-us/office/vba/api/excel.window.left).
        */
        double GetLeft();

        /**
        Sets a Double value that represents the distance, in points, from the left edge of the client area to the left edge of the window.

        [MSDN documentation for Window.Left](https://learn.microsoft.com/en-us/office/vba/api/excel.window.left).
        */
        void SetLeft(double left);

        /**
        Returns a Panes collection that represents all the panes in the specified window.

        [MSDN documentation for Window.Panes](https://learn.microsoft.com/en-us/office/vba/api/excel.window.panes).
        */
        wxExcelPanes GetPanes();

        /**
        Returns a Range Represents the selected cells on the worksheet in the specified window even if a graphic object is active or selected on the worksheet.

        [MSDN documentation for Window.RangeSelection](https://learn.microsoft.com/en-us/office/vba/api/excel.window.rangeselection).
        */
        wxExcelRange GetRangeSelection();


        /**
        Returns the number of the leftmost column in the pane or window.

        [MSDN documentation for Window.ScrollColumn](https://learn.microsoft.com/en-us/office/vba/api/excel.window.scrollcolumn).
        */
        long GetScrollColumn();

        /**
        Sets the number of the leftmost column in the pane or window.

        [MSDN documentation for Window.ScrollColumn](https://learn.microsoft.com/en-us/office/vba/api/excel.window.scrollcolumn).
        */
        void SetScrollColumn(long scrollColumn);

        /**
        Returns the number of the row that appears at the top of the pane or window.

        [MSDN documentation for Window.ScrollRow](https://learn.microsoft.com/en-us/office/vba/api/excel.window.scrollrow).
        */
        long GetScrollRow();

        /**
        Sets the number of the row that appears at the top of the pane or window.

        [MSDN documentation for Window.ScrollRow](https://learn.microsoft.com/en-us/office/vba/api/excel.window.scrollrow).
        */
        void SetScrollRow(long scrollRow);

        /**
        Returns a Sheets collection that represents all the selected sheets in the specified window.

        [MSDN documentation for Window.SelectedSheets](https://learn.microsoft.com/en-us/office/vba/api/excel.window.selectedsheets).
        */
        wxExcelSheets GetSelectedSheets();

        /**
        Returns the SheetViews object for the specified window. Since Excel 2007.

        [MSDN documentation for Window.SheetViews](https://learn.microsoft.com/en-us/office/vba/api/excel.window.sheetviews).
        */
        wxExcelSheetViews GetSheetViews();


        /**
        True if the window is split.

        [MSDN documentation for Window.Split](https://learn.microsoft.com/en-us/office/vba/api/excel.window.split).
        */
        bool GetSplit();

        /**
        True if the window is split.

        [MSDN documentation for Window.Split](https://learn.microsoft.com/en-us/office/vba/api/excel.window.split).
        */
        void SetSplit(bool split);

        /**
        Returns the column number where the window is split into panes (the number of columns to the left of the split line).

        [MSDN documentation for Window.SplitColumn](https://learn.microsoft.com/en-us/office/vba/api/excel.window.splitcolumn).
        */
        long GetSplitColumn();

        /**
        Sets the column number where the window is split into panes (the number of columns to the left of the split line).

        [MSDN documentation for Window.SplitColumn](https://learn.microsoft.com/en-us/office/vba/api/excel.window.splitcolumn).
        */
        void SetSplitColumn(long splitColumn);

        /**
        Returns the location of the horizontal window split, in points. Read/write Double.

        [MSDN documentation for Window.SplitHorizontal](https://learn.microsoft.com/en-us/office/vba/api/excel.window.splithorizontal).
        */
        double GetSplitHorizontal();

        /**
        Sets the location of the horizontal window split, in points. Read/write Double.

        [MSDN documentation for Window.SplitHorizontal](https://learn.microsoft.com/en-us/office/vba/api/excel.window.splithorizontal).
        */
        void SetSplitHorizontal(double splitHorizontal);

        /**
        Returns the row number where the window is split into panes (the number of rows above the split).

        [MSDN documentation for Window.SplitRow](https://learn.microsoft.com/en-us/office/vba/api/excel.window.splitrow).
        */
        long GetSplitRow();

        /**
        Sets the row number where the window is split into panes (the number of rows above the split).

        [MSDN documentation for Window.SplitRow](https://learn.microsoft.com/en-us/office/vba/api/excel.window.splitrow).
        */
        void SetSplitRow(long splitRow);

        /**
        Returns the location of the vertical window split, in points. Read/write Double.

        [MSDN documentation for Window.SplitVertical](https://learn.microsoft.com/en-us/office/vba/api/excel.window.splitvertical).
        */
        double GetSplitVertical();

        /**
        Sets the location of the vertical window split, in points. Read/write Double.

        [MSDN documentation for Window.SplitVertical](https://learn.microsoft.com/en-us/office/vba/api/excel.window.splitvertical).
        */
        void SetSplitVertical(double splitVertical);

        /**
        Returns the ratio of the width of the workbook's tab area to the width of the window's horizontal scroll bar (as a number between 0 (zero) and 1; the default value is 0.6). Read/write Double.

        [MSDN documentation for Window.TabRatio](https://learn.microsoft.com/en-us/office/vba/api/excel.window.tabratio).
        */
        double GetTabRatio();

        /**
        Sets the ratio of the width of the workbook's tab area to the width of the window's horizontal scroll bar (as a number between 0 (zero) and 1; the default value is 0.6). Read/write Double.

        [MSDN documentation for Window.TabRatio](https://learn.microsoft.com/en-us/office/vba/api/excel.window.tabratio).
        */
        void SetTabRatio(double tabRatio);

        /**
        Returns a Double value that represents the distance, in points, from the top edge of the window to the top edge of the usable area (below the menus, any toolbars docked at the top, and the formula bar).

        [MSDN documentation for Window.Top](https://learn.microsoft.com/en-us/office/vba/api/excel.window.top).
        */
        double GetTop();

        /**
        Sets a Double value that represents the distance, in points, from the top edge of the window to the top edge of the usable area (below the menus, any toolbars docked at the top, and the formula bar).

        [MSDN documentation for Window.Top](https://learn.microsoft.com/en-us/office/vba/api/excel.window.top).
        */
        void SetTop(double top);

        /**
        Returns a XlWindowType value that represents the window type.

        [MSDN documentation for Window.Type](https://learn.microsoft.com/en-us/office/vba/api/excel.window.type).
        */
        XlWindowType GetType();

        /**
        Returns the maximum height of the space that a window can occupy in the application window area, in points. Read-only Double.

        [MSDN documentation for Window.UsableHeight](https://learn.microsoft.com/en-us/office/vba/api/excel.window.usableheight).
        */
        double GetUsableHeight();

        /**
        Returns the maximum width of the space that a window can occupy in the application window area, in points. Read-only Double.

        [MSDN documentation for Window.UsableWidth](https://learn.microsoft.com/en-us/office/vba/api/excel.window.usablewidth).
        */
        double GetUsableWidth();

        /**
        Returns the view showing in the window. Read/write XlWindowView.

        [MSDN documentation for Window.View](https://learn.microsoft.com/en-us/office/vba/api/excel.window.view).
        */
        XlWindowView GetView();

        /**
        Sets the view showing in the window. Read/write XlWindowView.

        [MSDN documentation for Window.View](https://learn.microsoft.com/en-us/office/vba/api/excel.window.view).
        */
        void SetView(XlWindowView view);

        /**
        Returns a Boolean value that determines whether the object is visible.

        [MSDN documentation for Window.Visible](https://learn.microsoft.com/en-us/office/vba/api/excel.window.visible).
        */
        bool GetVisible();

        /**
        Sets a Boolean value that determines whether the object is visible.

        [MSDN documentation for Window.Visible](https://learn.microsoft.com/en-us/office/vba/api/excel.window.visible).
        */
        void SetVisible(bool visible);

        /**
        Returns a Range Represents the range of cells that are visible in the window or pane. If a column or row is partially visible, it's included in the range.

        [MSDN documentation for Window.VisibleRange](https://learn.microsoft.com/en-us/office/vba/api/excel.window.visiblerange).
        */
        wxExcelRange GetVisibleRange();

        /**
        Returns a Double value that represents the width, in points, of the window.

        [MSDN documentation for Window.Width](https://learn.microsoft.com/en-us/office/vba/api/excel.window.width).
        */
        double GetWidth();

        /**
        Sets a Double value that represents the width, in points, of the window.

        [MSDN documentation for Window.Width](https://learn.microsoft.com/en-us/office/vba/api/excel.window.width).
        */
        void SetWidth(double width);

        /**
        Returns the window number. For example, a window named "Book1.xls:2" has 2 as its window number. Most windows have the window number 1.

        [MSDN documentation for Window.WindowNumber](https://learn.microsoft.com/en-us/office/vba/api/excel.window.windownumber).
        */
        long GetWindowNumber();

        /**
        Returns the state of the window. Read/write XlWindowState.

        [MSDN documentation for Window.WindowState](https://learn.microsoft.com/en-us/office/vba/api/excel.window.windowstate).
        */
        XlWindowState GetWindowState();

        /**
        Sets the state of the window. Read/write XlWindowState.

        [MSDN documentation for Window.WindowState](https://learn.microsoft.com/en-us/office/vba/api/excel.window.windowstate).
        */
        void SetWindowState(XlWindowState windowState);

        /**
        Returns the display size of the window, as a percentage (100 equals normal size, 200 equals double size, and so on).

        [MSDN documentation for Window.Zoom](https://learn.microsoft.com/en-us/office/vba/api/excel.window.zoom).
        */
        long GetZoom();

        /**
        Sets the display size of the window, as a percentage (100 equals normal size, 200 equals double size, and so on).

        [MSDN documentation for Window.Zoom](https://learn.microsoft.com/en-us/office/vba/api/excel.window.zoom).
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

        [MSDN documentation for Windows.Arrange](https://learn.microsoft.com/en-us/office/vba/api/excel.windows.arrange).
        */
        bool Arrange(XlArrangeStyle* arrangeStyle = NULL, wxXlTribool activeWorkbook =  wxDefaultXlTribool,
            wxXlTribool syncHorizontal = wxDefaultXlTribool, wxXlTribool syncVertica = wxDefaultXlTribool);

        /**
        Ends side-by-side mode if two windows are in side-by-side mode. Returns a Boolean value that represents whether the method was successful.

        [MSDN documentation for Windows.BreakSideBySide](https://learn.microsoft.com/en-us/office/vba/api/excel.windows.breaksidebyside).
        */
        bool BreakSideBySide();

        /**
        Opens two windows in side-by-side mode. Returns a Boolean value.

        [MSDN documentation for Windows.CompareSideBySideWith](https://learn.microsoft.com/en-us/office/vba/api/excel.windows.comparesidebysidewith).
        */
        bool CompareSideBySideWith(const wxString& windowName);

        /**
        Resets the position of two worksheet windows that are being compared side by side.

        [MSDN documentation for Windows.ResetPositionsSideBySide](https://learn.microsoft.com/en-us/office/vba/api/excel.windows.resetpositionssidebyside).
        */
        void ResetPositionsSideBySide();


        // ***** PROPERTIES *****

        /**
        Returns a Long value that represents the number of objects in the collection.

        [MSDN documentation for Windows.Count](https://learn.microsoft.com/en-us/office/vba/api/excel.windows.count).
        */
        long GetCount();

        //@{
        /**
        Returns a single object from a collection.

        [MSDN documentation for Windows.Item](https://learn.microsoft.com/en-us/office/vba/api/excel.windows.item).
        */
        wxExcelWindow GetItem(long index);
        wxExcelWindow operator[](long index);
        wxExcelWindow GetItem(const wxString& name);
        wxExcelWindow operator[](const wxString& name);
        //@}

        /**
        True enables scrolling the contents of windows at the same time when documents are being compared side by side. False disables scrolling the windows at the same time.

        [MSDN documentation for Windows.SyncScrollingSideBySide](https://learn.microsoft.com/en-us/office/vba/api/excel.windows.syncscrollingsidebyside).
        */
        bool GetSyncScrollingSideBySide();

        /**
        True enables scrolling the contents of windows at the same time when documents are being compared side by side. False disables scrolling the windows at the same time.

        [MSDN documentation for Windows.SyncScrollingSideBySide](https://learn.microsoft.com/en-us/office/vba/api/excel.windows.syncscrollingsidebyside).
        */
        void SetSyncScrollingSideBySide(bool syncScrollingSideBySide);


        /**
        Returns "Windows".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("Windows"); }
    };


} // namespace wxAutoExcel

#endif //_WXAUTOEXCEL_WINDOWS_H
