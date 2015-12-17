/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Modified by:
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////


#include "wx/wxAutoExcel_prec.h"

#include "wx/wxAutoExcelWindows.h"

#include "wx/wxAutoExcelPanes.h"
#include "wx/wxAutoExcelRange.h"
#include "wx/wxAutoExcelShape.h"
#include "wx/wxAutoExcelSheet.h"
#include "wx/wxAutoExcelSheets.h"

#include "wx/wxAutoExcel_private.h"

namespace wxAutoExcel {

// ***** class wxExcelWindow METHODS *****

bool wxExcelWindow::Activate()
{
    WXAUTOEXCEL_CALL_METHOD0_BOOL("Activate");
}

bool wxExcelWindow::ActivateNext()
{
    WXAUTOEXCEL_CALL_METHOD0_BOOL("ActivateNext");
}

bool wxExcelWindow::ActivatePrevious()
{
    WXAUTOEXCEL_CALL_METHOD0_BOOL("ActivatePrevious");
}

bool wxExcelWindow::Close()
{
    WXAUTOEXCEL_CALL_METHOD0_BOOL("Close");
}

bool wxExcelWindow::LargeScroll(long* down, long* up, long* toRight, long* toLeft)
{
    WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT_NAME(Down, down);
    WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT_NAME(Up, up);
    WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT_NAME(ToRight, toRight);
    WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT_NAME(ToLeft, toLeft);

    WXAUTOEXCEL_CALL_METHOD4("LargeScroll", vDown, vUp, vToRight, vToLeft, "bool", false);
    return vResult.GetBool();
}

wxExcelWindow wxExcelWindow::NewWindow()
{
    wxExcelWindow window;

    WXAUTOEXCEL_CALL_METHOD0_OBJECT("NewWindow", window);    
}

long wxExcelWindow::PointsToScreenPixelsX(double points)
{
    WXAUTOEXCEL_CALL_METHOD1_LONG("PointsToScreenPixelsX", points, 0);
}

long wxExcelWindow::PointsToScreenPixelsY(double points)
{
   WXAUTOEXCEL_CALL_METHOD1_LONG("PointsToScreenPixelsY", points, 0);
}

bool wxExcelWindow::PrintOut(long* from, long* to, long* copies, wxXlTribool preview, const wxString& activePrinter,
                                 wxXlTribool printToFile, wxXlTribool collate, const wxString& prToFileName)
{
    wxVariantVector args;

    WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT_NAME_VECTOR(From, from, args);
    WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT_NAME_VECTOR(To, to, args);
    WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT_NAME_VECTOR(Copies, copies, args);
    WXAUTOEXCEL_OPTIONALCPPTBOOL_TO_OPTIONALVARIANT_NAME_VECTOR(Preview, preview, args);
    WXAUTOEXCEL_OPTIONALCPPSTR_TO_OPTIONALVARIANT_NAME_VECTOR(ActivePrinter, activePrinter, args);
    WXAUTOEXCEL_OPTIONALCPPTBOOL_TO_OPTIONALVARIANT_NAME_VECTOR(PrintToFile, printToFile, args);
    WXAUTOEXCEL_OPTIONALCPPTBOOL_TO_OPTIONALVARIANT_NAME_VECTOR(Collate, collate, args);
    WXAUTOEXCEL_OPTIONALCPPSTR_TO_OPTIONALVARIANT_NAME_VECTOR(PrToFileName, prToFileName, args);

    return PrintOut(args);
}

bool wxExcelWindow::PrintOut(const wxVariantVector& args)
{
    WXAUTOEXCEL_CALL_METHODARR("PrintOut", args, "bool", false);
    return vResult.GetBool();
}

bool wxExcelWindow::PrintPreview(wxXlTribool enableChanges)
{
    WXAUTOEXCEL_OPTIONALCPPTBOOL_TO_OPTIONALVARIANT_NAME(EnableChanges, enableChanges);

    WXAUTOEXCEL_CALL_METHOD1_BOOL("PrintPreview", vEnableChanges);
}

#if WXAUTOEXCEL_USE_SHAPES

bool wxExcelWindow::RangeFromPoint(long x, long y, wxExcelRange& range, wxExcelShape& shape)
{
    wxExcelObject object;
    wxVariant vResult;

    if ( InvokeMethod(wxS("RangeFromPoint"), vResult, x, y) )
    {
        if ( !vResult.IsNull() )
        {         
            VariantToObject(vResult, &object);

            if ( object.GetAutomationObjectName_(true) == wxS("Range") )            
                CloneDispatch(&object, &range);            
            else
                CloneDispatch(&object, &shape);            
            
            return true;
        }
    }
    return false;
}

#endif // #if WXAUTOEXCEL_USE_SHAPES

void wxExcelWindow::ScrollIntoView(long left, long top, long width, long height, wxXlTribool start)
{
    WXAUTOEXCEL_OPTIONALCPPTBOOL_TO_OPTIONALVARIANT(Start, start);
    WXAUTOEXCEL_CALL_METHOD5_RET("ScrollIntoView", left, top, width, height, vStart, "null");
}


bool wxExcelWindow::ScrollWorkbookTabs(long* sheets, long* position)
{
    WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT_NAME(Sheets, sheets);
    WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT_NAME(Position, position);

    WXAUTOEXCEL_CALL_METHOD2("ScrollWorkbookTabs", vSheets, vPosition, "bool", false);
    return vResult.GetBool();
}

bool wxExcelWindow::SmallScroll(long* down, long* up, long* toRight, long* toLeft)
{
    WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT_NAME(Down, down);
    WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT_NAME(Up, up);
    WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT_NAME(ToRight, toRight);
    WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT_NAME(ToLeft, toLeft);

    WXAUTOEXCEL_CALL_METHOD4("SmallScroll", vDown, vUp, vToRight, vToLeft, "bool", false);
    return vResult.GetBool();
}

// ***** class wxExcelWindow PROPERTIES *****

wxExcelRange wxExcelWindow::GetActiveCell()
{
    wxExcelRange range;
    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("ActiveCell", range);
}


wxExcelPane wxExcelWindow::GetActivePane()
{
    wxExcelPane pane;
    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("ActivePane", pane);
}

wxExcelSheet wxExcelWindow::GetActiveSheet()
{
    wxExcelSheet sheet;
    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("ActiveSheet", sheet);
}


bool wxExcelWindow::GetAutoFilterDateGrouping()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("AutoFilterDateGrouping");
}

void wxExcelWindow::SetAutoFilterDateGrouping(bool autoFilterDateGrouping)
{
    InvokePutProperty(wxS("AutoFilterDateGrouping"), autoFilterDateGrouping);
}

wxString wxExcelWindow::GetCaption()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("Caption");
}

void wxExcelWindow::SetCaption(const wxString& caption)
{
    InvokePutProperty(wxS("Caption"), caption);
}

bool wxExcelWindow::GetDisplayFormulas()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("DisplayFormulas");
}

void wxExcelWindow::SetDisplayFormulas(bool displayFormulas)
{
    InvokePutProperty(wxS("DisplayFormulas"), displayFormulas);
}

bool wxExcelWindow::GetDisplayGridlines()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("DisplayGridlines");
}

void wxExcelWindow::SetDisplayGridlines(bool displayGridlines)
{
    InvokePutProperty(wxS("DisplayGridlines"), displayGridlines);
}

bool wxExcelWindow::GetDisplayHeadings()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("DisplayHeadings");
}

void wxExcelWindow::SetDisplayHeadings(bool displayHeadings)
{
    InvokePutProperty(wxS("DisplayHeadings"), displayHeadings);
}

bool wxExcelWindow::GetDisplayHorizontalScrollBar()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("DisplayHorizontalScrollBar");
}

void wxExcelWindow::SetDisplayHorizontalScrollBar(bool displayHorizontalScrollBar)
{
    InvokePutProperty(wxS("DisplayHorizontalScrollBar"), displayHorizontalScrollBar);
}

bool wxExcelWindow::GetDisplayOutline()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("DisplayOutline");
}

void wxExcelWindow::SetDisplayOutline(bool displayOutline)
{
    InvokePutProperty(wxS("DisplayOutline"), displayOutline);
}

bool wxExcelWindow::GetDisplayRightToLeft()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("DisplayRightToLeft");
}

void wxExcelWindow::SetDisplayRightToLeft(bool displayRightToLeft)
{
    InvokePutProperty(wxS("DisplayRightToLeft"), displayRightToLeft);
}

bool wxExcelWindow::GetDisplayRuler()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("DisplayRuler");
}

void wxExcelWindow::SetDisplayRuler(bool displayRuler)
{
    InvokePutProperty(wxS("DisplayRuler"), displayRuler);
}

bool wxExcelWindow::GetDisplayVerticalScrollBar()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("DisplayVerticalScrollBar");
}

void wxExcelWindow::SetDisplayVerticalScrollBar(bool displayVerticalScrollBar)
{
    InvokePutProperty(wxS("DisplayVerticalScrollBar"), displayVerticalScrollBar);
}

bool wxExcelWindow::GetDisplayWhitespace()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("DisplayWhitespace");
}

void wxExcelWindow::SetDisplayWhitespace(bool displayWhitespace)
{
    InvokePutProperty(wxS("DisplayWhitespace"), displayWhitespace);
}

bool wxExcelWindow::GetDisplayWorkbookTabs()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("DisplayWorkbookTabs");
}

void wxExcelWindow::SetDisplayWorkbookTabs(bool displayWorkbookTabs)
{
    InvokePutProperty(wxS("DisplayWorkbookTabs"), displayWorkbookTabs);
}

bool wxExcelWindow::GetDisplayZeros()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("DisplayZeros");
}

void wxExcelWindow::SetDisplayZeros(bool displayZeros)
{
    InvokePutProperty(wxS("DisplayZeros"), displayZeros);
}

bool wxExcelWindow::GetEnableResize()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("EnableResize");
}

void wxExcelWindow::SetEnableResize(bool enableResize)
{
    InvokePutProperty(wxS("EnableResize"), enableResize);
}

bool wxExcelWindow::GetFreezePanes()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("FreezePanes");
}

void wxExcelWindow::SetFreezePanes(bool freezePanes)
{
    InvokePutProperty(wxS("FreezePanes"), freezePanes);
}

wxColour wxExcelWindow::GetGridlineColor()
{
    WXAUTOEXCEL_PROPERTY_COLOR_GET0("GridlineColor");
}

void wxExcelWindow::SetGridlineColor(const wxColour& gridlineColor)
{    
    InvokePutProperty(wxS("GridlineColor"), (long)gridlineColor.GetRGB());
}

long wxExcelWindow::GetGridlineColorIndex()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("GridlineColorIndex");
}

void wxExcelWindow::SetGridlineColorIndex(long gridlineColorIndex)
{
    InvokePutProperty(wxS("GridlineColorIndex"), gridlineColorIndex);
}

double wxExcelWindow::GetHeight()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("Height");
}

void wxExcelWindow::SetHeight(double height)
{
    InvokePutProperty(wxS("Height"), height);
}

long wxExcelWindow::GetIndex()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("Index");
}

double wxExcelWindow::GetLeft()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("Left");
}

void wxExcelWindow::SetLeft(double left)
{
    InvokePutProperty(wxS("Left"), left);
}

wxExcelPanes wxExcelWindow::GetPanes()
{
    wxExcelPanes panes;
    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("Panes", panes);
}


wxExcelRange wxExcelWindow::GetRangeSelection()
{
    wxExcelRange range;
    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("RangeSelection", range);    
}

long wxExcelWindow::GetScrollColumn()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("ScrollColumn");
}

void wxExcelWindow::SetScrollColumn(long scrollColumn)
{
    InvokePutProperty(wxS("ScrollColumn"), scrollColumn);
}

long wxExcelWindow::GetScrollRow()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("ScrollRow");
}

void wxExcelWindow::SetScrollRow(long scrollRow)
{
    InvokePutProperty(wxS("ScrollRow"), scrollRow);
}

wxExcelSheets wxExcelWindow::GetSelectedSheets()
{
    wxExcelSheets sheets;
    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("SelectedSheets", sheets);
}

wxExcelSheetViews wxExcelWindow::GetSheetViews()
{
    wxExcelSheetViews sheetViews;
    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("SheetViews", sheetViews);
}

bool wxExcelWindow::GetSplit()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("Split");
}

void wxExcelWindow::SetSplit(bool split)
{
    InvokePutProperty(wxS("Split"), split);
}

long wxExcelWindow::GetSplitColumn()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("SplitColumn");
}

void wxExcelWindow::SetSplitColumn(long splitColumn)
{
    InvokePutProperty(wxS("SplitColumn"), splitColumn);
}

double wxExcelWindow::GetSplitHorizontal()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("SplitHorizontal");
}

void wxExcelWindow::SetSplitHorizontal(double splitHorizontal)
{
    InvokePutProperty(wxS("SplitHorizontal"), splitHorizontal);
}

long wxExcelWindow::GetSplitRow()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("SplitRow");
}

void wxExcelWindow::SetSplitRow(long splitRow)
{
    InvokePutProperty(wxS("SplitRow"), splitRow);
}

double wxExcelWindow::GetSplitVertical()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("SplitVertical");
}

void wxExcelWindow::SetSplitVertical(double splitVertical)
{
    InvokePutProperty(wxS("SplitVertical"), splitVertical);
}

double wxExcelWindow::GetTabRatio()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("TabRatio");
}

void wxExcelWindow::SetTabRatio(double tabRatio)
{
    InvokePutProperty(wxS("TabRatio"), tabRatio);
}

double wxExcelWindow::GetTop()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("Top");
}

void wxExcelWindow::SetTop(double top)
{
    InvokePutProperty(wxS("Top"), top);
}

XlWindowType wxExcelWindow::GetType()
{    
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("Type", XlWindowType, xlChartAsWindow);
}

void wxExcelWindow::SetType(XlWindowType type)
{
    InvokePutProperty("Type", long(type));
}

double wxExcelWindow::GetUsableHeight()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("UsableHeight");
}

double wxExcelWindow::GetUsableWidth()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("UsableWidth");
}

XlWindowView wxExcelWindow::GetView()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("View", XlWindowView, xlNormalView);    
}

void wxExcelWindow::SetView(XlWindowView view)
{
    InvokePutProperty("View", long(view));
}

bool wxExcelWindow::GetVisible()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("Visible");
}

void wxExcelWindow::SetVisible(bool visible)
{
    InvokePutProperty(wxS("Visible"), visible);
}

wxExcelRange wxExcelWindow::GetVisibleRange()
{
    wxExcelRange range;
    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("VisibleRange", range);    
}

double wxExcelWindow::GetWidth()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("Width");
}

void wxExcelWindow::SetWidth(double width)
{
    InvokePutProperty(wxS("Width"), width);
}

long wxExcelWindow::GetWindowNumber()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("WindowNumber");
}

XlWindowState wxExcelWindow::GetWindowState()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("WindowState", XlWindowState, (XlWindowState)0);
}

void wxExcelWindow::SetWindowState(XlWindowState windowState)
{
    InvokePutProperty(wxS("WindowState"), (long)windowState);
}

long wxExcelWindow::GetZoom()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("Zoom");
}

void wxExcelWindow::SetZoom(long zoom)
{
    InvokePutProperty(wxS("Zoom"), zoom);
}



// ***** class wxExcelWindows METHODS *****

bool wxExcelWindows::Arrange(XlArrangeStyle* arrangeStyle, wxXlTribool activeWorkbook,
                             wxXlTribool syncHorizontal, wxXlTribool syncVertical)
{
    WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT_NAME(ArrangeStyle, ((long*)arrangeStyle));
    WXAUTOEXCEL_OPTIONALCPPTBOOL_TO_OPTIONALVARIANT_NAME(ActiveWorkbook, activeWorkbook);
    WXAUTOEXCEL_OPTIONALCPPTBOOL_TO_OPTIONALVARIANT_NAME(SyncHorizontal, syncHorizontal);
    WXAUTOEXCEL_OPTIONALCPPTBOOL_TO_OPTIONALVARIANT_NAME(SyncVertical, syncVertical);

    WXAUTOEXCEL_CALL_METHOD4("Arrange", vArrangeStyle, vActiveWorkbook, vSyncHorizontal, vSyncVertical, "bool", false);
    return vResult.GetBool();
}

bool wxExcelWindows::BreakSideBySide()
{
    WXAUTOEXCEL_CALL_METHOD0_BOOL("BreakSideBySide");
}

bool wxExcelWindows::CompareSideBySideWith(const wxString& windowName)
{
    WXAUTOEXCEL_CALL_METHOD1_BOOL("CompareSideBySideWith", windowName);
}

void wxExcelWindows::ResetPositionsSideBySide()
{
    WXAUTOEXCEL_CALL_METHOD0_RET("ResetPositionsSideBySide", "null");
}


// ***** class wxExcelWindows PROPERTIES *****


long wxExcelWindows::GetCount()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("Count");
}

wxExcelWindow wxExcelWindows::GetItem(long index)
{
    wxASSERT( index > 0 );

    wxExcelWindow window;
    WXAUTOEXCEL_PROPERTY_OBJECT_GET1("Item", index, window);    
}

wxExcelWindow wxExcelWindows::operator[](long index)
{
    return GetItem(index);
}

wxExcelWindow wxExcelWindows::GetItem(const wxString& name)
{
    wxExcelWindow window;
    WXAUTOEXCEL_PROPERTY_OBJECT_GET1("Item", name, window);    
}

wxExcelWindow wxExcelWindows::operator[](const wxString& name)
{
    return GetItem(name);
}


bool wxExcelWindows::GetSyncScrollingSideBySide()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("SyncScrollingSideBySide");
}

void wxExcelWindows::SetSyncScrollingSideBySide(bool syncScrollingSideBySide)
{
    InvokePutProperty(wxS("SyncScrollingSideBySide"), syncScrollingSideBySide);
}




} // namespace wxAutoExcel
