/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// License:     MIT license
/////////////////////////////////////////////////////////////////////////////


#ifndef _WXAUTOEXCEL_WORKSHEET_H
#define _WXAUTOEXCEL_WORKSHEET_H

#include "wx/wxAutoExcel_defs.h"

#include "wx/wxAutoExcelListObject.h"
#include "wx/wxAutoExcelRangeOwner.h"
#include "wx/wxAutoExcel_enums.h"

namespace wxAutoExcel {

    /**
    @brief Represents Microsoft Excel Worksheet.
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelWorksheet : public wxExcelRangeOwner
    {
    public:

        // ***** METHODS *****

        /**
        Makes the current sheet the active sheet.

        [MSDN documentation for Worksheet.Activate](https://learn.microsoft.com/en-us/office/vba/api/excel.worksheet.activate%28method%29)).
        */

        bool Activate();

        /**
        Calculates a specific worksheet.

        [MSDN documentation for Worksheet.Calculate](https://learn.microsoft.com/en-us/office/vba/api/excel.worksheet.calculate%28method%29)).
        */
        bool Calculate();

#if WXAUTOEXCEL_USE_CHARTS
        //@{
        /**
        Returns an object that represents either a single embedded chart (a ChartObject object) or a collection of all the embedded charts (a ChartObjects object) on the sheet.

        [MSDN documentation for Worksheet.ChartObjects](https://learn.microsoft.com/en-us/office/vba/api/excel.worksheet.chartobjects).
        */
        wxExcelChartObject ChartObjects(long index);
        wxExcelChartObject ChartObjects(const wxString& name);
        wxExcelChartObjects ChartObjects();
        wxExcelChartObjects ChartObjects(const wxVector<long>& indices);
        wxExcelChartObjects ChartObjects(const wxArrayString& names);
        //@}
#endif // #if WXAUTOEXCEL_USE_CHARTS

        /**
        Checks the spelling of a worksheet.

        [MSDN documentation for Worksheet.CheckSpelling](https://learn.microsoft.com/en-us/office/vba/api/excel.worksheet.checkspelling).
        */
        bool CheckSpelling(const wxString& customDictionary = wxEmptyString, wxXlTribool ignoreUpperCase = wxDefaultXlTribool,
            wxXlTribool alwaysSuggest = wxDefaultXlTribool, MsoLanguageID* spellLang = NULL);

        /**
        Circles invalid entries on the worksheet.

        [MSDN documentation for Worksheet.CircleInvalid](https://learn.microsoft.com/en-us/office/vba/api/excel.worksheet.circleinvalid).
        */
        void CircleInvalid();

        /**
        Clears the tracer arrows from the worksheet. Tracer arrows are added by using the auditing feature.

        [MSDN documentation for Worksheet.ClearArrows](https://learn.microsoft.com/en-us/office/vba/api/excel.worksheet.cleararrows).
        */
        bool ClearArrows();

        /**
        Clears circles from invalid entries on the worksheet.

        [MSDN documentation for Worksheet.ClearCircles](https://learn.microsoft.com/en-us/office/vba/api/excel.worksheet.clearcircles).
        */
        void ClearCircles();

        //@{
        /**
        Method without parameters copies the worksheet into a new workbook. Method with parameters copies the worksheet to another location in the workbook.

        [MSDN documentation for Worksheet.Copy](https://learn.microsoft.com/en-us/office/vba/api/excel.worksheet.copy).
        */
        bool Copy();
        bool CopyAfterOrBefore(wxExcelSheet worksheetAfterOrBefore, bool after);
        //@}

        /**
        Deletes the object.

        [MSDN documentation for Worksheet.Delete](https://learn.microsoft.com/en-us/office/vba/api/excel.worksheet.delete).
        */
        bool Delete();


        //@{
        /**
        Exports to a file of the specified format.

        [MSDN documentation for Worksheet.ExportAsFixedFormat](https://learn.microsoft.com/en-us/office/vba/api/excel.worksheet.exportasfixedformat).
        */
        void ExportAsFixedFormat(XlFixedFormatType type, const wxString& fileName = wxEmptyString,
            XlFixedFormatQuality* quality = NULL, wxXlTribool includeDocProperties = wxDefaultXlTribool,
            wxXlTribool ignorePrintAreas = wxDefaultXlTribool,
            long* from = NULL, long* to = NULL, wxXlTribool openAfterPublish = wxDefaultXlTribool);

        void ExportAsFixedFormat(XlFixedFormatType type, const wxVariantVector& optionalArgs);
        //@}

        //@{
        /**
        Method without parameters moves the worksheet into a new workbook. Method with parameters moves the worksheet to another location in the workbook.

        [MSDN documentation for Worksheet.Move](https://learn.microsoft.com/en-us/office/vba/api/excel.worksheet.move).
        */
        bool Move();
        bool MoveAfterOrBefore(wxExcelSheet sheetAfterOrBefore, bool after);
        //@}

        //@{
        /**
        Returns an either a single OLE object or a collection of all OLE objects (an OLEObjects collection) on the chart or sheet. Read-only.

        [MSDN documentation for Worksheet.OLEObjects](https://learn.microsoft.com/en-us/office/vba/api/excel.worksheet.oleobjects).
        */
        wxExcelOLEObjects OLEObjects();
        wxExcelOLEObject OLEObjects(long index);
        wxExcelOLEObject OLEObjects(const wxString& name);
        //@}

        /**
        Pastes the contents of the Clipboard onto the worksheet.

        [MSDN documentation for Worksheet.Paste](https://learn.microsoft.com/en-us/office/vba/api/excel.worksheet.paste).
        */
        bool Paste(wxExcelRange* destination = NULL, wxXlTribool link = wxDefaultXlTribool);

        /**
        Pastes the contents of the Clipboard onto the worksheet, using a specified format. Use this method to paste data from other applications or to paste data in a specific format.

        [MSDN documentation for Worksheet.PasteSpecial](https://learn.microsoft.com/en-us/office/vba/api/excel.worksheet.pastespecial).
        */
        bool PasteSpecial(const wxString& format = wxEmptyString, wxXlTribool link = wxDefaultXlTribool,
            wxXlTribool displayAsIcon = wxDefaultXlTribool, const wxString& iconFileName = wxEmptyString,
            long* iconIndex = NULL, const wxString& iconLabel = wxEmptyString,
            wxXlTribool noHTMLFormatting = wxDefaultXlTribool);


        //@{
        /**
        Prints the worksheet.

        [MSDN documentation for Worksheet.PrintOut](https://learn.microsoft.com/en-us/office/vba/api/excel.worksheet.printout).
        */
        bool PrintOut(long* from = NULL, long* to = NULL, long* copies = NULL, wxXlTribool preview = wxDefaultXlTribool,
            const wxString& activePrinter = wxEmptyString, wxXlTribool printToFile = wxDefaultXlTribool,
            wxXlTribool collate = wxDefaultXlTribool, const wxString& prToFileName= wxEmptyString, wxXlTribool ignorePrintAreas = wxDefaultXlTribool);
        bool PrintOut(const wxVariantVector& args);
        //@}

        /**
        Shows a preview of the worksheet as it would look when printed.

        [MSDN documentation for Worksheet.PrintPreview](https://learn.microsoft.com/en-us/office/vba/api/excel.worksheet.printpreview).
        */
        bool PrintPreview(wxXlTribool enableChanges = wxDefaultXlTribool);

        //@{
        /**
        Protects a worksheet so that it cannot be modified.

        The overload taking optionalArgs passes them to Excel as they are, so they should be
        named, e.g. wxVariant(true, "AllowSorting").

        [MSDN documentation for Worksheet.Protect](https://learn.microsoft.com/en-us/office/vba/api/excel.worksheet.protect).
        */
        void Protect(const wxString& password = wxEmptyString,
            wxXlTribool drawingObjects = wxDefaultXlTribool, wxXlTribool contents = wxDefaultXlTribool,
            wxXlTribool scenarios = wxDefaultXlTribool, wxXlTribool userInterfaceOnly = wxDefaultXlTribool,
            wxXlTribool allowFormattingCells = wxDefaultXlTribool, wxXlTribool allowFormattingColumns = wxDefaultXlTribool,
            wxXlTribool allowFormattingRows = wxDefaultXlTribool, wxXlTribool allowInsertingColumns = wxDefaultXlTribool,
            wxXlTribool allowInsertingRows = wxDefaultXlTribool, wxXlTribool allowInsertingHyperlinks = wxDefaultXlTribool,
            wxXlTribool allowDeletingColumns = wxDefaultXlTribool, wxXlTribool allowDeletingRows = wxDefaultXlTribool,
            wxXlTribool allowSorting = wxDefaultXlTribool, wxXlTribool allowFiltering = wxDefaultXlTribool,
            wxXlTribool allowUsingPivotTables = wxDefaultXlTribool);

        void Protect(const wxVariantVector& optionalArgs);
        //@}

        /**
        Resets all page breaks on the specified worksheet.

        [MSDN documentation for Worksheet.ResetAllPageBreaks](https://learn.microsoft.com/en-us/office/vba/api/excel.worksheet.resetallpagebreaks).
        */
        void ResetAllPageBreaks();

        //@{
        /**
        Saves changes to the worksheet in a different file.

        [MSDN documentation for Worksheet.SaveAs](https://learn.microsoft.com/en-us/office/vba/api/excel.worksheet.saveas).
        */
        void SaveAs(const wxString& fileName = wxEmptyString, XlFileFormat* fileFormat = NULL,
            const wxString& password = wxEmptyString, const wxString& writeResPassword = wxEmptyString,
            wxXlTribool readOnlyRecommended = wxDefaultXlTribool, wxXlTribool createBackup = wxDefaultXlTribool,
            wxXlTribool addToMru = wxDefaultXlTribool, wxXlTribool local = wxDefaultXlTribool);

        void SaveAs(const wxVariantVector& optionalArgs);
        //@}

        /**
        Selects the worksheet.

        [MSDN documentation for Worksheet.Select](https://learn.microsoft.com/en-us/office/vba/api/excel.worksheet.select).
        */
        bool Select(wxXlTribool replace = wxDefaultXlTribool);

        /**
        Sets the background graphic for a worksheet.

        [MSDN documentation for Worksheet.SetBackgroundPicture](https://learn.microsoft.com/en-us/office/vba/api/excel.worksheet.setbackgroundpicture).
        */
        void SetBackgroundPicture(const wxString& fileName);

        /**
        Makes all rows of the currently filtered list visible. If AutoFilter is in use, this method changes the arrows to "All."

        [MSDN documentation for Worksheet.ShowAllData](https://learn.microsoft.com/en-us/office/vba/api/excel.worksheet.showalldata).
        */
        void ShowAllData();

        /**
        Displays the data form associated with the worksheet.

        [MSDN documentation for Worksheet.ShowDataForm](https://learn.microsoft.com/en-us/office/vba/api/excel.worksheet.showdataform).
        */
        void ShowDataForm();

        /**
        Removes protection from the worksheet. This method has no effect if it isn't protected.

        Returns true if the call succeeded.

        [MSDN documentation for Worksheet.Unprotect](https://learn.microsoft.com/en-us/office/vba/api/excel.worksheet.unprotect).
        */
        bool Unprotect(const wxString& password = wxEmptyString);


        // ***** PROPERTIES *****

        /**
        Returns an AutoFilter object if filtering is on.

        [MSDN documentation for Worksheet.AutoFilter](https://learn.microsoft.com/en-us/office/vba/api/excel.worksheet.autofilter).
        */
        wxExcelAutoFilter GetAutoFilter();

        /**
        True if the AutoFilter drop-down arrows are currently displayed on the sheet. This property is independent of the FilterMode property.

        [MSDN documentation for Worksheet.AutoFilterMode](https://learn.microsoft.com/en-us/office/vba/api/excel.worksheet.autofiltermode).
        */
        bool GetAutoFilterMode();
        /**
        Call to remove the filter drop-down arrows.

        [MSDN documentation for Worksheet.AutoFilterMode](https://learn.microsoft.com/en-us/office/vba/api/excel.worksheet.autofiltermode).
        */
        void SetAutoFilterMode();

        /**
        Returns a Range Represents the range containing the first circular reference on the sheet.

        [MSDN documentation for Worksheet.CircularReference](https://learn.microsoft.com/en-us/office/vba/api/excel.worksheet.circularreference).
        */
        wxExcelRange GetCircularReference();

        /**
        Returns a Comments collection that represents all the comments for the specified worksheet.

        [MSDN documentation for Worksheet.Comments](https://learn.microsoft.com/en-us/office/vba/api/excel.worksheet.comments).
        */
        wxExcelComments GetComments();

        /**
        Returns a CommentsThreaded collection that represents all the top-level/root comments (no replies) for the specified worksheet. Includes legacy and modern comments.

        [MSDN documentation for Worksheet.CommentsThreaded](https://learn.microsoft.com/en-us/office/vba/api/excel.worksheet.commentsthreaded).
        */
        wxExcelCommentsThreaded GetCommentsThreaded();

        /**
        Returns the function code used for the current consolidation.

        [MSDN documentation for Worksheet.ConsolidationFunction](https://learn.microsoft.com/en-us/office/vba/api/excel.worksheet.consolidationfunction).
        */
        XlConsolidationFunction GetConsolidationFunction();

        /**
        Returns a three-element array of consolidation options. If the element is 1, that option is set.

        [MSDN documentation for Worksheet.ConsolidationOptions](https://learn.microsoft.com/en-us/office/vba/api/excel.worksheet.consolidationoptions).
        */
        wxArrayShort GetConsolidationOptions();

        /**
        Returns an array of string values that name the source sheets for the worksheet's current consolidation. Returns an empty array if the worksheet has no consolidation.

        [MSDN documentation for Worksheet.ConsolidationSources](https://learn.microsoft.com/en-us/office/vba/api/excel.worksheet.consolidationsources).
        */
        wxArrayString GetConsolidationSources();

        /**
        Returns a CustomProperties object representing the identifier information associated with a worksheet.

        [MSDN documentation for Worksheet.CustomProperties](https://learn.microsoft.com/en-us/office/vba/api/excel.worksheet.customproperties).
        */
        wxExcelCustomProperties GetCustomProperties();

        /**
        True if page breaks (both automatic and manual) on the specified worksheet are displayed.

        [MSDN documentation for Worksheet.DisplayPageBreaks](https://learn.microsoft.com/en-us/office/vba/api/excel.worksheet.displaypagebreaks).
        */
        bool GetDisplayPageBreaks();
        /**
        True if page breaks (both automatic and manual) on the specified worksheet are displayed.

        [MSDN documentation for Worksheet.DisplayPageBreaks](https://learn.microsoft.com/en-us/office/vba/api/excel.worksheet.displaypagebreaks).
        */
        void SetDisplayPageBreaks(bool displayPageBreaks);

        /**
        True if the specified worksheet is displayed from right to left instead of from left to right. False if the object is displayed from left to right.

        [MSDN documentation for Worksheet.DisplayRightToLeft](https://learn.microsoft.com/en-us/office/vba/api/excel.worksheet.displayrighttoleft).
        */
        bool GetDisplayRightToLeft();

        /**
        Sets whether the specified worksheet is displayed from right to left instead of from left to right. Microsoft documents the property as read-only, but Excel sets it.

        [MSDN documentation for Worksheet.DisplayRightToLeft](https://learn.microsoft.com/en-us/office/vba/api/excel.worksheet.displayrighttoleft).
        */
        void SetDisplayRightToLeft(bool displayRightToLeft);

        /**
        True if AutoFilter arrows are enabled when user-interface-only protection is turned on.

        [MSDN documentation for Worksheet.EnableAutoFilter](https://learn.microsoft.com/en-us/office/vba/api/excel.worksheet.enableautofilter).
        */
        bool GetEnableAutoFilter();
        /**
        True if AutoFilter arrows are enabled when user-interface-only protection is turned on.

        [MSDN documentation for Worksheet.EnableAutoFilter](https://learn.microsoft.com/en-us/office/vba/api/excel.worksheet.enableautofilter).
        */
        void SetEnableAutoFilter(bool enableAutoFilter);

        /**
        True if Microsoft Excel automatically recalculates the worksheet when necessary. False if Excel doesn't recalculate the sheet.

        [MSDN documentation for Worksheet.EnableCalculation](https://learn.microsoft.com/en-us/office/vba/api/excel.worksheet.enablecalculation).
        */
        bool GetEnableCalculation();
        /**
        True if Microsoft Excel automatically recalculates the worksheet when necessary. False if Excel doesn't recalculate the sheet.

        [MSDN documentation for Worksheet.EnableCalculation](https://learn.microsoft.com/en-us/office/vba/api/excel.worksheet.enablecalculation).
        */
        void SetEnableCalculation(bool enableCalculation);

        /**
        Returms or sets if conditional formats will occur automatically as needed. Since MS Excel 2007.

        [MSDN documentation for Worksheet.EnableFormatConditionsCalculation](https://learn.microsoft.com/en-us/office/vba/api/excel.worksheet.enableformatconditionscalculation).
        */
        bool GetEnableFormatConditionsCalculation();
        /**
        Returms or sets if conditional formats will will occur automatically as needed. Since MS Excel 2007.

        [MSDN documentation for Worksheet.EnableFormatConditionsCalculation](https://learn.microsoft.com/en-us/office/vba/api/excel.worksheet.enableformatconditionscalculation).
        */
        void SetEnableFormatConditionsCalculation(bool enableFormatConditionsCalculation);

        /**
        True if outlining symbols are enabled when user-interface-only protection is turned on.

        [MSDN documentation for Worksheet.EnableOutlining](https://learn.microsoft.com/en-us/office/vba/api/excel.worksheet.enableoutlining).
        */
        bool GetEnableOutlining();
        /**
        True if outlining symbols are enabled when user-interface-only protection is turned on.

        [MSDN documentation for Worksheet.EnableOutlining](https://learn.microsoft.com/en-us/office/vba/api/excel.worksheet.enableoutlining).
        */
        void SetEnableOutlining(bool enableOutlining);

        /**
        True if PivotTable controls and actions are enabled when user-interface-only protection is turned on.

        [MSDN documentation for Worksheet.EnablePivotTable](https://learn.microsoft.com/en-us/office/vba/api/excel.worksheet.enablepivottable).
        */
        bool GetEnablePivotTable();
        /**
        True if PivotTable controls and actions are enabled when user-interface-only protection is turned on.

        [MSDN documentation for Worksheet.EnablePivotTable](https://learn.microsoft.com/en-us/office/vba/api/excel.worksheet.enablepivottable).
        */
        void SetEnablePivotTable(bool enablePivotTable);

        /**
        Returns what can be selected on the sheet.

        [MSDN documentation for Worksheet.EnableSelection](https://learn.microsoft.com/en-us/office/vba/api/excel.worksheet.enableselection).
        */
        XlEnableSelection GetEnableSelection();

        /**
        Sets what can be selected on the sheet.

        [MSDN documentation for Worksheet.EnableSelection](https://learn.microsoft.com/en-us/office/vba/api/excel.worksheet.enableselection).
        */
        void SetEnableSelection(XlEnableSelection enableSelection);

        /**
        True if the worksheet is in the filter mode.

        [MSDN documentation for Worksheet.FilterMode](https://learn.microsoft.com/en-us/office/vba/api/excel.worksheet.filtermode).
        */
        bool GetFilterMode();

        /**
        Returns a PageBreaks collection that represents the horizontal page breaks on the sheet.

        [MSDN documentation for Worksheet.HPageBreaks](https://learn.microsoft.com/en-us/office/vba/api/excel.worksheet.hpagebreaks).
        */
        wxExcelPageBreaks GetHPageBreaks();

        /**
        Returns a Hyperlinks collection that represents the hyperlinks for the worksheet.

        [MSDN documentation for Worksheet.Hyperlinks](https://learn.microsoft.com/en-us/office/vba/api/excel.worksheet.hyperlinks).
        */
        wxExcelHyperlinks GetHyperlinks();


        /**
        Returns the index number of the worksheet in the Worksheets collection.

        [MSDN documentation for Worksheet.Index](https://learn.microsoft.com/en-us/office/vba/api/excel.worksheet.index).
        */
        long GetIndex();

        /**
        Returns a collection of ListObject objects on the worksheet.

        [MSDN documentation for Worksheet.ListObjects](https://learn.microsoft.com/en-us/office/vba/api/excel.worksheet.listobjects).
        */
        wxExcelListObjects GetListObjects();

        /**
        Returns a String value representing the worksheet name.

        [MSDN documentation for Worksheet.Name](https://learn.microsoft.com/en-us/office/vba/api/excel.worksheet.name).
        */
        wxString GetName();
        /**
        Sets a String value representing the worksheet name.

        [MSDN documentation for Worksheet.Name](https://learn.microsoft.com/en-us/office/vba/api/excel.worksheet.name).
        */
        void SetName(const wxString& name);

        /**
        Returns a Names collection that represents all the worksheet-specific names (names defined with the "WorksheetName!" prefix)

        [MSDN documentation for Worksheet.Names](https://learn.microsoft.com/en-us/office/vba/api/excel.worksheet.names).
        */
        wxExcelNames GetNames();

        /**
        Returns a Worksheet Represents the next sheet.

        [MSDN documentation for Worksheet.Next](https://learn.microsoft.com/en-us/office/vba/api/excel.worksheet.next).
        */
        wxExcelWorksheet GetNext();

        /**
        Returns an Outline object that represents the outline for the specified worksheet.

        [MSDN documentation for Worksheet.Outline](https://learn.microsoft.com/en-us/office/vba/api/excel.worksheet.outline).
        */
        wxExcelOutline GetOutline();

        /**
       Returns an Outline object that represents the outline for the specified worksheet.

        [MSDN documentation for Worksheet.PageSetup](https://learn.microsoft.com/en-us/office/vba/api/excel.worksheet.pagesetup).
        */
        wxExcelPageSetup GetPageSetup();

        /**
        Returns a Worksheet Represents the next sheet.

        [MSDN documentation for Worksheet.Previous](https://learn.microsoft.com/en-us/office/vba/api/excel.worksheet.previous).
        */
        wxExcelWorksheet GetPrevious();

        /**
        Returns the number of comment pages that will be printed for the current worksheet. Since Excel 2010.

        [MSDN documentation for Worksheet.PrintedCommentPages](https://learn.microsoft.com/en-us/office/vba/api/excel.worksheet.printedcommentpages).
        */
        long GetPrintedCommentPages();

        /**
        True if the contents of the sheet are protected. This protects the individual cells. To turn on content protection, use the Protect method with the Contents argument set to True.

        [MSDN documentation for Worksheet.ProtectContents](https://learn.microsoft.com/en-us/office/vba/api/excel.worksheet.protectcontents).
        */
        bool GetProtectContents();

        /**
        True if shapes are protected. To turn on shape protection, use the Protect method with the DrawingObjects argument set to True.

        [MSDN documentation for Worksheet.ProtectDrawingObjects](https://learn.microsoft.com/en-us/office/vba/api/excel.worksheet.protectdrawingobjects).
        */
        bool GetProtectDrawingObjects();

        /**
        Returns a Protection object that represents the protection options of the worksheet.

        [MSDN documentation for Worksheet.Protection](https://learn.microsoft.com/en-us/office/vba/api/excel.worksheet.protection).
        */
        wxExcelProtection GetProtection();

        /**
        True if user-interface-only protection is turned on. To turn on user interface protection, use the Protect method with the UserInterfaceOnly argument set to True.

        [MSDN documentation for Worksheet.ProtectionMode](https://learn.microsoft.com/en-us/office/vba/api/excel.worksheet.protectionmode).
        */
        bool GetProtectionMode();

        /**
        True if the worksheet scenarios are protected.

        [MSDN documentation for Worksheet.ProtectScenarios](https://learn.microsoft.com/en-us/office/vba/api/excel.worksheet.protectscenarios).
        */
        bool GetProtectScenarios();

        /**
        Returns the range where scrolling is allowed, as an A1-style range reference. Cells outside the scroll area cannot be selected.

        [MSDN documentation for Worksheet.ScrollArea](https://learn.microsoft.com/en-us/office/vba/api/excel.worksheet.scrollarea).
        */
        wxString GetScrollArea();
        /**
        Sets the range where scrolling is allowed, as an A1-style range reference. Cells outside the scroll area cannot be selected.

        [MSDN documentation for Worksheet.ScrollArea](https://learn.microsoft.com/en-us/office/vba/api/excel.worksheet.scrollarea).
        */
        void SetScrollArea(const wxString& scrollArea);

#if WXAUTOEXCEL_USE_SHAPES
        /**
        Returns a Shapes collection that represents all the shapes on the worksheet.

        [MSDN documentation for Worksheet.Shapes](https://learn.microsoft.com/en-us/office/vba/api/excel.worksheet.shapes).
        */
        wxExcelShapes GetShapes();
#endif // #if WXAUTOEXCEL_USE_SHAPES

        /**
        Returns a Sort object.

        [MSDN documentation for Worksheet.Sort](https://learn.microsoft.com/en-us/office/vba/api/excel.worksheet.sort).
        */
        wxExcelSort GetSort();


        /**
        Returns the standard (default) height of all the rows in the worksheet, in points.

        [MSDN documentation for Worksheet.StandardHeight](https://learn.microsoft.com/en-us/office/vba/api/excel.worksheet.standardheight).
        */
        double GetStandardHeight();

        /**
        Returns the standard (default) width of all the columns in the worksheet.

        [MSDN documentation for Worksheet.StandardWidth](https://learn.microsoft.com/en-us/office/vba/api/excel.worksheet.standardwidth).
        */
        double GetStandardWidth();

        /**
        Sets the standard (default) width of all the columns in the worksheet.

        [MSDN documentation for Worksheet.StandardWidth](https://learn.microsoft.com/en-us/office/vba/api/excel.worksheet.standardwidth).
        */
        void SetStandardWidth(double standardWidth);

        /**
        Returns a Tab object for a worksheet.

        [MSDN documentation for Worksheet.Tab](https://learn.microsoft.com/en-us/office/vba/api/excel.worksheet.tab).
        */
        wxExcelTab GetTab();

        /**
        True if Microsoft Excel uses Lotus 1-2-3 expression evaluation rules for the worksheet.

        [MSDN documentation for Worksheet.TransitionExpEval](https://learn.microsoft.com/en-us/office/vba/api/excel.worksheet.transitionexpeval).
        */
        bool GetTransitionExpEval();
        /**
        True if Microsoft Excel uses Lotus 1-2-3 expression evaluation rules for the worksheet.

        [MSDN documentation for Worksheet.TransitionExpEval](https://learn.microsoft.com/en-us/office/vba/api/excel.worksheet.transitionexpeval).
        */
        void SetTransitionExpEval(bool transitionExpEval);

        /**
        True if Microsoft Excel uses Lotus 1-2-3 formula entry rules for the worksheet.

        [MSDN documentation for Worksheet.TransitionFormEntry](https://learn.microsoft.com/en-us/office/vba/api/excel.worksheet.transitionformentry).
        */
        bool GetTransitionFormEntry();
        /**
        True if Microsoft Excel uses Lotus 1-2-3 formula entry rules for the worksheet.

        [MSDN documentation for Worksheet.TransitionFormEntry](https://learn.microsoft.com/en-us/office/vba/api/excel.worksheet.transitionformentry).
        */
        void SetTransitionFormEntry(bool transitionFormEntry);

        /**
        Returns an XlSheetType value that represents the worksheet type.

        [MSDN documentation for Worksheet.Type](https://learn.microsoft.com/en-us/office/vba/api/excel.worksheet.type).
        */
        XlSheetType GetType();

        /**
        Returns a Range Represents the used range on the specified worksheet.

        [MSDN documentation for Worksheet.UsedRange](https://learn.microsoft.com/en-us/office/vba/api/excel.worksheet.usedrange).
        */
        wxExcelRange GetUsedRange();

        /**
        Returns an XlSheetVisibility value that determines whether the object is visible.

        [MSDN documentation for Worksheet.Visible](https://learn.microsoft.com/en-us/office/vba/api/excel.worksheet.visible).
        */
        XlSheetVisibility GetVisible();
        /**
        Sets an XlSheetVisibility value that determines whether the object is visible.

        [MSDN documentation for Worksheet.Visible](https://learn.microsoft.com/en-us/office/vba/api/excel.worksheet.visible).
        */
        void SetVisible(XlSheetVisibility visible);

        /**
        Returns a PageBreaks collection that represents the vertical page breaks on the sheet.

        [MSDN documentation for Worksheet.VPageBreaks](https://learn.microsoft.com/en-us/office/vba/api/excel.worksheet.vpagebreaks).
        */
        wxExcelPageBreaks GetVPageBreaks();

        /**
        Automatically converts the wxExcelWorksheet object to wxExcelSheet object so it can be used anywhere wxExcelSheet can.
        */
        operator wxExcelSheet();

        /**
        Returns "Worksheet".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("Worksheet"); }
    private:
        bool DoOrderedCopyOrMove(bool copy, wxExcelSheet sheetAfterOrBefore, bool after);
    };


} // namespace wxAutoExcel

#endif //_WXAUTOEXCEL_WORKSHEET_H
