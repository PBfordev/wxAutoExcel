/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Modified by:
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////


#ifndef _WXAUTOEXCEL_WORKSHEET_H
#define _WXAUTOEXCEL_WORKSHEET_H

#include "wx/wxAutoExcel_defs.h"

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

        [MSDN documentation for Worksheet.Activate](http://msdn.microsoft.com/en-us/library/bb179169.aspx).
        */

        bool Activate();

        /**
        Calculates a specific worksheet.

        [MSDN documentation for Worksheet.Calculate](http://msdn.microsoft.com/en-us/library/bb179170.aspx).
        */
        bool Calculate();

#if WXAUTOEXCEL_USE_CHARTS
        //@{
        /**
        Returns an object that represents either a single embedded chart (a ChartObject object) or a collection of all the embedded charts (a ChartObjects object) on the sheet.

        [MSDN documentation for Worksheet.ChartObjects](http://msdn.microsoft.com/en-us/library/bb179172.aspx).
        */
        wxExcelChartObjects ChartObjects();
        wxExcelChartObjects ChartObjects(const wxVector<long>& indices);
        wxExcelChartObjects ChartObjects(const wxArrayString& names);        
        //@}
#endif // #if WXAUTOEXCEL_USE_CHARTS

        /**
        Checks the spelling of a worksheet.

        [MSDN documentation for Worksheet.CheckSpelling](http://msdn.microsoft.com/en-us/library/bb179174.aspx).
        */
        bool CheckSpelling(const wxString& customDictionary = wxEmptyString, wxXlTribool ignoreUpperCase = wxDefaultXlTribool,
            wxXlTribool alwaysSuggest = wxDefaultXlTribool, MsoLanguageID* spellLang = NULL);

        /**
        Circles invalid entries on the worksheet.

        [MSDN documentation for Worksheet.CircleInvalid](http://msdn.microsoft.com/en-us/library/bb223252.aspx).
        */
        void CircleInvalid();

        /**
        Clears the tracer arrows from the worksheet. Tracer arrows are added by using the auditing feature.

        [MSDN documentation for Worksheet.ClearArrows](http://msdn.microsoft.com/en-us/library/bb223254.aspx).
        */
        bool ClearArrows();

        /**
        Clears circles from invalid entries on the worksheet.

        [MSDN documentation for Worksheet.ClearCircles](http://msdn.microsoft.com/en-us/library/bb223258.aspx).
        */
        bool ClearCircles();

        //@{
        /**
        Method without parameters copies the worksheet into a new workbook. Method with parameters copies the worksheet to another location in the workbook.

        [MSDN documentation for Worksheet.Copy](http://msdn.microsoft.com/en-us/library/bb179176.aspx).
        */
        bool Copy();
        bool CopyAfterOrBefore(wxExcelSheet worksheetAfterOrBefore, bool after);
        //@}

        /**
        Deletes the object.

        [MSDN documentation for Worksheet.Delete](http://msdn.microsoft.com/en-us/library/bb179178.aspx).
        */
        bool Delete();


        //@{
        /**
        Exports to a file of the specified format.

        [MSDN documentation for Worksheet.ExportAsFixedFormat](http://msdn.microsoft.com/en-us/library/bb238919.aspx).
        */
        bool ExportAsFixedFormat(XlFixedFormatType type, const wxString& fileName = wxEmptyString,
            XlFixedFormatQuality* quality = NULL, wxXlTribool includeDocProperties = wxDefaultXlTribool,
            wxXlTribool ignorePrintAreas = wxDefaultXlTribool,
            long* from = NULL, long* to = NULL, wxXlTribool openAfterPublish = wxDefaultXlTribool);

        bool ExportAsFixedFormat(XlFixedFormatType type, const wxVariantVector& optionalArgs);
        //@}

        //@{
        /**
        Method without parameters moves the worksheet into a new workbook. Method with parameters moves the worksheet to another location in the workbook.

        [MSDN documentation for Worksheet.Move](http://msdn.microsoft.com/en-us/library/bb179196.aspx).
        */
        bool Move();
        bool MoveAfterOrBefore(wxExcelSheet sheetAfterOrBefore, bool after);
        //@}
        
        /**
        Returns an Represents either a single OLE object (an OLEObject ) or a collection of all OLE objects (an OLEObjects collection) on the chart or sheet. Read-only.

        [MSDN documentation for Worksheet.OLEObjects](http://msdn.microsoft.com/en-us/library/bb179198.aspx).
        */
        wxExcelOLEObjects OLEObjects();
        
        /**
        Pastes the contents of the Clipboard onto the worksheet.

        [MSDN documentation for Worksheet.Paste](http://msdn.microsoft.com/en-us/library/bb179200.aspx).
        */
        bool Paste(wxExcelRange* destination = NULL, wxXlTribool link = wxDefaultXlTribool);

        /**
        Pastes the contents of the Clipboard onto the worksheet, using a specified format. Use this method to paste data from other applications or to paste data in a specific format.

        [MSDN documentation for Worksheet.PasteSpecial](http://msdn.microsoft.com/en-us/library/bb179201.aspx).
        */
        bool PasteSpecial(const wxString& format = wxEmptyString, wxXlTribool link = wxDefaultXlTribool,
            wxXlTribool displayAsIcon = wxDefaultXlTribool, const wxString& iconFileName = wxEmptyString,
            long* iconIndex = NULL, const wxString& iconLabel = wxEmptyString,
            wxXlTribool noHTMLFormatting = wxDefaultXlTribool);


        //@{
        /**
        Prints the worksheet.

        [MSDN documentation for Worksheet.PrintOut](http://msdn.microsoft.com/en-us/library/bb179206.aspx).
        */
        bool PrintOut(long* from = NULL, long* to = NULL, long* copies = NULL, wxXlTribool preview = wxDefaultXlTribool,
            const wxString& activePrinter = wxEmptyString, wxXlTribool printToFile = wxDefaultXlTribool,
            wxXlTribool collate = wxDefaultXlTribool, const wxString& prToFileName= wxEmptyString, wxXlTribool ignorePrintAreas = wxDefaultXlTribool);
        bool PrintOut(const wxVariantVector& args);
        //@}

        /**
        Shows a preview of the worksheet as it would look when printed.

        [MSDN documentation for Worksheet.PrintPreview](http://msdn.microsoft.com/en-us/library/bb179208.aspx).
        */
        bool PrintPreview(wxXlTribool enableChanges = wxDefaultXlTribool);

        /**
        Protects a worksheet so that it cannot be modified.

        [MSDN documentation for Worksheet.Protect](http://msdn.microsoft.com/en-us/library/bb179210.aspx).
        */
        void Protect(const wxString& password = wxEmptyString, wxXlTribool structure = wxDefaultXlTribool, wxXlTribool windows = wxDefaultXlTribool);

        /**
        Resets all page breaks on the specified worksheet.

        [MSDN documentation for Worksheet.ResetAllPageBreaks](http://msdn.microsoft.com/en-us/library/bb177960.aspx).
        */
        void ResetAllPageBreaks();

        //@{
        /**
        Saves changes to the worksheet in a different file.

        [MSDN documentation for Worksheet.SaveAs](http://msdn.microsoft.com/en-us/library/bb214156.aspx).
        */
        bool SaveAs(const wxString& fileName = wxEmptyString, XlFileFormat* fileFormat = NULL,
            const wxString& password = wxEmptyString, const wxString& writeResPassword = wxEmptyString,
            wxXlTribool readOnlyRecommended = wxDefaultXlTribool, wxXlTribool createBackup = wxDefaultXlTribool,
            wxXlTribool addToMru = wxDefaultXlTribool, wxXlTribool local = wxDefaultXlTribool);

        bool SaveAs(const wxVariantVector& optionalArgs);
        //@}        

        /**
        Selects the worksheet.

        [MSDN documentation for Worksheet.Select](http://msdn.microsoft.com/en-us/library/bb214162.aspx).
        */
        bool Select(wxXlTribool replace = wxDefaultXlTribool);

        /**
        Sets the background graphic for a worksheet.

        [MSDN documentation for Worksheet.SetBackgroundPicture](http://msdn.microsoft.com/en-us/library/bb214164.aspx).
        */
        void SetBackgroundPicture(const wxString& fileName);

        /**
        Makes all rows of the currently filtered list visible. If AutoFilter is in use, this method changes the arrows to "All."

        [MSDN documentation for Worksheet.ShowAllData](http://msdn.microsoft.com/en-us/library/bb178108.aspx).
        */
        void ShowAllData();

        /**
        Displays the data form associated with the worksheet.

        [MSDN documentation for Worksheet.ShowDataForm](http://msdn.microsoft.com/en-us/library/bb178113.aspx).
        */
        void ShowDataForm();

        /**
        Removes protection from the worksheet. This method has no effect if it isn't protected.

        [MSDN documentation for Worksheet.Unprotect](http://msdn.microsoft.com/en-us/library/bb214174.aspx).
        */
        void Unprotect(const wxString& password = wxEmptyString);


        // ***** PROPERTIES *****         

        /**
        Returns an AutoFilter object if filtering is on.

        [MSDN documentation for Worksheet.AutoFilter](http://msdn.microsoft.com/en-us/library/bb148832.aspx).
        */
        wxExcelAutoFilter GetAutoFilter();

        /**
        True if the AutoFilter drop-down arrows are currently displayed on the sheet. This property is independent of the FilterMode property.

        [MSDN documentation for Worksheet.AutoFilterMode](http://msdn.microsoft.com/en-us/library/bb220852.aspx).
        */
        bool GetAutoFilterMode();
        /**
        Call to remove the filter drop-down arrows.

        [MSDN documentation for Worksheet.AutoFilterMode](http://msdn.microsoft.com/en-us/library/bb220852.aspx).
        */
        void SetAutoFilterMode();

        /**
        Returns a Range Represents the range containing the first circular reference on the sheet.

        [MSDN documentation for Worksheet.CircularReference](http://msdn.microsoft.com/en-us/library/bb177355.aspx).
        */
        wxExcelRange GetCircularReference();

        /**
        Returns a Comments collection that represents all the comments for the specified worksheet.

        [MSDN documentation for Worksheet.Comments](http://msdn.microsoft.com/en-us/library/bb177380.aspx).
        */
        wxExcelComments GetComments();

        /**
        Returns the function code used for the current consolidation. 

        [MSDN documentation for Worksheet.ConsolidationFunction](http://msdn.microsoft.com/en-us/library/bb177383.aspx).
        */
        XlConsolidationFunction GetConsolidationFunction();

        /**
        Returns a three-element array of consolidation options. If the element is 1, that option is set.

        [MSDN documentation for Worksheet.ConsolidationOptions](http://msdn.microsoft.com/en-us/library/bb177385.aspx).
        */
        wxArrayShort GetConsolidationOptions();

        /**
        Returns an array of string values that name the source sheets for the worksheet's current consolidation. 

        [MSDN documentation for Worksheet.ConsolidationSources](http://msdn.microsoft.com/en-us/library/bb177387.aspx).
        */
        wxArrayString GetConsolidationSources();

        /**
        True if page breaks (both automatic and manual) on the specified worksheet are displayed.

        [MSDN documentation for Worksheet.DisplayPageBreaks](http://msdn.microsoft.com/en-us/library/bb220979.aspx).
        */
        bool GetDisplayPageBreaks();
        /**
        True if page breaks (both automatic and manual) on the specified worksheet are displayed.

        [MSDN documentation for Worksheet.DisplayPageBreaks](http://msdn.microsoft.com/en-us/library/bb220979.aspx).
        */
        void SetDisplayPageBreaks(bool displayPageBreaks);

        /**
        True if the specified worksheet is displayed from right to left instead of from left to right. False if the object is displayed from left to right.

        [MSDN documentation for Worksheet.DisplayRightToLeft](http://msdn.microsoft.com/en-us/library/bb148848.aspx).
        */
        bool GetDisplayRightToLeft();

        /**
        True if AutoFilter arrows are enabled when user-interface-only protection is turned on.

        [MSDN documentation for Worksheet.EnableAutoFilter](http://msdn.microsoft.com/en-us/library/bb221123.aspx).
        */
        bool GetEnableAutoFilter();
        /**
        True if AutoFilter arrows are enabled when user-interface-only protection is turned on.

        [MSDN documentation for Worksheet.EnableAutoFilter](http://msdn.microsoft.com/en-us/library/bb221123.aspx).
        */
        void SetEnableAutoFilter(bool enableAutoFilter);

        /**
        True if Microsoft Excel automatically recalculates the worksheet when necessary. False if Excel doesn't recalculate the sheet.

        [MSDN documentation for Worksheet.EnableCalculation](http://msdn.microsoft.com/en-us/library/bb221135.aspx).
        */
        bool GetEnableCalculation();
        /**
        True if Microsoft Excel automatically recalculates the worksheet when necessary. False if Excel doesn't recalculate the sheet.

        [MSDN documentation for Worksheet.EnableCalculation](http://msdn.microsoft.com/en-us/library/bb221135.aspx).
        */
        void SetEnableCalculation(bool enableCalculation);

        /**
        Returms or sets if conditional formats will occur automatically as needed. Since MS Excel 2007.

        [MSDN documentation for Worksheet.EnableFormatConditionsCalculation](http://msdn.microsoft.com/en-us/library/bb216237.aspx).
        */
        bool GetEnableFormatConditionsCalculation();
        /**
        Returms or sets if conditional formats will will occur automatically as needed. Since MS Excel 2007.

        [MSDN documentation for Worksheet.EnableFormatConditionsCalculation](http://msdn.microsoft.com/en-us/library/bb216237.aspx).
        */
        void SetEnableFormatConditionsCalculation(bool enableFormatConditionsCalculation);

        /**
        True if outlining symbols are enabled when user-interface-only protection is turned on.

        [MSDN documentation for Worksheet.EnableOutlining](http://msdn.microsoft.com/en-us/library/bb221161.aspx).
        */
        bool GetEnableOutlining();
        /**
        True if outlining symbols are enabled when user-interface-only protection is turned on.

        [MSDN documentation for Worksheet.EnableOutlining](http://msdn.microsoft.com/en-us/library/bb221161.aspx).
        */
        void SetEnableOutlining(bool enableOutlining);

        /**
        True if PivotTable controls and actions are enabled when user-interface-only protection is turned on.

        [MSDN documentation for Worksheet.EnablePivotTable](http://msdn.microsoft.com/en-us/library/bb221164.aspx).
        */
        bool GetEnablePivotTable();
        /**
        True if PivotTable controls and actions are enabled when user-interface-only protection is turned on.

        [MSDN documentation for Worksheet.EnablePivotTable](http://msdn.microsoft.com/en-us/library/bb221164.aspx).
        */
        void SetEnablePivotTable(bool enablePivotTable);

        /**
        Returns what can be selected on the sheet.

        [MSDN documentation for Worksheet.EnableSelection](http://msdn.microsoft.com/en-us/library/bb221170.aspx).
        */
        XlEnableSelection GetEnableSelection();

        /**
        True if the worksheet is in the filter mode.

        [MSDN documentation for Worksheet.FilterMode](http://msdn.microsoft.com/en-us/library/bb148850.aspx).
        */
        bool GetFilterMode();

        /**
        Returns a PageBreaks collection that represents the horizontal page breaks on the sheet.

        [MSDN documentation for Worksheet.HPageBreaks](http://msdn.microsoft.com/en-us/library/bb148855.aspx).
        */
        wxExcelPageBreaks GetHPageBreaks();

        /**
        Returns a Hyperlinks collection that represents the hyperlinks for the worksheet.

        [MSDN documentation for Worksheet.Hyperlinks](http://msdn.microsoft.com/en-us/library/bb148859.aspx).
        */
        wxExcelHyperlinks GetHyperlinks();


        /**
        Returns the index number of the worksheet in the Worksheets collection.

        [MSDN documentation for Worksheet.Index](http://msdn.microsoft.com/en-us/library/bb148863.aspx).
        */
        long GetIndex();        

        /**
        Returns a String value representing the worksheet name.

        [MSDN documentation for Worksheet.Name](http://msdn.microsoft.com/en-us/library/bb148869.aspx).
        */
        wxString GetName();
        /**
        Sets a String value representing the worksheet name.

        [MSDN documentation for Worksheet.Name](http://msdn.microsoft.com/en-us/library/bb148869.aspx).
        */
        void SetName(const wxString& name);

        /**
        Returns a Worksheet Represents the next sheet.

        [MSDN documentation for Worksheet.Next](http://msdn.microsoft.com/en-us/library/bb148878.aspx).
        */
        wxExcelWorksheet GetNext();

        /**
        Returns a PageSetup object that contains all the page setup settings for the worksheet.

        [MSDN documentation for Worksheet.PageSetup](http://msdn.microsoft.com/en-us/library/bb148891.aspx).
        */
        wxExcelPageSetup GetPageSetup();

        /**
        Returns a Worksheet Represents the next sheet.

        [MSDN documentation for Worksheet.Previous](http://msdn.microsoft.com/en-us/library/bb238395.aspx).
        */
        wxExcelWorksheet GetPrevious();

        /**
        Returns the number of comment pages that will be printed for the current worksheet. Since Excel 2010.

        [MSDN documentation for Worksheet.PrintedCommentPages](http://msdn.microsoft.com/en-us/library/office/ff196864%28v=office.14%29.aspx).
        */
        long GetPrintedCommentPages();

        /**
        True if the contents of the sheet are protected. This protects the individual cells. To turn on content protection, use the Protect method with the Contents argument set to True.

        [MSDN documentation for Worksheet.ProtectContents](http://msdn.microsoft.com/en-us/library/bb238406.aspx).
        */
        bool GetProtectContents();
        
        /**
        True if shapes are protected. To turn on shape protection, use the Protect method with the DrawingObjects argument set to True.

        [MSDN documentation for Worksheet.ProtectDrawingObjects](http://msdn.microsoft.com/en-us/library/bb238414.aspx).
        */
        bool GetProtectDrawingObjects();

        /**
        True if user-interface-only protection is turned on. To turn on user interface protection, use the Protect method with the UserInterfaceOnly argument set to True.

        [MSDN documentation for Worksheet.ProtectionMode](http://msdn.microsoft.com/en-us/library/bb238419.aspx).
        */
        bool GetProtectionMode();

        /**
        True if the worksheet scenarios are protected.

        [MSDN documentation for Worksheet.ProtectScenarios](http://msdn.microsoft.com/en-us/library/bb209041.aspx).
        */
        bool GetProtectScenarios();        

        /**
        Returns the range where scrolling is allowed, as an A1-style range reference. Cells outside the scroll area cannot be selected.

        [MSDN documentation for Worksheet.ScrollArea](http://msdn.microsoft.com/en-us/library/bb221614.aspx).
        */
        wxString GetScrollArea();
        /**
        Sets the range where scrolling is allowed, as an A1-style range reference. Cells outside the scroll area cannot be selected.

        [MSDN documentation for Worksheet.ScrollArea](http://msdn.microsoft.com/en-us/library/bb221614.aspx).
        */
        void SetScrollArea(const wxString& scrollArea);

#if WXAUTOEXCEL_USE_SHAPES
        /**
        Returns a Shapes collection that represents all the shapes on the worksheet.

        [MSDN documentation for Worksheet.Shapes](http://msdn.microsoft.com/en-us/library/bb215254.aspx).
        */
        wxExcelShapes GetShapes();
#endif // #if WXAUTOEXCEL_USE_SHAPES

        /**
        Returns the standard (default) height of all the rows in the worksheet, in points. 

        [MSDN documentation for Worksheet.StandardHeight](http://msdn.microsoft.com/en-us/library/bb209292.aspx).
        */
        double GetStandardHeight();

        /**
        Returns the standard (default) width of all the columns in the worksheet.

        [MSDN documentation for Worksheet.StandardWidth](http://msdn.microsoft.com/en-us/library/bb209294.aspx).
        */
        double GetStandardWidth();

        /**
        Sets the standard (default) width of all the columns in the worksheet. 

        [MSDN documentation for Worksheet.StandardWidth](http://msdn.microsoft.com/en-us/library/bb209294.aspx).
        */
        void SetStandardWidth(double standardWidth);

        /**
        Returns a Tab object for a worksheet.

        [MSDN documentation for Worksheet.Tab](http://msdn.microsoft.com/en-us/library/bb224514.aspx).
        */
        wxExcelTab GetTab();

        /**
        True if Microsoft Excel uses Lotus 1-2-3 expression evaluation rules for the worksheet.

        [MSDN documentation for Worksheet.TransitionExpEval](http://msdn.microsoft.com/en-us/library/bb221886.aspx).
        */
        bool GetTransitionExpEval();
        /**
        True if Microsoft Excel uses Lotus 1-2-3 expression evaluation rules for the worksheet.

        [MSDN documentation for Worksheet.TransitionExpEval](http://msdn.microsoft.com/en-us/library/bb221886.aspx).
        */
        void SetTransitionExpEval(bool transitionExpEval);

        /**
        True if Microsoft Excel uses Lotus 1-2-3 formula entry rules for the worksheet.

        [MSDN documentation for Worksheet.TransitionFormEntry](http://msdn.microsoft.com/en-us/library/bb221893.aspx).
        */
        bool GetTransitionFormEntry();
        /**
        True if Microsoft Excel uses Lotus 1-2-3 formula entry rules for the worksheet.

        [MSDN documentation for Worksheet.TransitionFormEntry](http://msdn.microsoft.com/en-us/library/bb221893.aspx).
        */
        void SetTransitionFormEntry(bool transitionFormEntry);

        /**
        Returns an XlSheetType value that represents the worksheet type.

        [MSDN documentation for Worksheet.Type](http://msdn.microsoft.com/en-us/library/bb224517.aspx).
        */
        XlSheetType GetType();

        /**
        Returns a Range Represents the used range on the specified worksheet.

        [MSDN documentation for Worksheet.UsedRange](http://msdn.microsoft.com/en-us/library/bb221970.aspx).
        */
        wxExcelRange GetUsedRange();

        /**
        Returns an XlSheetVisibility value that determines whether the object is visible.

        [MSDN documentation for Worksheet.Visible](http://msdn.microsoft.com/en-us/library/bb224519.aspx).
        */
        XlSheetVisibility GetVisible();
        /**
        Sets an XlSheetVisibility value that determines whether the object is visible.

        [MSDN documentation for Worksheet.Visible](http://msdn.microsoft.com/en-us/library/bb224519.aspx).
        */
        void SetVisible(XlSheetVisibility visible);

        /**
        Returns a PageBreaks collection that represents the vertical page breaks on the sheet.

        [MSDN documentation for Worksheet.VPageBreaks](http://msdn.microsoft.com/en-us/library/bb224522.aspx).
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
