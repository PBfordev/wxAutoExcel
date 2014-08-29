/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Modified by:
// RCS-ID:      $Id: $
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////


#ifndef _WXAUTOEXCEL_APPLICATION_H
#define _WXAUTOEXCEL_APPLICATION_H

#include "wx/wxAutoExcel_defs.h"
#include "wx/wxAutoExcel_enums.h"
#include "wx/wxAutoExcelRange.h"

namespace wxAutoExcel {
    /**
    @brief Represents Microsoft Excel Application.
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelApplication : public wxExcelRangeOwner
    {
    public:

        /**
        Launches MS Excel and creates an instance of MS Excel Application. If failed then returned wxExcelApplication::IsOk_() returns false.

        */
        static wxExcelApplication CreateInstance();

        /**
        Analogical to wxAutomationObject::GetInstance() If failed then returned wxExcelApplication::IsOk_() returns false.

        */
        static wxExcelApplication GetInstance(int flags = wxAutomationInstance_CreateIfNeeded);

        // ***** METHODS *****

        /**
        Activates a Microsoft application. If the application is already running, this method activates the running application. If the application is not running, this method starts a new instance of the application.

        [MSDN documentation for Application.ActivateMicrosoftApp](http://msdn.microsoft.com/en-us/library/bb209531.aspx).
        */
        void ActivateMicrosoftApp(XlMSApplication index);

        //@{
        /**
        Adds a custom list for custom autofill and/or custom sort.

        [MSDN documentation for Application.AddCustomList](http://msdn.microsoft.com/en-us/library/bb209560.aspx).
        */
        bool AddCustomList(const wxArrayString& listArray);
        bool AddCustomList(wxExcelRange listArray, bool byRow);
        //@}

        /**
        Calculates all open workbooks.

        [MSDN documentation for Application.Calculate](http://msdn.microsoft.com/en-us/library/bb211549.aspx).
        */
        void Calculate();

        /**
        Forces a full calculation of the data in all open workbooks.

        [MSDN documentation for Application.CalculateFull](http://msdn.microsoft.com/en-us/library/bb223210.aspx).
        */
        void CalculateFull();

        /**
        For all open workbooks, forces a full calculation of the data and rebuilds the dependencies.

        [MSDN documentation for Application.CalculateFullRebuild](http://msdn.microsoft.com/en-us/library/bb223213.aspx).
        */
        void CalculateFullRebuild();

        /**
        Runs all pending queries to OLEDB and OLAP data sources.

        [MSDN documentation for Application.CalculateUntilAsyncQueriesDone](http://msdn.microsoft.com/en-us/library/bb242478.aspx).
        */
        void CalculateUntilAsyncQueriesDone();

        /**
        Converts a measurement from centimeters to points (one point equals 0.035 centimeters).

        [MSDN documentation for Application.CentimetersToPoints](http://msdn.microsoft.com/en-us/library/bb223223.aspx).
        */
        double CentimetersToPoints(double centimeters);

        /**
        Stops recalculation. If keepAbort is a valid range, the recalculation will be finished for it.

        [MSDN documentation for Application.CheckAbort](http://msdn.microsoft.com/en-us/library/bb223243.aspx).
        */
        void CheckAbort(wxExcelRange* keepAbort = NULL);

        /**
        Checks the spelling of a single word.

        [MSDN documentation for Application.CheckSpelling](http://msdn.microsoft.com/en-us/library/bb211557.aspx).
        */
        bool CheckSpelling(const wxString& word, const wxString& customDictionary = wxEmptyString, wxXlTribool ignoreUpperCase = wxDefaultXlTribool);

        /**
        Converts cell references in a formula between the A1 and R1C1 reference styles, between relative and absolute references, or both.

        [MSDN documentation for Application.ConvertFormula](http://msdn.microsoft.com/en-us/library/bb223281.aspx).
        */
        wxString ConvertFormula(const wxString& formula, XlReferenceStyle fromReferenceStyle,
                                XlReferenceStyle* toReferenceStyle = NULL, XlReferenceStyle* toAbsolute = NULL,
                                wxExcelRange* relativeTo = NULL);

        /**
        Runs a command or performs some other action or actions in another application by way of the specified DDE channel.

        [MSDN documentation for Application.DDEExecute](http://msdn.microsoft.com/en-us/library/bb223316.aspx).
        */
        void DDEExecute(long channel, const wxString& str);

        /**
        Opens a DDE channel to an application.

        [MSDN documentation for Application.DDEInitiate](http://msdn.microsoft.com/en-us/library/bb223318.aspx).
        */
        long DDEInitiate(const wxString& app, const wxString& topic);

        /**
        Sends data to an application.

        [MSDN documentation for Application.DDEPoke](http://msdn.microsoft.com/en-us/library/bb223320.aspx).
        */
        void DDEPoke(long channel, const wxVariant& item, const wxVariant& data);

        /**
        Requests information from the specified application. This method always returns a list.

        [MSDN documentation for Application.DDERequest](http://msdn.microsoft.com/en-us/library/bb223322.aspx).
        */
        wxVariant DDERequest(long channel, const wxString& item);

        /**
        Closes a channel to another application.

        [MSDN documentation for Application.DDETerminate](http://msdn.microsoft.com/en-us/library/bb223325.aspx).
        */
        void DDETerminate(long channel);

        /**
        Deletes a custom list.

        [MSDN documentation for Application.DeleteCustomList](http://msdn.microsoft.com/en-us/library/bb223327.aspx).
        */
        bool DeleteCustomList(long listNum);


        /**
        Equivalent to double-clicking the active cell.

        [MSDN documentation for Application.DoubleClick](http://msdn.microsoft.com/en-us/library/bb209796.aspx).
        */
        void DoubleClick();


        /**
        Displays the Open dialog box. Returns true if the file chosen by the user was successfully open, false otherwise.

        [MSDN documentation for Application.FindFile](http://msdn.microsoft.com/en-us/library/bb209853.aspx).
        */
        bool FindFile();

        /**
        Returns a custom list.

        [MSDN documentation for Application.GetCustomListContents](http://msdn.microsoft.com/en-us/library/bb209881.aspx).
        */
        wxArrayString GetCustomListContents(long listNum);

        /**
        Returns the custom list number for an array of strings. You can use this method to match both built-in lists and custom-defined lists.

        [MSDN documentation for Application.GetCustomListNum](http://msdn.microsoft.com/en-us/library/bb209884.aspx).
        */
        long GetCustomListNum(const wxArrayString& listArray);

        /**
        Displays the standard Open dialog box and gets a file name from the user without actually opening any files.

        [MSDN documentation for Application.GetOpenFilename](http://msdn.microsoft.com/en-us/library/bb209892.aspx).
        */
        wxArrayString GetOpenFilename(const wxString& fileFilter = wxEmptyString, long* filterIndex = NULL,
                                      const wxString& title = wxEmptyString, wxXlTribool multiSelect = wxDefaultXlTribool);

        /**
        Returns the Japanese phonetic text of the specified text string. This method is available to you only if you have selected or installed Japanese language support for Microsoft Office.

        [MSDN documentation for Application.GetPhonetic](http://msdn.microsoft.com/en-us/library/bb209895.aspx).
        */
        wxString GetPhonetic(const wxString& text);

        /**
        Displays the standard Save As dialog box and gets a file name from the user without actually saving any files.

        [MSDN documentation for Application.GetSaveAsFilename](http://msdn.microsoft.com/en-us/library/bb209903.aspx).
        */
        wxString GetSaveAsFilename(const wxString& initialFilename = wxEmptyString,
                                   const wxString& fileFilter = wxEmptyString, long* filterIndex = NULL,
                                   const wxString& title = wxEmptyString);
         //@{
        /**
        Selects any range or Visual Basic procedure in any workbook, and activates that workbook if it is not already active.

        [MSDN documentation for Application.Goto](http://msdn.microsoft.com/en-us/library/bb209910.aspx).
        */
        void Goto(const wxString& reference = wxEmptyString, bool scroll = false);
        void Goto(wxExcelRange range, bool scroll = false);
         //@}

        /**
        Displays a Help topic.

        [MSDN documentation for Application.Help](http://msdn.microsoft.com/en-us/library/bb209917.aspx).
        */
        void Help(const wxString& helpFile = wxEmptyString, long* helpContextID = NULL);

        /**
        Converts a measurement from inches to points.

        [MSDN documentation for Application.InchesToPoints](http://msdn.microsoft.com/en-us/library/bb209927.aspx).
        */
        double InchesToPoints(double inches);        

        /**
        Returns a Range representing the rectangular intersection of two or more ranges.

        [MSDN documentation for Application.Intersect](http://msdn.microsoft.com/en-us/library/bb209961.aspx).
        */
        wxExcelRange Intersect(const wxExcelRangeVector& ranges);

        /**
        Closes a MAPI mail session established by Microsoft Excel.

        [MSDN documentation for Application.MailLogoff](http://msdn.microsoft.com/en-us/library/bb223477.aspx).
        */
        void MailLogoff();

        /**
        Logs in to MAPI Mail or Microsoft Exchange and establishes a mail session. If Microsoft Mail is not already running, you must use this method to establish a mail session before mail or document routing functions can be used.

        [MSDN documentation for Application.MailLogon](http://msdn.microsoft.com/en-us/library/bb223480.aspx).
        */
        void MailLogon(const wxString& name = wxEmptyString, const wxString& password = wxEmptyString, 
                       wxXlTribool downloadNewMail = wxDefaultXlTribool);

        /**
        Quits Microsoft Excel.

        [MSDN documentation for Application.Quit](http://msdn.microsoft.com/en-us/library/bb223560.aspx).
        */
        void Quit();

        /**
        Records code if the macro recorder is on.

        [MSDN documentation for Application.RecordMacro](http://msdn.microsoft.com/en-us/library/bb223571.aspx).
        */
        void RecordMacro(const wxString& basicCode, const wxString& XlmCode = wxEmptyString);

        /**
        Loads an XLL code resource and automatically registers the functions and commands contained in the resource.

        [MSDN documentation for Application.RegisterXLL](http://msdn.microsoft.com/en-us/library/bb223580.aspx).
        */
        bool RegisterXLL(const wxString& fileName);


        /**
        Runs a macro.

        [MSDN documentation for Application.Run](http://msdn.microsoft.com/en-us/library/bb213784.aspx).
        */
        wxVariant Run(const wxString& macro, const wxVariantVector& arguments);

        /**
        Saves the current workspace.

        [MSDN documentation for Application.SaveWorkspace](http://msdn.microsoft.com/en-us/library/bb178007.aspx).
        */
        void SaveWorkspace(const wxString& fileName = wxEmptyString);

        /**
        Sends keystrokes to Excel application.

        [MSDN documentation for Application.SendKeys](http://msdn.microsoft.com/en-us/library/bb178027.aspx).
        */
        void SendKeys(wxString& keys, bool wait = false);

        /**
        The SharePoint version number.

        [MSDN documentation for Application.SharePointVersion](http://msdn.microsoft.com/en-us/library/bb257006.aspx).
        */
        long SharePointVersion(const wxString& url);


        /**
        Returns the union of two or more ranges.

        [MSDN documentation for Application.Union](http://msdn.microsoft.com/en-us/library/bb178176.aspx).
        */
        wxExcelRange Union(const wxExcelRangeVector& ranges);

        // ***** PROPERTIES *****

        /**
        Returns a Range Represents the active cell in the active window (the window on top) or in the specified window. If the window is not displaying a worksheet, this property fails.

        [MSDN documentation for Application.ActiveCell](http://msdn.microsoft.com/en-us/library/bb212488.aspx).
        */
        wxExcelRange GetActiveCell();

        /**
        Returns the name of the active printer.

        [MSDN documentation for Application.ActivePrinter](http://msdn.microsoft.com/en-us/library/bb220818.aspx).
        */
        wxString GetActivePrinter();

        /**
        Sets the name of the active printer.

        [MSDN documentation for Application.ActivePrinter](http://msdn.microsoft.com/en-us/library/bb220818.aspx).
        */
        void SetActivePrinter(const wxString& activePrinter);

        /**
        Returns the active sheet (the sheet on top) in the active workbook or in the specified window or workbook.

        [MSDN documentation for Application.ActiveSheet](http://msdn.microsoft.com/en-us/library/bb212495.aspx).
        */
        wxExcelSheet GetActiveSheet();

        /**
        Returns the active window (the window on top).

        [MSDN documentation for Application.ActiveWindow](http://msdn.microsoft.com/en-us/library/bb220819.aspx).
        */
        wxExcelWindow GetActiveWindow();

        /**
        Returns the workbook in the active window (the window on top).

        [MSDN documentation for Application.ActiveWorkbook](http://msdn.microsoft.com/en-us/library/bb220820.aspx).
        */
        wxExcelWorkbook GetActiveWorkbook();

        /**
        True if Microsoft Excel displays a message before overwriting nonblank cells during a drag-and-drop editing operation.

        [MSDN documentation for Application.AlertBeforeOverwriting](http://msdn.microsoft.com/en-us/library/bb220824.aspx).
        */
        bool GetAlertBeforeOverwriting();
        
        /**
        True if Microsoft Excel displays a message before overwriting nonblank cells during a drag-and-drop editing operation.

        [MSDN documentation for Application.AlertBeforeOverwriting](http://msdn.microsoft.com/en-us/library/bb220824.aspx).
        */
        void SetAlertBeforeOverwriting(bool alertBeforeOverwriting);

        /**
        Returns the name of the alternate startup folder.

        [MSDN documentation for Application.AltStartupPath](http://msdn.microsoft.com/en-us/library/bb220837.aspx).
        */
        wxString GetAltStartupPath();
        
        /**
        Sets the name of the alternate startup folder.

        [MSDN documentation for Application.AltStartupPath](http://msdn.microsoft.com/en-us/library/bb220837.aspx).
        */
        void SetAltStartupPath(const wxString& altStartupPath);

        /**
        Returns a Boolean that represents whether to use ClearType to display fonts in the menu, Ribbon, and dialog box text. Since MS Excel 2007.

        [MSDN documentation for Application.AlwaysUseClearType](http://msdn.microsoft.com/en-us/library/bb224791.aspx).
        */
        bool GetAlwaysUseClearType();
        
        /**
        Sets a Boolean that represents whether to use ClearType to display fonts in the menu, Ribbon, and dialog box text. Since MS Excel 2007.

        [MSDN documentation for Application.AlwaysUseClearType](http://msdn.microsoft.com/en-us/library/bb224791.aspx).
        */
        void SetAlwaysUseClearType(bool alwaysUseClearType);

        /**
        Returns a Boolean value that indicates whether the XML features in Microsoft Excel are available.

        [MSDN documentation for Application.ArbitraryXMLSupportAvailable](http://msdn.microsoft.com/en-us/library/bb220844.aspx).
        */
        bool GetArbitraryXMLSupportAvailable();
        
        /**
        True if Microsoft Excel asks the user to update links when opening files with links. False if links are automatically updated with no dialog box.

        [MSDN documentation for Application.AskToUpdateLinks](http://msdn.microsoft.com/en-us/library/bb220846.aspx).
        */
        bool GetAskToUpdateLinks();
        
        /**
        True if Microsoft Excel asks the user to update links when opening files with links. False if links are automatically updated with no dialog box.

        [MSDN documentation for Application.AskToUpdateLinks](http://msdn.microsoft.com/en-us/library/bb220846.aspx).
        */
        void SetAskToUpdateLinks(bool askToUpdateLinks);

        /**
        True (default) if Microsoft Excel automatically formats hyperlinks as you type. False if Excel does not automatically format hyperlinks as you type.

        [MSDN documentation for Application.AutoFormatAsYouTypeReplaceHyperlinks](http://msdn.microsoft.com/en-us/library/bb220853.aspx).
        */
        bool GetAutoFormatAsYouTypeReplaceHyperlinks();
        /**
        True (default) if Microsoft Excel automatically formats hyperlinks as you type. False if Excel does not automatically format hyperlinks as you type.

        [MSDN documentation for Application.AutoFormatAsYouTypeReplaceHyperlinks](http://msdn.microsoft.com/en-us/library/bb220853.aspx).
        */
        void SetAutoFormatAsYouTypeReplaceHyperlinks(bool autoFormatAsYouTypeReplaceHyperlinks);

        /**
        Returns an MsoAutomationSecurity constant that represents the security mode Microsoft Excel uses when programmatically opening files.

        [MSDN documentation for Application.AutomationSecurity](http://msdn.microsoft.com/en-us/library/bb220857.aspx).
        */
        MsoAutomationSecurity GetAutomationSecurity();

        /**
        Sets an MsoAutomationSecurity constant that represents the security mode Microsoft Excel uses when programmatically opening files.

        [MSDN documentation for Application.AutomationSecurity](http://msdn.microsoft.com/en-us/library/bb220857.aspx).
        */
        void SetAutomationSecurity(MsoAutomationSecurity security);

        /**
        True if entries in cells formatted as percentages aren't automatically multiplied by 100 as soon as they are entered.

        [MSDN documentation for Application.AutoPercentEntry](http://msdn.microsoft.com/en-us/library/bb220858.aspx).
        */
        bool GetAutoPercentEntry();
        
        /**
        True if entries in cells formatted as percentages aren't automatically multiplied by 100 as soon as they are entered.

        [MSDN documentation for Application.AutoPercentEntry](http://msdn.microsoft.com/en-us/library/bb220858.aspx).
        */
        void SetAutoPercentEntry(bool autoPercentEntry);

        /**
        Returns the Microsoft Excel build number.

        [MSDN documentation for Application.Build](http://msdn.microsoft.com/en-us/library/bb220895.aspx).
        */
        double GetBuild();

        /**
        True if workbooks are calculated before they're saved to disk (if the Calculation property is set to xlManual). This property is preserved even if you change the Calculation property.

        [MSDN documentation for Application.CalculateBeforeSave](http://msdn.microsoft.com/en-us/library/bb220898.aspx).
        */
        bool GetCalculateBeforeSave();
        
        /**
        True if workbooks are calculated before they're saved to disk (if the Calculation property is set to xlManual). This property is preserved even if you change the Calculation property.

        [MSDN documentation for Application.CalculateBeforeSave](http://msdn.microsoft.com/en-us/library/bb220898.aspx).
        */
        void SetCalculateBeforeSave(bool calculateBeforeSave);

        /**
        Returns an XlCalculation value that represents the calculation mode.

        [MSDN documentation for Application.Calculation](http://msdn.microsoft.com/en-us/library/bb212499.aspx).
        */
        XlCalculation GetCalculation();
        
        /**
        Sets an XlCalculation value that represents the calculation mode. There must be at least one active workbook, else the call fails.

        [MSDN documentation for Application.Calculation](http://msdn.microsoft.com/en-us/library/bb212499.aspx).
        */
        void SetCalculation(XlCalculation calculation);

        /**
        Returns an XlCalculationInterruptKey constant that specifies the key that can interrupt Microsoft Excel when performing calculations.

        [MSDN documentation for Application.CalculationInterruptKey](http://msdn.microsoft.com/en-us/library/bb220900.aspx).
        */
        XlCalculationInterruptKey GetCalculationInterruptKey();
        
        /**
        Sets an XlCalculationInterruptKey constant that specifies the key that can interrupt Microsoft Excel when performing calculations.

        [MSDN documentation for Application.CalculationInterruptKey](http://msdn.microsoft.com/en-us/library/bb220900.aspx).
        */
        void SetCalculationInterruptKey(XlCalculationInterruptKey calculationInterruptKey);

        /**
        Returns an XlCalculationState constant that indicates the calculation state of the application, for any calculations that are being performed in Microsoft Excel.

        [MSDN documentation for Application.CalculationState](http://msdn.microsoft.com/en-us/library/bb220901.aspx).
        */
        XlCalculationState GetCalculationState();

        /**
        Returns a number whose rightmost four digits are the minor calculation engine version number, and whose other digits (on the left) are the major version of Microsoft Excel.

        [MSDN documentation for Application.CalculationVersion](http://msdn.microsoft.com/en-us/library/bb212504.aspx).
        */
        long GetCalculationVersion();

        /**
        Returns a String value that represents the name that appears in the title bar of the main Microsoft Excel window.

        [MSDN documentation for Application.Caption](http://msdn.microsoft.com/en-us/library/bb212508.aspx).
        */
        wxString GetCaption();
        /**
        Sets a String value that represents the name that appears in the title bar of the main Microsoft Excel window.

        [MSDN documentation for Application.Caption](http://msdn.microsoft.com/en-us/library/bb212508.aspx).
        */
        void SetCaption(const wxString& caption);

        /**
        True if dragging and dropping cells is enabled.

        [MSDN documentation for Application.CellDragAndDrop](http://msdn.microsoft.com/en-us/library/bb220911.aspx).
        */
        bool GetCellDragAndDrop();
        /**
        True if dragging and dropping cells is enabled.

        [MSDN documentation for Application.CellDragAndDrop](http://msdn.microsoft.com/en-us/library/bb220911.aspx).
        */
        void SetCellDragAndDrop(bool cellDragAndDrop);

        /**
        Returns the formats that are currently on the Clipboard, as an array of numeric values. To determine whether a particular format is on the Clipboard, compare each element in the array with the appropriate XlClipboardFormat constant.

        [MSDN documentation for Application.ClipboardFormats](http://msdn.microsoft.com/en-us/library/bb177357.aspx).
        */
        wxArrayLong GetClipboardFormats();

        /**
        True if handwriting recognition is limited to numbers and punctuation only. This property is available only if you're using Microsoft Windows for Pen Computing.

        [MSDN documentation for Application.ConstrainNumeric](http://msdn.microsoft.com/en-us/library/bb177388.aspx).
        */
        bool GetConstrainNumeric();
        /**
        True if handwriting recognition is limited to numbers and punctuation only. This property is available only if you're using Microsoft Windows for Pen Computing.

        [MSDN documentation for Application.ConstrainNumeric](http://msdn.microsoft.com/en-us/library/bb177388.aspx).
        */
        void SetConstrainNumeric(bool constrainNumeric);

        /**
        True if Microsoft Excel displays control characters for right-to-left languages.

        [MSDN documentation for Application.ControlCharacters](http://msdn.microsoft.com/en-us/library/bb177392.aspx).
        */
        bool GetControlCharacters();
        /**
        True if Microsoft Excel displays control characters for right-to-left languages. This property can be set only when right to left language support has been installed and selected.

        [MSDN documentation for Application.ControlCharacters](http://msdn.microsoft.com/en-us/library/bb177392.aspx).
        */
        void SetControlCharacters(bool controlCharacters);

        /**
        True if objects are cut, copied, extracted, and sorted with cells.

        [MSDN documentation for Application.CopyObjectsWithCells](http://msdn.microsoft.com/en-us/library/bb177395.aspx).
        */
        bool GetCopyObjectsWithCells();
        /**
        True if objects are cut, copied, extracted, and sorted with cells.

        [MSDN documentation for Application.CopyObjectsWithCells](http://msdn.microsoft.com/en-us/library/bb177395.aspx).
        */
        void SetCopyObjectsWithCells(bool copyObjectsWithCells);

        /**
        Returns the appearance of the mouse pointer in Microsoft Excel.

        [MSDN documentation for Application.Cursor](http://msdn.microsoft.com/en-us/library/bb177420.aspx).
        */
        XlMousePointer GetCursor();
        /**
        Sets the appearance of the mouse pointer in Microsoft Excel.

        [MSDN documentation for Application.Cursor](http://msdn.microsoft.com/en-us/library/bb177420.aspx).
        */
        void SetCursor(XlMousePointer cursor);

        /**
        Returns a value that indicates whether a visual cursor or a logical cursor is used. Can be one of the following constants: xlVisualCursor or xlLogicalCursor.

        [MSDN documentation for Application.CursorMovement](http://msdn.microsoft.com/en-us/library/bb177422.aspx).
        */
        long GetCursorMovement();
        /**
        Returns the number of defined custom lists (including built-in lists).

        [MSDN documentation for Application.CustomListCount](http://msdn.microsoft.com/en-us/library/bb177425.aspx).
        */
        long GetCustomListCount();
        /**
        Returns the status of Cut or Copy mode. Can be 0 ( for false) or an XlCutCopyMode constant.

        [MSDN documentation for Application.CutCopyMode](http://msdn.microsoft.com/en-us/library/bb177432.aspx).
        */
        long GetCutCopyMode();

        /**
        Cancels Cut or Copy mode and removes the moving border.

        [MSDN documentation for Application.CutCopyMode](http://msdn.microsoft.com/en-us/library/bb177432.aspx).
        */
        void SetCutCopyMode();

        /**
        Returns Data Entry mode (xlOn, xlOff or xlStrict). When in Data Entry mode, you can enter data only in the unlocked cells in the currently selected range.

        [MSDN documentation for Application.DataEntryMode](http://msdn.microsoft.com/en-us/library/bb177438.aspx).
        */
        long GetDataEntryMode();
        /**
        Sets Data Entry mode (xlOn, xlOff or xlStrict). When in Data Entry mode, you can enter data only in the unlocked cells in the currently selected range.

        [MSDN documentation for Application.DataEntryMode](http://msdn.microsoft.com/en-us/library/bb177438.aspx).
        */
        void SetDataEntryMode(long dataEntryMode);

        /**
        Returns the application-specific DDE return code that was contained in the last DDE acknowledge message received by Microsoft Excel.

        [MSDN documentation for Application.DDEAppReturnCode](http://msdn.microsoft.com/en-us/library/bb177449.aspx).
        */
        long GetDDEAppReturnCode();

        /**
        Returns the character used for the decimal separator.

        [MSDN documentation for Application.DecimalSeparator](http://msdn.microsoft.com/en-us/library/bb177453.aspx).
        */
        wxString GetDecimalSeparator();
        /**
        Sets the character used for the decimal separator.

        [MSDN documentation for Application.DecimalSeparator](http://msdn.microsoft.com/en-us/library/bb177453.aspx).
        */
        void SetDecimalSeparator(const wxString& decimalSeparator);

        /**
        Returns the default path that Microsoft Excel uses when it opens files.

        [MSDN documentation for Application.DefaultFilePath](http://msdn.microsoft.com/en-us/library/bb177454.aspx).
        */
        wxString GetDefaultFilePath();
        /**
        Sets the default path that Microsoft Excel uses when it opens files.

        [MSDN documentation for Application.DefaultFilePath](http://msdn.microsoft.com/en-us/library/bb177454.aspx).
        */
        void SetDefaultFilePath(const wxString& defaultFilePath);

        /**
        Returns the default format for saving files.

        [MSDN documentation for Application.DefaultSaveFormat](http://msdn.microsoft.com/en-us/library/bb177456.aspx).
        */
        XlFileFormat GetDefaultSaveFormat();
        /**
        Sets the default format for saving files.

        [MSDN documentation for Application.DefaultSaveFormat](http://msdn.microsoft.com/en-us/library/bb177456.aspx).
        */
        void SetDefaultSaveFormat(XlFileFormat defaultSaveFormat);

        /**
        Returns the default direction in which Microsoft Excel displays new windows and worksheets. Can be one of the following constants: xlRTL (right to left) or xlLTR (left to right).

        [MSDN documentation for Application.DefaultSheetDirection](http://msdn.microsoft.com/en-us/library/bb177458.aspx).
        */
        long GetDefaultSheetDirection();
        /**
        Sets the default direction in which Microsoft Excel displays new windows and worksheets. Can be one of the following constants: xlRTL (right to left) or xlLTR (left to right).

        [MSDN documentation for Application.DefaultSheetDirection](http://msdn.microsoft.com/en-us/library/bb177458.aspx).
        */
        void SetDefaultSheetDirection(long defaultSheetDirection);

        /**
        True if Microsoft Excel displays certain alerts and messages.

        [MSDN documentation for Application.DisplayAlerts](http://msdn.microsoft.com/en-us/library/bb177478.aspx).
        */
        bool GetDisplayAlerts();
        /**
        True if Microsoft Excel displays certain alerts and messages.

        [MSDN documentation for Application.DisplayAlerts](http://msdn.microsoft.com/en-us/library/bb177478.aspx).
        */
        void SetDisplayAlerts(bool displayAlerts);

        /**
        Returns the way cells display comments and indicators. Can be one of the XlCommentDisplayMode constants.

        [MSDN documentation for Application.DisplayCommentIndicator](http://msdn.microsoft.com/en-us/library/bb177487.aspx).
        */
        XlCommentDisplayMode GetDisplayCommentIndicator();
        /**
        Sets the way cells display comments and indicators. Can be one of the XlCommentDisplayMode constants.

        [MSDN documentation for Application.DisplayCommentIndicator](http://msdn.microsoft.com/en-us/library/bb177487.aspx).
        */
        void SetDisplayCommentIndicator(XlCommentDisplayMode displayCommentIndicator);

        /**
        Gets or sets whether to show a list of relevant functions and defined names when building cell formulas. Since MS Excel 2007.

        [MSDN documentation for Application.DisplayFormulaAutoComplete](http://msdn.microsoft.com/en-us/library/bb224793.aspx).
        */
        bool GetDisplayFormulaAutoComplete();
        /**
        Gets or sets whether to show a list of relevant functions and defined names when building cell formulas. Since MS Excel 2007.

        [MSDN documentation for Application.DisplayFormulaAutoComplete](http://msdn.microsoft.com/en-us/library/bb224793.aspx).
        */
        void SetDisplayFormulaAutoComplete(bool displayFormulaAutoComplete);

        /**
        True if the formula bar is displayed.

        [MSDN documentation for Application.DisplayFormulaBar](http://msdn.microsoft.com/en-us/library/bb177505.aspx).
        */
        bool GetDisplayFormulaBar();
        /**
        True if the formula bar is to be displayed.

        [MSDN documentation for Application.DisplayFormulaBar](http://msdn.microsoft.com/en-us/library/bb177505.aspx).
        */
        void SetDisplayFormulaBar(bool displayFormulaBar);

        /**
        True if Microsoft Excel is in full-screen mode.

        [MSDN documentation for Application.DisplayFullScreen](http://msdn.microsoft.com/en-us/library/bb177506.aspx).
        */
        bool GetDisplayFullScreen();
        /**
        True if Microsoft Excel should switch to a full-screen mode.

        [MSDN documentation for Application.DisplayFullScreen](http://msdn.microsoft.com/en-us/library/bb177506.aspx).
        */
        void SetDisplayFullScreen(bool displayFullScreen);

        /**
        True if function ToolTips can be displayed.

        [MSDN documentation for Application.DisplayFunctionToolTips](http://msdn.microsoft.com/en-us/library/bb177507.aspx).
        */
        bool GetDisplayFunctionToolTips();
        /**
        True if function ToolTips are to be displayed.

        [MSDN documentation for Application.DisplayFunctionToolTips](http://msdn.microsoft.com/en-us/library/bb177507.aspx).
        */
        void SetDisplayFunctionToolTips(bool displayFunctionToolTips);

        /**
        True if cells containing notes display cell tips and contain note indicators (small dots in their upper-right corners).

        [MSDN documentation for Application.DisplayNoteIndicator](http://msdn.microsoft.com/en-us/library/bb177512.aspx).
        */
        bool GetDisplayNoteIndicator();
        /**
        True if cells containing notes display cell tips and contain note indicators (small dots in their upper-right corners).

        [MSDN documentation for Application.DisplayNoteIndicator](http://msdn.microsoft.com/en-us/library/bb177512.aspx).
        */
        void SetDisplayNoteIndicator(bool displayNoteIndicator);

        /**
        True if scroll bars are visible for all workbooks.

        [MSDN documentation for Application.DisplayScrollBars](http://msdn.microsoft.com/en-us/library/bb220990.aspx).
        */
        bool GetDisplayScrollBars();
        /**
        True if scroll bars are to be visible for all workbooks.

        [MSDN documentation for Application.DisplayScrollBars](http://msdn.microsoft.com/en-us/library/bb220990.aspx).
        */
        void SetDisplayScrollBars(bool displayScrollBars);

        /**
        True if the status bar is displayed.

        [MSDN documentation for Application.DisplayStatusBar](http://msdn.microsoft.com/en-us/library/bb220996.aspx).
        */
        bool GetDisplayStatusBar();
        /**
        True if the status bar is to be displayed.

        [MSDN documentation for Application.DisplayStatusBar](http://msdn.microsoft.com/en-us/library/bb220996.aspx).
        */
        void SetDisplayStatusBar(bool displayStatusBar);

        /**
        True if Microsoft Excel allows editing in cells.

        [MSDN documentation for Application.EditDirectlyInCell](http://msdn.microsoft.com/en-us/library/bb221076.aspx).
        */
        bool GetEditDirectlyInCell();
        /**
        True if Microsoft Excel should allow editing in cells.

        [MSDN documentation for Application.EditDirectlyInCell](http://msdn.microsoft.com/en-us/library/bb221076.aspx).
        */
        void SetEditDirectlyInCell(bool editDirectlyInCell);

        /**
        True if animated insertion and deletion is enabled.

        [MSDN documentation for Application.EnableAnimations](http://msdn.microsoft.com/en-us/library/bb221104.aspx).
        */
        bool GetEnableAnimations();
        /**
        True if animated insertion and deletion is to be enabled.

        [MSDN documentation for Application.EnableAnimations](http://msdn.microsoft.com/en-us/library/bb221104.aspx).
        */
        void SetEnableAnimations(bool enableAnimations);

        /**
        True if the AutoComplete feature is enabled.

        [MSDN documentation for Application.EnableAutoComplete](http://msdn.microsoft.com/en-us/library/bb221107.aspx).
        */
        bool GetEnableAutoComplete();
        /**
        True if the AutoComplete feature is to be enabled.

        [MSDN documentation for Application.EnableAutoComplete](http://msdn.microsoft.com/en-us/library/bb221107.aspx).
        */
        void SetEnableAutoComplete(bool enableAutoComplete);

        /**
        Controls how Microsoft Excel handles CTRL+BREAK (or ESC or COMMAND+PERIOD) user interruptions to the running procedure.

        [MSDN documentation for Application.EnableCancelKey](http://msdn.microsoft.com/en-us/library/bb221139.aspx).
        */
        XlEnableCancelKey GetEnableCancelKey();
        /**
        Controls how Microsoft Excel handles CTRL+BREAK (or ESC or COMMAND+PERIOD) user interruptions to the running procedure.

        [MSDN documentation for Application.EnableCancelKey](http://msdn.microsoft.com/en-us/library/bb221139.aspx).
        */
        void SetEnableCancelKey(XlEnableCancelKey enableCancelKey);

        /**
        Returns a Boolean that represents whether to display an alert message when a user attempts to perform an operation that affects a larger number of cells than is specified in the Office center UI. Since MS Excel 2007.

        [MSDN documentation for Application.EnableLargeOperationAlert](http://msdn.microsoft.com/en-us/library/bb242602.aspx).
        */
        bool GetEnableLargeOperationAlert();
        /**
        Sets a Boolean that represents whether to display an alert message when a user attempts to perform an operation that affects a larger number of cells than is specified in the Office center UI. Since MS Excel 2007.

        [MSDN documentation for Application.EnableLargeOperationAlert](http://msdn.microsoft.com/en-us/library/bb242602.aspx).
        */
        void SetEnableLargeOperationAlert(bool enableLargeOperationAlert);

        /**
        Returns a Boolean that represents whether to show or hide gallery previews that appear when using galleries that support previewing. Setting this property to True shows a preview of your workbook before applying the command. Since MS Excel 2007.

        [MSDN documentation for Application.EnableLivePreview](http://msdn.microsoft.com/en-us/library/bb224794.aspx).
        */
        bool GetEnableLivePreview();
        /**
        Sets a Boolean that represents whether to show or hide gallery previews that appear when using galleries that support previewing. Setting this property to True shows a preview of your workbook before applying the command. Since MS Excel 2007.

        [MSDN documentation for Application.EnableLivePreview](http://msdn.microsoft.com/en-us/library/bb224794.aspx).
        */
        void SetEnableLivePreview(bool enableLivePreview);

        /**
        True if sound is enabled for Microsoft Office.

        [MSDN documentation for Application.EnableSound](http://msdn.microsoft.com/en-us/library/bb221172.aspx).
        */
        bool GetEnableSound();
        /**
        True if sound is enabled for Microsoft Office.

        [MSDN documentation for Application.EnableSound](http://msdn.microsoft.com/en-us/library/bb221172.aspx).
        */
        void SetEnableSound(bool enableSound);

        /**
        Returns an ErrorCheckingOptions object, which represents the error checking options for an application. 

        [MSDN documentation for Application.ErrorCheckingOptions](http://msdn.microsoft.com/en-us/library/bb208474.aspx).
        */
        wxExcelErrorCheckingOptions GetErrorCheckingOptions();

        /**
        True if Microsoft Excel automatically extends formatting and formulas to new data that is added to a list.

        [MSDN documentation for Application.ExtendList](http://msdn.microsoft.com/en-us/library/bb208488.aspx).
        */
        bool GetExtendList();
        /**
        True if Microsoft Excel automatically extends formatting and formulas to new data that is added to a list.

        [MSDN documentation for Application.ExtendList](http://msdn.microsoft.com/en-us/library/bb208488.aspx).
        */
        void SetExtendList(bool extendList);

        /**
        Returns a value (constant) that specifies how Microsoft Excel handles calls to methods and properties that require features that aren't yet installed. Can be one of the MsoFeatureInstall constants listed in the following table.

        [MSDN documentation for Application.FeatureInstall](http://msdn.microsoft.com/en-us/library/bb208494.aspx).
        */
        MsoFeatureInstall GetFeatureInstall();
        /**
        Sets a value (constant) that specifies how Microsoft Excel handles calls to methods and properties that require features that aren't yet installed. Can be one of the MsoFeatureInstall constants listed in the following table.

        [MSDN documentation for Application.FeatureInstall](http://msdn.microsoft.com/en-us/library/bb208494.aspx).
        */
        void SetFeatureInstall(MsoFeatureInstall featureInstall);

        
        /**
        Returns how Excel will validate files before opening them. Since Excel 2010.

        [MSDN documentation for Application.FileValidation](http://msdn.microsoft.com/en-us/library/office/ff822746%28v=office.14%29.aspx).
        */
        MsoFileValidationMode GetFileValidation();

        /**
        Sets how Excel will validate files before opening them. Since Excel 2010.

        [MSDN documentation for Application.FileValidation](http://msdn.microsoft.com/en-us/library/office/ff822746%28v=office.14%29.aspx).
        */
        void SetFileValidation(MsoFileValidationMode fileValidation);

        /**
        All data entered after this property is set to True will be formatted with the number of fixed decimal places set by the FixedDecimalPlaces property.

        [MSDN documentation for Application.FixedDecimal](http://msdn.microsoft.com/en-us/library/bb208517.aspx).
        */
        bool GetFixedDecimal();
        /**
        All data entered after this property is set to True will be formatted with the number of fixed decimal places set by the FixedDecimalPlaces property.

        [MSDN documentation for Application.FixedDecimal](http://msdn.microsoft.com/en-us/library/bb208517.aspx).
        */
        void SetFixedDecimal(bool fixedDecimal);

        /**
        Returns the number of fixed decimal places used when the FixedDecimal property is set to True.

        [MSDN documentation for Application.FixedDecimalPlaces](http://msdn.microsoft.com/en-us/library/bb208519.aspx).
        */
        long GetFixedDecimalPlaces();
        /**
        Sets the number of fixed decimal places used when the FixedDecimal property is set to True.

        [MSDN documentation for Application.FixedDecimalPlaces](http://msdn.microsoft.com/en-us/library/bb208519.aspx).
        */
        void SetFixedDecimalPlaces(long fixedDecimalPlaces);

        /**
        Allows the user to specify the height of the formula bar in lines. Since MS Excel 2007.

        [MSDN documentation for Application.FormulaBarHeight](http://msdn.microsoft.com/en-us/library/bb224795.aspx).
        */
        long GetFormulaBarHeight();
        /**
        Allows the user to specify the height of the formula bar in lines. Since MS Excel 2007.

        [MSDN documentation for Application.FormulaBarHeight](http://msdn.microsoft.com/en-us/library/bb224795.aspx).
        */
        void SetFormulaBarHeight(long formulaBarHeight);

        /**
        Returns True when Microsoft Excel can get PivotTable report data.

        [MSDN documentation for Application.GenerateGetPivotData](http://msdn.microsoft.com/en-us/library/bb208576.aspx).
        */
        bool GetGenerateGetPivotData();
        /**
        Returns True when Microsoft Excel can get PivotTable report data.

        [MSDN documentation for Application.GenerateGetPivotData](http://msdn.microsoft.com/en-us/library/bb208576.aspx).
        */
        void SetGenerateGetPivotData(bool generateGetPivotData);

        /**
        The GenerateTableRefs property determines whether the traditional notation method or the new structured referencing notation method is used for referencing tables in formulas. Since MS Excel 2007.

        [MSDN documentation for Application.GenerateTableRefs](http://msdn.microsoft.com/en-us/library/bb224796.aspx).
        */
        XlGenerateTableRefs GetGenerateTableRefs();
        /**
        The GenerateTableRefs property determines whether the traditional notation method or the new structured referencing notation method is used for referencing tables in formulas. Since MS Excel 2007.

        [MSDN documentation for Application.GenerateTableRefs](http://msdn.microsoft.com/en-us/library/bb224796.aspx).
        */
        void SetGenerateTableRefs(XlGenerateTableRefs generateTableRefs);

        /**
        Returns a Double value that represents the height in points of the main application window.

        [MSDN documentation for Application.Height](http://msdn.microsoft.com/en-us/library/bb179253.aspx).
        */
        double GetHeight();
        /**
        Sets a Double value that represents the height in points of the main application window.

        [MSDN documentation for Application.Height](http://msdn.microsoft.com/en-us/library/bb179253.aspx).
        */
        void SetHeight(double height);

        /**
        Returns whether Excel uses high quality mode to print graphics.

        [MSDN documentation for Application.HighQualityModeForGraphics](http://msdn.microsoft.com/en-us/library/office/ff822842%28v=office.14%29.aspx).
        */
        bool GetHighQualityModeForGraphics();

        /**
        Sets whether Excel uses high quality mode to print graphics.

        [MSDN documentation for Application.HighQualityModeForGraphics](http://msdn.microsoft.com/en-us/library/office/ff822842%28v=office.14%29.aspx).
        */
        void SetHighQualityModeForGraphics(bool highQualityModeForGraphics);


        /**
        Returns the instance Microsoft Excel application.

        [MSDN documentation for Application.Hinstance](http://msdn.microsoft.com/en-us/library/bb177573.aspx).
        */
        long GetHinstance();

        /**
        Returns the instance Microsoft Excel application. Since Excel 2010.

        [MSDN documentation for Application.HinstancePtr](http://msdn.microsoft.com/en-us/library/office/ff841235%28v=office.14%29.aspx).
        */
        LONG_PTR GetHinstancePtr();

        /**
        Returns a Long indicating the top-level window handle of the Microsoft Excel window.

        [MSDN documentation for Application.Hwnd](http://msdn.microsoft.com/en-us/library/bb177579.aspx).
        */
        long GetHwnd();

        /**
        True if remote DDE requests are ignored.

        [MSDN documentation for Application.IgnoreRemoteRequests](http://msdn.microsoft.com/en-us/library/bb177599.aspx).
        */
        bool GetIgnoreRemoteRequests();

        /**
        True if remote DDE requests are ignored.

        [MSDN documentation for Application.IgnoreRemoteRequests](http://msdn.microsoft.com/en-us/library/bb177599.aspx).
        */
        void SetIgnoreRemoteRequests(bool ignoreRemoteRequests);

        /**
        True if Microsoft Excel is in interactive mode.

        [MSDN documentation for Application.Interactive](http://msdn.microsoft.com/en-us/library/bb177670.aspx).
        */
        bool GetInteractive();

        /**
        True to set Microsoft Excel into interactive mode.

        [MSDN documentation for Application.Interactive](http://msdn.microsoft.com/en-us/library/bb177670.aspx).
        */
        void SetInteractive(bool interactive);

        /**
        Returns information about the current country/region and international settings.

        [MSDN documentation for Application.International](http://msdn.microsoft.com/en-us/library/bb177675.aspx).
        */
        wxVariant GetInternational(XlApplicationInternational* index = NULL);

        /**
        True if Microsoft Excel will use iteration to resolve circular references.

        [MSDN documentation for Application.Iteration](http://msdn.microsoft.com/en-us/library/bb177805.aspx).
        */
        bool GetIteration();
        /**
        True if Microsoft Excel will use iteration to resolve circular references.

        [MSDN documentation for Application.Iteration](http://msdn.microsoft.com/en-us/library/bb177805.aspx).
        */
        void SetIteration(bool iteration);


        /**
        Returns true if the specified workbook is open in a Protected View window. Since Excel 2010.

        [MSDN documentation for Application.IsSandboxed](http://msdn.microsoft.com/en-us/library/office/ff839573%28v=office.14%29.aspx).
        */
        bool GetIsSandboxed();

        /**
        Returns the maximum number of cells needed in an operation beyond which an alert is triggered. Since MS Excel 2007.

        [MSDN documentation for Application.LargeOperationCellThousandCount](http://msdn.microsoft.com/en-us/library/bb242605.aspx).
        */
        long GetLargeOperationCellThousandCount();
        /**
        Sets the maximum number of cells needed in an operation beyond which an alert is triggered. Since MS Excel 2007.

        [MSDN documentation for Application.LargeOperationCellThousandCount](http://msdn.microsoft.com/en-us/library/bb242605.aspx).
        */
        void SetLargeOperationCellThousandCount(long largeOperationCellThousandCount);

        /**
        The distance, in points, from the left edge of the screen to the left edge of the main Microsoft Excel window.

        [MSDN documentation for Application.Left](http://msdn.microsoft.com/en-us/library/bb179255.aspx).
        */
        double GetLeft();
        /**
        The distance, in points, from the left edge of the screen to the left edge of the main Microsoft Excel window.

        [MSDN documentation for Application.Left](http://msdn.microsoft.com/en-us/library/bb179255.aspx).
        */
        void SetLeft(double left);

        /**
        Returns the path to the Library folder, but without the final separator.

        [MSDN documentation for Application.LibraryPath](http://msdn.microsoft.com/en-us/library/bb177910.aspx).
        */
        wxString GetLibraryPath();
        /**
        Returns the MAPI mail session number as a hexadecimal string (if there's an active session), or an empty string if there's no session.

        [MSDN documentation for Application.MailSession](http://msdn.microsoft.com/en-us/library/bb208707.aspx).
        */
        wxString GetMailSession();
        /**
        Returns the mail system that's installed on the host machine.

        [MSDN documentation for Application.MailSystem](http://msdn.microsoft.com/en-us/library/bb208709.aspx).
        */
        XlMailSystem GetMailSystem();

        /**
        True if documents formatted for the standard paper size of another country/region (for example, A4) are automatically adjusted so that they're printed correctly on the standard paper size (for example, Letter) of your country/region.

        [MSDN documentation for Application.MapPaperSize](http://msdn.microsoft.com/en-us/library/bb208722.aspx).
        */
        bool GetMapPaperSize();
        /**
        True if documents formatted for the standard paper size of another country/region (for example, A4) are automatically adjusted so that they're printed correctly on the standard paper size (for example, Letter) of your country/region.

        [MSDN documentation for Application.MapPaperSize](http://msdn.microsoft.com/en-us/library/bb208722.aspx).
        */
        void SetMapPaperSize(bool mapPaperSize);

        /**
        True if a math coprocessor is available.

        [MSDN documentation for Application.MathCoprocessorAvailable](http://msdn.microsoft.com/en-us/library/bb208735.aspx).
        */
        bool GetMathCoprocessorAvailable();

        /**
        Returns the maximum amount of change between each iteration as Microsoft Excel resolves circular references.

        [MSDN documentation for Application.MaxChange](http://msdn.microsoft.com/en-us/library/bb208738.aspx).
        */
        double GetMaxChange();
        /**
        Sets the maximum amount of change between each iteration as Microsoft Excel resolves circular references.

        [MSDN documentation for Application.MaxChange](http://msdn.microsoft.com/en-us/library/bb208738.aspx).
        */
        void SetMaxChange(double maxChange);

        /**
        Returns the maximum number of iterations that Microsoft Excel can use to resolve a circular reference.

        [MSDN documentation for Application.MaxIterations](http://msdn.microsoft.com/en-us/library/bb208750.aspx).
        */
        long GetMaxIterations();

        /**
        Sets the maximum number of iterations that Microsoft Excel can use to resolve a circular reference.

        [MSDN documentation for Application.MaxIterations](http://msdn.microsoft.com/en-us/library/bb208750.aspx).
        */
        void SetMaxIterations(long iterations);

        /**
        Specifies the measurement unit used in the application. Since MS Excel 2007.

        [MSDN documentation for Application.MeasurementUnit](http://msdn.microsoft.com/en-us/library/bb224797.aspx).
        */
        XlMeasurementUnits GetMeasurementUnit();
        /**
        Specifies the measurement unit used in the application. Since MS Excel 2007.

        [MSDN documentation for Application.MeasurementUnit](http://msdn.microsoft.com/en-us/library/bb224797.aspx).
        */
        void SetMeasurementUnit(XlMeasurementUnits measurementUnit);

        /**
        True if a mouse is available.

        [MSDN documentation for Application.MouseAvailable](http://msdn.microsoft.com/en-us/library/bb208792.aspx).
        */
        bool GetMouseAvailable();

        /**
        True if the active cell will be moved as soon as the ENTER (RETURN) key is pressed.

        [MSDN documentation for Application.MoveAfterReturn](http://msdn.microsoft.com/en-us/library/bb208798.aspx).
        */
        bool GetMoveAfterReturn();
        /**
        True if the active cell will be moved as soon as the ENTER (RETURN) key is pressed.

        [MSDN documentation for Application.MoveAfterReturn](http://msdn.microsoft.com/en-us/library/bb208798.aspx).
        */
        void SetMoveAfterReturn(bool moveAfterReturn);

        /**
        Returns the direction in which the active cell is moved when the user presses ENTER.

        [MSDN documentation for Application.MoveAfterReturnDirection](http://msdn.microsoft.com/en-us/library/bb208803.aspx).
        */
        XlDirection GetMoveAfterReturnDirection();
        /**
        Sets the direction in which the active cell is moved when the user presses ENTER.

        [MSDN documentation for Application.MoveAfterReturnDirection](http://msdn.microsoft.com/en-us/library/bb208803.aspx).
        */
        void SetMoveAfterReturnDirection(XlDirection moveAfterReturnDirection);

        /**
        Returns the network path where templates are stored. If the network path doesn't exist, this property returns an empty string.

        [MSDN documentation for Application.NetworkTemplatesPath](http://msdn.microsoft.com/en-us/library/bb208825.aspx).
        */
        wxString GetNetworkTemplatesPath();

        /**
        Returns the ODBC query time limit, in seconds. The default value is 45 seconds.

        [MSDN documentation for Application.ODBCTimeout](http://msdn.microsoft.com/en-us/library/bb208848.aspx).
        */
        long GetODBCTimeout();
        /**
        Sets the ODBC query time limit, in seconds. The default value is 45 seconds.

        [MSDN documentation for Application.ODBCTimeout](http://msdn.microsoft.com/en-us/library/bb208848.aspx).
        */
        void SetODBCTimeout(long oDBCTimeout);

        /**
        Returns the name and version number of the current operating system - for example, "Windows (32-bit) 4.00" or "Macintosh 7.00".

        [MSDN documentation for Application.OperatingSystem](http://msdn.microsoft.com/en-us/library/bb208895.aspx).
        */
        wxString GetOperatingSystem();

        /**
        Returns the registered organization name.

        [MSDN documentation for Application.OrganizationName](http://msdn.microsoft.com/en-us/library/bb208901.aspx).
        */
        wxString GetOrganizationName();

        /**
        Returns a String value that represents the complete path to the application, excluding the final separator and name of the application.

        [MSDN documentation for Application.Path](http://msdn.microsoft.com/en-us/library/bb179263.aspx).
        */
        wxString GetPath();

        /**
        Returns the path separator character ("\").

        [MSDN documentation for Application.PathSeparator](http://msdn.microsoft.com/en-us/library/bb221413.aspx).
        */
        wxString GetPathSeparator();
        /**
        True if PivotTable reports use structured selection.

        [MSDN documentation for Application.PivotTableSelection](http://msdn.microsoft.com/en-us/library/bb221427.aspx).
        */
        bool GetPivotTableSelection();
        /**
        True if PivotTable reports use structured selection.

        [MSDN documentation for Application.PivotTableSelection](http://msdn.microsoft.com/en-us/library/bb221427.aspx).
        */
        void SetPivotTableSelection(bool pivotTableSelection);

        /**
        Specifies whether communication with the printer is turned on. Since Excel 2010.

        [MSDN documentation for Application.PrintCommunication](http://msdn.microsoft.com/en-us/library/office/ff835544%28v=office.14%29.aspx).
        */
        bool GetPrintCommunication();

        /**
        Specifies whether communication with the printer is turned on. Since Excel 2010.

        [MSDN documentation for Application.PrintCommunication](http://msdn.microsoft.com/en-us/library/office/ff835544%28v=office.14%29.aspx).
        */
        void SetPrintCommunication(bool printCommunication);

        /**
        Returns the globally unique identifier (GUID) for Microsoft Excel.

        [MSDN documentation for Application.ProductCode](http://msdn.microsoft.com/en-us/library/bb209019.aspx).
        */
        wxString GetProductCode();

        /**
        True if Microsoft Excel asks for summary information when files are first saved.

        [MSDN documentation for Application.PromptForSummaryInfo](http://msdn.microsoft.com/en-us/library/bb209022.aspx).
        */
        bool GetPromptForSummaryInfo();
        /**
        True if Microsoft Excel asks for summary information when files are first saved.

        [MSDN documentation for Application.PromptForSummaryInfo](http://msdn.microsoft.com/en-us/library/bb209022.aspx).
        */
        void SetPromptForSummaryInfo(bool promptForSummaryInfo);

        /**
        Returns True when the Microsoft Excel application is ready

        [MSDN documentation for Application.Ready](http://msdn.microsoft.com/en-us/library/bb209060.aspx).
        */
        bool GetReady();

        /**
        Returns how Microsoft Excel displays cell references and row and column headings in either A1 or R1C1 reference style. 

        [MSDN documentation for Application.ReferenceStyle](http://msdn.microsoft.com/en-us/library/bb209074.aspx).
        */
        XlReferenceStyle GetReferenceStyle();
        /**
        Sets how Microsoft Excel displays cell references and row and column headings in either A1 or R1C1 reference style. 

        [MSDN documentation for Application.ReferenceStyle](http://msdn.microsoft.com/en-us/library/bb209074.aspx).
        */
        void SetReferenceStyle(XlReferenceStyle referenceStyle);

        /**
        True if the IntelliMouse zooms instead of scrolling.

        [MSDN documentation for Application.RollZoom](http://msdn.microsoft.com/en-us/library/bb209180.aspx).
        */
        bool GetRollZoom();
        /**
        True if the IntelliMouse zooms instead of scrolling.

        [MSDN documentation for Application.RollZoom](http://msdn.microsoft.com/en-us/library/bb209180.aspx).
        */
        void SetRollZoom(bool rollZoom);

        /**
        True if screen updating is turned on.

        [MSDN documentation for Application.ScreenUpdating](http://msdn.microsoft.com/en-us/library/bb221605.aspx).
        */
        bool GetScreenUpdating();
        /**
        True if screen updating is turned on.

        [MSDN documentation for Application.ScreenUpdating](http://msdn.microsoft.com/en-us/library/bb221605.aspx).
        */
        void SetScreenUpdating(bool screenUpdating);

        /**
        Returns a Sheets collection that represents all the sheets in the active workbook. 

        [MSDN documentation for Application.Sheets](http://msdn.microsoft.com/en-us/library/bb214193.aspx).
        */
        wxExcelSheets GetSheets();

        /**
        Returns the number of sheets that Microsoft Excel automatically inserts into new workbooks.

        [MSDN documentation for Application.SheetsInNewWorkbook](http://msdn.microsoft.com/en-us/library/bb221669.aspx).
        */
        long GetSheetsInNewWorkbook();
        /**
        Sets the number of sheets that Microsoft Excel automatically inserts into new workbooks.

        [MSDN documentation for Application.SheetsInNewWorkbook](http://msdn.microsoft.com/en-us/library/bb221669.aspx).
        */
        void SetSheetsInNewWorkbook(long sheetsInNewWorkbook);

        /**
        Returns a Boolean that represents whether the Developer tab is displayed in the Ribbon.

        [MSDN documentation for Application.ShowDevTools](http://msdn.microsoft.com/en-us/library/bb224799.aspx). Since MS Excel 2007.
        */
        bool GetShowDevTools();
        /**
        Sets a Boolean that represents whether the Developer tab is displayed in the Ribbon. Since MS Excel 2007.

        [MSDN documentation for Application.ShowDevTools](http://msdn.microsoft.com/en-us/library/bb224799.aspx).
        */
        void SetShowDevTools(bool showDevTools);

        /**
        Returns a Boolean that represents whether to display Mini toolbars when the user right-clicks in the workbook window. Since MS Excel 2007.

        [MSDN documentation for Application.ShowMenuFloaties](http://msdn.microsoft.com/en-us/library/bb224800.aspx).
        */
        bool GetShowMenuFloaties();
        /**
        Sets a Boolean that represents whether to display Mini toolbars when the user right-clicks in the workbook window. Since MS Excel 2007.

        [MSDN documentation for Application.ShowMenuFloaties](http://msdn.microsoft.com/en-us/library/bb224800.aspx).
        */
        void SetShowMenuFloaties(bool showMenuFloaties);

        /**
        Returns a Boolean that represents whether Mini toolbars displays when a user selects text. Since MS Excel 2007.

        [MSDN documentation for Application.ShowSelectionFloaties](http://msdn.microsoft.com/en-us/library/bb224801.aspx).
        */
        bool GetShowSelectionFloaties();
        /**
        Sets a Boolean that represents whether Mini toolbars displays when a user selects text. Since MS Excel 2007.

        [MSDN documentation for Application.ShowSelectionFloaties](http://msdn.microsoft.com/en-us/library/bb224801.aspx).
        */
        void SetShowSelectionFloaties(bool showSelectionFloaties);

        /**
        Returns True (default is False) when the New Workbook task pane appears for a Microsoft Excel application.

        [MSDN documentation for Application.ShowStartupDialog](http://msdn.microsoft.com/en-us/library/bb209233.aspx).
        */
        bool GetShowStartupDialog();
        /**
        Returns True (default is False) when the New Workbook task pane appears for a Microsoft Excel application.

        [MSDN documentation for Application.ShowStartupDialog](http://msdn.microsoft.com/en-us/library/bb209233.aspx).
        */
        void SetShowStartupDialog(bool showStartupDialog);

        /**
        True if ToolTips are turned on.

        [MSDN documentation for Application.ShowToolTips](http://msdn.microsoft.com/en-us/library/bb209235.aspx).
        */
        bool GetShowToolTips();
        /**
        True if ToolTips are turned on.

        [MSDN documentation for Application.ShowToolTips](http://msdn.microsoft.com/en-us/library/bb209235.aspx).
        */
        void SetShowToolTips(bool showToolTips);

        /**
        True if there's a separate Windows taskbar button for each open workbook. The default value is True.

        [MSDN documentation for Application.ShowWindowsInTaskbar](http://msdn.microsoft.com/en-us/library/bb209238.aspx).
        */
        bool GetShowWindowsInTaskbar();
        /**
        True if there's a separate Windows taskbar button for each open workbook. The default value is True.

        [MSDN documentation for Application.ShowWindowsInTaskbar](http://msdn.microsoft.com/en-us/library/bb209238.aspx).
        */
        void SetShowWindowsInTaskbar(bool showWindowsInTaskbar);

        /**
        Returns the name of the standard font.

        [MSDN documentation for Application.StandardFont](http://msdn.microsoft.com/en-us/library/bb209287.aspx).
        */
        wxString GetStandardFont();
        /**
        Sets the name of the standard font. The change doesn't take effect until you restart Microsoft Excel.

        [MSDN documentation for Application.StandardFont](http://msdn.microsoft.com/en-us/library/bb209287.aspx).
        */
        void SetStandardFont(const wxString& standardFont);

        /**
        Returns the standard font size, in points.

        [MSDN documentation for Application.StandardFontSize](http://msdn.microsoft.com/en-us/library/bb209289.aspx).
        */
        long GetStandardFontSize();
        /**
        Sets the standard font size, in points. The change doesn't take effect until you restart Microsoft Excel.

        [MSDN documentation for Application.StandardFontSize](http://msdn.microsoft.com/en-us/library/bb209289.aspx).
        */
        void SetStandardFontSize(long standardFontSize);

        /**
        Returns the complete path of the startup folder, excluding the final separator.

        [MSDN documentation for Application.StartupPath](http://msdn.microsoft.com/en-us/library/bb209298.aspx).
        */
        wxString GetStartupPath();

        /**
        Returns the text in the status bar.

        [MSDN documentation for Application.StatusBar](http://msdn.microsoft.com/en-us/library/bb209301.aspx).
        */
        wxString GetStatusBar();
        //@{
        /**
        Sets the text in the status bar. The second version restores the default status bar text.

        [MSDN documentation for Application.StatusBar](http://msdn.microsoft.com/en-us/library/bb209301.aspx).
        */
        void SetStatusBar(const wxString& statusBar);
        void SetStatusBar();
        //@}

        /**
        Returns the local path where templates are stored.

        [MSDN documentation for Application.TemplatesPath](http://msdn.microsoft.com/en-us/library/bb221720.aspx).
        */
        wxString GetTemplatesPath();

        /**
        Returns the character used for the thousands separator as a String.

        [MSDN documentation for Application.ThousandsSeparator](http://msdn.microsoft.com/en-us/library/bb221814.aspx).
        */
        wxString GetThousandsSeparator();
        /**
        Sets the character used for the thousands separator as a String.

        [MSDN documentation for Application.ThousandsSeparator](http://msdn.microsoft.com/en-us/library/bb221814.aspx).
        */
        void SetThousandsSeparator(const wxString& thousandsSeparator);

        /**
        Returns a Double value that represents the distance, in points, from the top edge of the screen to the top edge of the main Microsoft Excel window.

        [MSDN documentation for Application.Top](http://msdn.microsoft.com/en-us/library/bb214199.aspx).
        */
        double GetTop();
        /**
        Sets a Double value that represents the distance, in points, from the top edge of the screen to the top edge of the main Microsoft Excel window.

        [MSDN documentation for Application.Top](http://msdn.microsoft.com/en-us/library/bb214199.aspx).
        */
        void SetTop(double top);

        /**
        Returns the Microsoft Excel menu or help key, which is usually "/".

        [MSDN documentation for Application.TransitionMenuKey](http://msdn.microsoft.com/en-us/library/bb221897.aspx).
        */
        wxString GetTransitionMenuKey();
        /**
        Sets the Microsoft Excel menu or help key, which is usually "/".

        [MSDN documentation for Application.TransitionMenuKey](http://msdn.microsoft.com/en-us/library/bb221897.aspx).
        */
        void SetTransitionMenuKey(const wxString& transitionMenuKey);

        /**
        Returns the action taken when the Microsoft Excel menu key is pressed. Can be either xlExcelMenus or xlLotusHelp.

        [MSDN documentation for Application.TransitionMenuKeyAction](http://msdn.microsoft.com/en-us/library/bb221902.aspx).
        */
        long GetTransitionMenuKeyAction();
        /**
        Sets the action taken when the Microsoft Excel menu key is pressed. Can be either xlExcelMenus or xlLotusHelp.

        [MSDN documentation for Application.TransitionMenuKeyAction](http://msdn.microsoft.com/en-us/library/bb221902.aspx).
        */
        void SetTransitionMenuKeyAction(long transitionMenuKeyAction);

        /**
        True if transition navigation keys are active.

        [MSDN documentation for Application.TransitionNavigKeys](http://msdn.microsoft.com/en-us/library/bb221905.aspx).
        */
        bool GetTransitionNavigKeys();
        /**
        True if transition navigation keys are active.

        [MSDN documentation for Application.TransitionNavigKeys](http://msdn.microsoft.com/en-us/library/bb221905.aspx).
        */
        void SetTransitionNavigKeys(bool transitionNavigKeys);

        /**
        Returns the maximum height of the space that a window can occupy in the application window area, in points.

        [MSDN documentation for Application.UsableHeight](http://msdn.microsoft.com/en-us/library/bb214391.aspx).
        */
        double GetUsableHeight();

        /**
        Returns the maximum width of the space that a window can occupy in the application window area, in points.

        [MSDN documentation for Application.UsableWidth](http://msdn.microsoft.com/en-us/library/bb214397.aspx).
        */
        double GetUsableWidth();

        /**
        True if the application is visible or if it was created or started by the user. False if you created or started the application programmatically by using the CreateObject or GetObject functions, and the application is hidden.

        [MSDN documentation for Application.UserControl](http://msdn.microsoft.com/en-us/library/bb221971.aspx).
        */
        bool GetUserControl();
        /**
        True if the application is visible or if it was created or started by the user. False if you created or started the application programmatically by using the CreateObject or GetObject functions, and the application is hidden.

        [MSDN documentation for Application.UserControl](http://msdn.microsoft.com/en-us/library/bb221971.aspx).
        */
        void SetUserControl(bool userControl);

        /**
        Returns the path to the location on the user's computer where the COM add-ins are installed.

        [MSDN documentation for Application.UserLibraryPath](http://msdn.microsoft.com/en-us/library/bb221976.aspx).
        */
        wxString GetUserLibraryPath();
        /**
        Returns the name of the current user.

        [MSDN documentation for Application.UserName](http://msdn.microsoft.com/en-us/library/bb221980.aspx).
        */
        wxString GetUserName();
        /**
        Sets the name of the current user.

        [MSDN documentation for Application.UserName](http://msdn.microsoft.com/en-us/library/bb221980.aspx).
        */
        void SetUserName(const wxString& userName);

        /**
        True (default) if the system separators of Microsoft Excel are enabled.

        [MSDN documentation for Application.UseSystemSeparators](http://msdn.microsoft.com/en-us/library/bb221996.aspx).
        */
        bool GetUseSystemSeparators();
        /**
        True (default) if the system separators of Microsoft Excel are enabled.

        [MSDN documentation for Application.UseSystemSeparators](http://msdn.microsoft.com/en-us/library/bb221996.aspx).
        */
        void SetUseSystemSeparators(bool useSystemSeparators);

        /**
        Returns a String value that represents the Microsoft Excel version number. 
        
        Version numbers are as follows:
        Excel 97   =  8
        Excel 2000 =  9
        Excel 2002 = 10
        Excel 2003 = 11
        Excel 2007 = 12
        Excel 2010 = 14
        Excel 2013 = 15

        [MSDN documentation for Application.Version](http://msdn.microsoft.com/en-us/library/bb214414.aspx).

        see also Is2007OrNewer(), Is2010OrNewer() and GetVersionAsDouble()
        */
        wxString GetVersion();

        /**
        Returns a Boolean value that determines whether the object is visible.

        [MSDN documentation for Application.Visible](http://msdn.microsoft.com/en-us/library/bb214421.aspx).
        */
        bool GetVisible();
        /**
        Sets a Boolean value that determines whether the object is visible.

        [MSDN documentation for Application.Visible](http://msdn.microsoft.com/en-us/library/bb214421.aspx).
        */
        void SetVisible(bool visible);

        /**
        Returns a Double value that represents the distance, in points, from the left edge of the application window to its right edge.

        [MSDN documentation for Application.Width](http://msdn.microsoft.com/en-us/library/bb214430.aspx).
        */
        double GetWidth();
        /**
        Sets a Double value that represents the distance, in points, from the left edge of the application window to its right edge.

        [MSDN documentation for Application.Width](http://msdn.microsoft.com/en-us/library/bb214430.aspx).
        */
        void SetWidth(double width);

        /**
        Returns a Windows collection that represents all the windows in all the workbooks. 

        [MSDN documentation for Application.Windows](http://msdn.microsoft.com/en-us/library/bb214435.aspx).
        */
        wxExcelWindows GetWindows();

        /**
        True if the computer is running under Microsoft Windows for Pen Computing.

        [MSDN documentation for Application.WindowsForPens](http://msdn.microsoft.com/en-us/library/bb223061.aspx).
        */
        bool GetWindowsForPens();

        /**
        Returns the state of the window. 

        [MSDN documentation for Application.WindowState](http://msdn.microsoft.com/en-us/library/bb214451.aspx).
        */
        XlWindowState GetWindowState();
        /**
        Sets the state of the window. 

        [MSDN documentation for Application.WindowState](http://msdn.microsoft.com/en-us/library/bb214451.aspx).
        */
        void SetWindowState(XlWindowState windowState);

        /**
        Returns a Workbooks collection that represents all the open workbooks.

        [MSDN documentation for Application.Workbooks](http://msdn.microsoft.com/en-us/library/bb223063.aspx).
        */
        wxExcelWorkbooks GetWorkbooks();

        /**
        For an Application object, returns a Sheets collection that represents all the worksheets in the active workbook. For a Workbook object, returns a Sheets collection that represents all the worksheets in the specified workbook.

        [MSDN documentation for Application.Worksheets](http://msdn.microsoft.com/en-us/library/bb214454.aspx).
        */
        wxExcelWorksheets GetWorksheets();

        /**
            Returns true if the MS Excel is version 2007 or newer, false otherwise. See GetVersion() method.
        */
        bool Is2007OrNewer();

        /**
            Returns true if the MS Excel is version 2010 or newer, false otherwise. See GetVersion() method.
        */
        bool Is2010OrNewer();

        /**
            Returns MS Excel version as a double.
        */
        bool GetVersionAsDouble(double& version);

        /**
        Returns "Application".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("Application"); }
    private:
        bool RangesToVariants(const wxExcelRangeVector& ranges, wxVariantVector& variants);

        

    };



} // namespace wxAutoExcel

#endif // _WXAUTOEXCEL_APPLICATION_H

