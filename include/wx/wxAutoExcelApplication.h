/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// License:     MIT license
/////////////////////////////////////////////////////////////////////////////


#ifndef _WXAUTOEXCEL_APPLICATION_H
#define _WXAUTOEXCEL_APPLICATION_H

#if defined(_WIN64)
    #include <wx/longlong.h>
#endif


#include "wx/wxAutoExcel_defs.h"
#include "wx/wxAutoExcel_enums.h"

#include "wx/wxAutoExcelRange.h"

namespace wxAutoExcel {

#if defined(_WIN64)
    typedef wxLongLong wxAEHinstPtr;
#else
    typedef long wxAEHinstPtr;
#endif

    /**
    @brief Represents Microsoft Excel Application.
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelApplication : public wxExcelRangeOwner
    {
    public:
        /**  Constants for known MS Excel versions used by GetVersionAsEnum_() and IsVersionAtLeast_().
        */
        enum ExcelForWindowsVersions
        {
            evUnknown   = 0,    /*!< Unknown MS Excel version. */
            evExcel2000 = 90,   /*!< MS Excel for Windows 2000 (v9.0). */
            evExcel2002 = 100,  /*!< MS Excel for Windows 2002 (v10.0). */
            evExcel2003 = 110,  /*!< MS Excel for Windows 2003 (v11.0). */
            evExcel2007 = 120,  /*!< MS Excel for Windows 2007 (v12.0). */
            evExcel2010 = 140,  /*!< MS Excel for Windows 2010 (v14.0). */
            evExcel2013 = 150,  /*!< MS Excel for Windows 2013 (v15.0). */
            evExcel2016 = 160,  /*!< MS Excel for Windows 2016 (v16.0). */
            evExcel2016or2019 = evExcel2016,  /*!< MS Excel for Windows 2016 or 2019 (v16.0). */
        };

        /**
        Launches MS Excel and creates an instance of MS Excel Application. If failed then returned wxExcelApplication::IsOk_() returns false.

        */
        static wxExcelApplication CreateInstance();

        /**
        Analogical to wxAutomationObject::GetInstance(). If failed then returned wxExcelApplication::IsOk_() returns false.

        */
        static wxExcelApplication GetInstance(int flags = wxAutomationInstance_CreateIfNeeded);

        /**
        Attempts to connect to the instance of Excel where the workbook with the workbookPath
        (must be the full path) is opened, if any.
        Note: After getting the instance, the workbook with workbookPath may not be the ActiveWorkbook,
        in case the Excel instance has several documents open and some other workbook is active.
        If failed then returned wxExcelApplication::IsOk_() returns false.

        */
        static wxExcelApplication GetInstance(const wxString& workbookPath);

        // ***** METHODS *****

        /**
        Activates a Microsoft application. If the application is already running, this method activates the running application. If the application is not running, this method starts a new instance of the application.

        [MSDN documentation for Application.ActivateMicrosoftApp](https://learn.microsoft.com/en-us/office/vba/api/excel.application.activatemicrosoftapp).
        */
        void ActivateMicrosoftApp(XlMSApplication index);

        //@{
        /**
        Adds a custom list for custom autofill and/or custom sort.

        [MSDN documentation for Application.AddCustomList](https://learn.microsoft.com/en-us/office/vba/api/excel.application.addcustomlist).
        */
        bool AddCustomList(const wxArrayString& listArray);
        bool AddCustomList(wxExcelRange listArray, bool byRow);
        //@}

        /**
        Calculates all open workbooks.

        [MSDN documentation for Application.Calculate](https://learn.microsoft.com/en-us/office/vba/api/excel.application.calculate).
        */
        void Calculate();

        /**
        Forces a full calculation of the data in all open workbooks.

        [MSDN documentation for Application.CalculateFull](https://learn.microsoft.com/en-us/office/vba/api/excel.application.calculatefull).
        */
        void CalculateFull();

        /**
        For all open workbooks, forces a full calculation of the data and rebuilds the dependencies.

        [MSDN documentation for Application.CalculateFullRebuild](https://learn.microsoft.com/en-us/office/vba/api/excel.application.calculatefullrebuild).
        */
        void CalculateFullRebuild();

        /**
        Runs all pending queries to OLEDB and OLAP data sources.

        [MSDN documentation for Application.CalculateUntilAsyncQueriesDone](https://learn.microsoft.com/en-us/office/vba/api/excel.application.calculateuntilasyncqueriesdone).
        */
        void CalculateUntilAsyncQueriesDone();

        /**
        Converts a measurement from centimeters to points (one point equals 0.035 centimeters).

        [MSDN documentation for Application.CentimetersToPoints](https://learn.microsoft.com/en-us/office/vba/api/excel.application.centimeterstopoints).
        */
        double CentimetersToPoints(double centimeters);

        /**
        Stops recalculation. If keepAbort is a valid range, the recalculation will be finished for it.

        [MSDN documentation for Application.CheckAbort](https://learn.microsoft.com/en-us/office/vba/api/excel.application.checkabort).
        */
        void CheckAbort(wxExcelRange* keepAbort = NULL);

        /**
        Checks the spelling of a single word.

        [MSDN documentation for Application.CheckSpelling](https://learn.microsoft.com/en-us/office/vba/api/excel.application.checkspelling).
        */
        bool CheckSpelling(const wxString& word, const wxString& customDictionary = wxEmptyString, wxXlTribool ignoreUpperCase = wxDefaultXlTribool);

        /**
        Converts cell references in a formula between the A1 and R1C1 reference styles, between relative and absolute references, or both.

        [MSDN documentation for Application.ConvertFormula](https://learn.microsoft.com/en-us/office/vba/api/excel.application.convertformula).
        */
        wxString ConvertFormula(const wxString& formula, XlReferenceStyle fromReferenceStyle,
                                XlReferenceStyle* toReferenceStyle = NULL, XlReferenceType* toAbsolute = NULL,
                                wxExcelRange* relativeTo = NULL);

        /**
        Runs a command or performs some other action or actions in another application by way of the specified DDE channel.

        [MSDN documentation for Application.DDEExecute](https://learn.microsoft.com/en-us/office/vba/api/excel.application.ddeexecute).
        */
        void DDEExecute(long channel, const wxString& str);

        /**
        Opens a DDE channel to an application.

        [MSDN documentation for Application.DDEInitiate](https://learn.microsoft.com/en-us/office/vba/api/excel.application.ddeinitiate).
        */
        long DDEInitiate(const wxString& app, const wxString& topic);

        /**
        Sends data to an application.

        [MSDN documentation for Application.DDEPoke](https://learn.microsoft.com/en-us/office/vba/api/excel.application.ddepoke).
        */
        void DDEPoke(long channel, const wxVariant& item, const wxVariant& data);

        /**
        Requests information from the specified application. This method always returns a list.

        [MSDN documentation for Application.DDERequest](https://learn.microsoft.com/en-us/office/vba/api/excel.application.dderequest).
        */
        wxVariant DDERequest(long channel, const wxString& item);

        /**
        Closes a channel to another application.

        [MSDN documentation for Application.DDETerminate](https://learn.microsoft.com/en-us/office/vba/api/excel.application.ddeterminate).
        */
        void DDETerminate(long channel);

        /**
        Deletes a custom list.

        [MSDN documentation for Application.DeleteCustomList](https://learn.microsoft.com/en-us/office/vba/api/excel.application.deletecustomlist).
        */
        bool DeleteCustomList(long listNum);


        /**
        Equivalent to double-clicking the active cell.

        [MSDN documentation for Application.DoubleClick](https://learn.microsoft.com/en-us/office/vba/api/excel.application.doubleclick).
        */
        void DoubleClick();


        /**
        Displays the Open dialog box. Returns true if the file chosen by the user was successfully open, false otherwise.

        [MSDN documentation for Application.FindFile](https://learn.microsoft.com/en-us/office/vba/api/excel.application.findfile).
        */
        bool FindFile();

        /**
        Returns a custom list.

        [MSDN documentation for Application.GetCustomListContents](https://learn.microsoft.com/en-us/office/vba/api/excel.application.getcustomlistcontents).
        */
        wxArrayString GetCustomListContents(long listNum);

        /**
        Returns the custom list number for an array of strings. You can use this method to match both built-in lists and custom-defined lists.

        [MSDN documentation for Application.GetCustomListNum](https://learn.microsoft.com/en-us/office/vba/api/excel.application.getcustomlistnum).
        */
        long GetCustomListNum(const wxArrayString& listArray);

        /**
        Displays the standard Open dialog box and gets a file name from the user without actually opening any files.

        [MSDN documentation for Application.GetOpenFilename](https://learn.microsoft.com/en-us/office/vba/api/excel.application.getopenfilename).
        */
        wxArrayString GetOpenFilename(const wxString& fileFilter = wxEmptyString, long* filterIndex = NULL,
                                      const wxString& title = wxEmptyString, wxXlTribool multiSelect = wxDefaultXlTribool);

        /**
        Returns the Japanese phonetic text of the specified text string. This method is available to you only if you have selected or installed Japanese language support for Microsoft Office.

        [MSDN documentation for Application.GetPhonetic](https://learn.microsoft.com/en-us/office/vba/api/excel.application.getphonetic).
        */
        wxString GetPhonetic(const wxString& text);

        /**
        Displays the standard Save As dialog box and gets a file name from the user without actually saving any files.

        [MSDN documentation for Application.GetSaveAsFilename](https://learn.microsoft.com/en-us/office/vba/api/excel.application.getsaveasfilename).
        */
        wxString GetSaveAsFilename(const wxString& initialFilename = wxEmptyString,
                                   const wxString& fileFilter = wxEmptyString, long* filterIndex = NULL,
                                   const wxString& title = wxEmptyString);
         //@{
        /**
        Selects any range or Visual Basic procedure in any workbook, and activates that workbook if it is not already active.

        [MSDN documentation for Application.Goto](https://learn.microsoft.com/en-us/office/vba/api/excel.application.goto).
        */
        void Goto(const wxString& reference = wxEmptyString, bool scroll = false);
        void Goto(wxExcelRange range, bool scroll = false);
         //@}

        /**
        Displays a Help topic.

        [MSDN documentation for Application.Help](https://learn.microsoft.com/en-us/office/vba/api/excel.application.help).
        */
        void Help(const wxString& helpFile = wxEmptyString, long* helpContextID = NULL);

        /**
        Converts a measurement from inches to points.

        [MSDN documentation for Application.InchesToPoints](https://learn.microsoft.com/en-us/office/vba/api/excel.application.inchestopoints).
        */
        double InchesToPoints(double inches);

        /**
        Returns a Range representing the rectangular intersection of two or more ranges.

        [MSDN documentation for Application.Intersect](https://learn.microsoft.com/en-us/office/vba/api/excel.application.intersect).
        */
        wxExcelRange Intersect(const wxExcelRangeVector& ranges);

        /**
        Closes a MAPI mail session established by Microsoft Excel.

        [MSDN documentation for Application.MailLogoff](https://learn.microsoft.com/en-us/office/vba/api/excel.application.maillogoff).
        */
        void MailLogoff();

        /**
        Logs in to MAPI Mail or Microsoft Exchange and establishes a mail session. If Microsoft Mail is not already running, you must use this method to establish a mail session before mail or document routing functions can be used.

        [MSDN documentation for Application.MailLogon](https://learn.microsoft.com/en-us/office/vba/api/excel.application.maillogon).
        */
        void MailLogon(const wxString& name = wxEmptyString, const wxString& password = wxEmptyString,
                       wxXlTribool downloadNewMail = wxDefaultXlTribool);

        /**
        Quits Microsoft Excel.

        [MSDN documentation for Application.Quit](https://learn.microsoft.com/en-us/office/vba/api/excel.application.quit).
        */
        void Quit();

        /**
        Records code if the macro recorder is on.

        [MSDN documentation for Application.RecordMacro](https://learn.microsoft.com/en-us/office/vba/api/excel.application.recordmacro).
        */
        void RecordMacro(const wxString& basicCode, const wxString& XlmCode = wxEmptyString);

        /**
        Loads an XLL code resource and automatically registers the functions and commands contained in the resource.

        [MSDN documentation for Application.RegisterXLL](https://learn.microsoft.com/en-us/office/vba/api/excel.application.registerxll).
        */
        bool RegisterXLL(const wxString& fileName);


        /**
        Runs a macro.

        [MSDN documentation for Application.Run](https://learn.microsoft.com/en-us/office/vba/api/excel.application.run).
        */
        wxVariant Run(const wxString& macro, const wxVariantVector& arguments);

        /**
        Saves the current workspace.

        [MSDN documentation for Application.SaveWorkspace](https://learn.microsoft.com/en-us/office/vba/api/excel.application.saveworkspace).
        */
        void SaveWorkspace(const wxString& fileName = wxEmptyString);

        /**
        Sends keystrokes to Excel application.

        [MSDN documentation for Application.SendKeys](https://learn.microsoft.com/en-us/office/vba/api/excel.application.sendkeys).
        */
        void SendKeys(const wxString& keys, bool wait = false);

        /**
        The SharePoint version number.

        [MSDN documentation for Application.SharePointVersion](https://learn.microsoft.com/en-us/office/vba/api/excel.application.sharepointversion).
        */
        long SharePointVersion(const wxString& url);


        /**
        Returns the union of two or more ranges.

        This is a way to get a range with several areas without an address, whose separator
        of the areas depends on the Windows regional settings, see GetListSeparator_().

        [MSDN documentation for Application.Union](https://learn.microsoft.com/en-us/office/vba/api/excel.application.union).
        */
        wxExcelRange Union(const wxExcelRangeVector& ranges);

        // ***** PROPERTIES *****

        /**
        Returns a Range Represents the active cell in the active window (the window on top) or in the specified window. If the window is not displaying a worksheet, this property fails.

        [MSDN documentation for Application.ActiveCell](https://learn.microsoft.com/en-us/office/vba/api/excel.application.activecell).
        */
        wxExcelRange GetActiveCell();

        /**
        Returns the name of the active printer.

        [MSDN documentation for Application.ActivePrinter](https://learn.microsoft.com/en-us/office/vba/api/excel.application.activeprinter).
        */
        wxString GetActivePrinter();

        /**
        Sets the name of the active printer.

        [MSDN documentation for Application.ActivePrinter](https://learn.microsoft.com/en-us/office/vba/api/excel.application.activeprinter).
        */
        void SetActivePrinter(const wxString& activePrinter);

        /**
        Returns the active sheet (the sheet on top) in the active workbook or in the specified window or workbook.

        [MSDN documentation for Application.ActiveSheet](https://learn.microsoft.com/en-us/office/vba/api/excel.application.activesheet).
        */
        wxExcelSheet GetActiveSheet();

        /**
        Returns the active window (the window on top).

        [MSDN documentation for Application.ActiveWindow](https://learn.microsoft.com/en-us/office/vba/api/excel.application.activewindow).
        */
        wxExcelWindow GetActiveWindow();

        /**
        Returns the workbook in the active window (the window on top).

        [MSDN documentation for Application.ActiveWorkbook](https://learn.microsoft.com/en-us/office/vba/api/excel.application.activeworkbook).
        */
        wxExcelWorkbook GetActiveWorkbook();

        /**
        Returns an AddIns collection that represents all the add-ins listed in the Add-Ins dialog box (Add-Ins command on the Developer tab).

        [MSDN documentation for Application.AddIns](https://learn.microsoft.com/en-us/office/vba/api/excel.application.addins).
        */
        wxExcelAddIns GetAddIns();

        /**
        Returns an AddIns2 collection that represents all the add-ins that are currently available or open in Microsoft Excel, regardless of whether they are installed

        [MSDN documentation for Application.AddIns2](https://learn.microsoft.com/en-us/office/vba/api/excel.application.addins2).
        */
        wxExcelAddIns2 GetAddIns2();

        /**
        True if Microsoft Excel displays a message before overwriting nonblank cells during a drag-and-drop editing operation.

        [MSDN documentation for Application.AlertBeforeOverwriting](https://learn.microsoft.com/en-us/office/vba/api/excel.application.alertbeforeoverwriting).
        */
        bool GetAlertBeforeOverwriting();

        /**
        True if Microsoft Excel displays a message before overwriting nonblank cells during a drag-and-drop editing operation.

        [MSDN documentation for Application.AlertBeforeOverwriting](https://learn.microsoft.com/en-us/office/vba/api/excel.application.alertbeforeoverwriting).
        */
        void SetAlertBeforeOverwriting(bool alertBeforeOverwriting);

        /**
        Returns the name of the alternate startup folder.

        [MSDN documentation for Application.AltStartupPath](https://learn.microsoft.com/en-us/office/vba/api/excel.application.altstartuppath).
        */
        wxString GetAltStartupPath();

        /**
        Sets the name of the alternate startup folder.

        [MSDN documentation for Application.AltStartupPath](https://learn.microsoft.com/en-us/office/vba/api/excel.application.altstartuppath).
        */
        void SetAltStartupPath(const wxString& altStartupPath);

        /**
        Returns a Boolean that represents whether to use ClearType to display fonts in the menu, Ribbon, and dialog box text. Since MS Excel 2007.

        [MSDN documentation for Application.AlwaysUseClearType](https://learn.microsoft.com/en-us/office/vba/api/excel.application.alwaysusecleartype).
        */
        bool GetAlwaysUseClearType();

        /**
        Sets a Boolean that represents whether to use ClearType to display fonts in the menu, Ribbon, and dialog box text. Since MS Excel 2007.

        [MSDN documentation for Application.AlwaysUseClearType](https://learn.microsoft.com/en-us/office/vba/api/excel.application.alwaysusecleartype).
        */
        void SetAlwaysUseClearType(bool alwaysUseClearType);

        /**
        Returns a Boolean value that indicates whether the XML features in Microsoft Excel are available.

        [MSDN documentation for Application.ArbitraryXMLSupportAvailable](https://learn.microsoft.com/en-us/office/vba/api/excel.application.arbitraryxmlsupportavailable).
        */
        bool GetArbitraryXMLSupportAvailable();

        /**
        True if Microsoft Excel asks the user to update links when opening files with links. False if links are automatically updated with no dialog box.

        [MSDN documentation for Application.AskToUpdateLinks](https://learn.microsoft.com/en-us/office/vba/api/excel.application.asktoupdatelinks).
        */
        bool GetAskToUpdateLinks();

        /**
        True if Microsoft Excel asks the user to update links when opening files with links. False if links are automatically updated with no dialog box.

        [MSDN documentation for Application.AskToUpdateLinks](https://learn.microsoft.com/en-us/office/vba/api/excel.application.asktoupdatelinks).
        */
        void SetAskToUpdateLinks(bool askToUpdateLinks);

        /**
        True (default) if Microsoft Excel automatically formats hyperlinks as you type. False if Excel does not automatically format hyperlinks as you type.

        [MSDN documentation for Application.AutoFormatAsYouTypeReplaceHyperlinks](https://learn.microsoft.com/en-us/office/vba/api/excel.application.autoformatasyoutypereplacehyperlinks).
        */
        bool GetAutoFormatAsYouTypeReplaceHyperlinks();
        /**
        True (default) if Microsoft Excel automatically formats hyperlinks as you type. False if Excel does not automatically format hyperlinks as you type.

        [MSDN documentation for Application.AutoFormatAsYouTypeReplaceHyperlinks](https://learn.microsoft.com/en-us/office/vba/api/excel.application.autoformatasyoutypereplacehyperlinks).
        */
        void SetAutoFormatAsYouTypeReplaceHyperlinks(bool autoFormatAsYouTypeReplaceHyperlinks);

        /**
        Returns an MsoAutomationSecurity constant that represents the security mode Microsoft Excel uses when programmatically opening files.

        [MSDN documentation for Application.AutomationSecurity](https://learn.microsoft.com/en-us/office/vba/api/excel.application.automationsecurity).
        */
        MsoAutomationSecurity GetAutomationSecurity();

        /**
        Sets an MsoAutomationSecurity constant that represents the security mode Microsoft Excel uses when programmatically opening files.

        [MSDN documentation for Application.AutomationSecurity](https://learn.microsoft.com/en-us/office/vba/api/excel.application.automationsecurity).
        */
        void SetAutomationSecurity(MsoAutomationSecurity security);

        /**
        True if entries in cells formatted as percentages aren't automatically multiplied by 100 as soon as they are entered.

        [MSDN documentation for Application.AutoPercentEntry](https://learn.microsoft.com/en-us/office/vba/api/excel.application.autopercententry).
        */
        bool GetAutoPercentEntry();

        /**
        True if entries in cells formatted as percentages aren't automatically multiplied by 100 as soon as they are entered.

        [MSDN documentation for Application.AutoPercentEntry](https://learn.microsoft.com/en-us/office/vba/api/excel.application.autopercententry).
        */
        void SetAutoPercentEntry(bool autoPercentEntry);

        /**
        Returns the Microsoft Excel build number.

        [MSDN documentation for Application.Build](https://learn.microsoft.com/en-us/office/vba/api/excel.application.build).
        */
        double GetBuild();

        /**
        True if workbooks are calculated before they're saved to disk (if the Calculation property is set to xlManual). This property is preserved even if you change the Calculation property.

        [MSDN documentation for Application.CalculateBeforeSave](https://learn.microsoft.com/en-us/office/vba/api/excel.application.calculatebeforesave).
        */
        bool GetCalculateBeforeSave();

        /**
        True if workbooks are calculated before they're saved to disk (if the Calculation property is set to xlManual). This property is preserved even if you change the Calculation property.

        [MSDN documentation for Application.CalculateBeforeSave](https://learn.microsoft.com/en-us/office/vba/api/excel.application.calculatebeforesave).
        */
        void SetCalculateBeforeSave(bool calculateBeforeSave);

        /**
        Returns an XlCalculation value that represents the calculation mode.

        [MSDN documentation for Application.Calculation](https://learn.microsoft.com/en-us/office/vba/api/excel.application.calculation).
        */
        XlCalculation GetCalculation();

        /**
        Sets an XlCalculation value that represents the calculation mode. There must be at least one active workbook, else the call fails.

        [MSDN documentation for Application.Calculation](https://learn.microsoft.com/en-us/office/vba/api/excel.application.calculation).
        */
        void SetCalculation(XlCalculation calculation);

        /**
        Returns an XlCalculationInterruptKey constant that specifies the key that can interrupt Microsoft Excel when performing calculations.

        [MSDN documentation for Application.CalculationInterruptKey](https://learn.microsoft.com/en-us/office/vba/api/excel.application.calculationinterruptkey).
        */
        XlCalculationInterruptKey GetCalculationInterruptKey();

        /**
        Sets an XlCalculationInterruptKey constant that specifies the key that can interrupt Microsoft Excel when performing calculations.

        [MSDN documentation for Application.CalculationInterruptKey](https://learn.microsoft.com/en-us/office/vba/api/excel.application.calculationinterruptkey).
        */
        void SetCalculationInterruptKey(XlCalculationInterruptKey calculationInterruptKey);

        /**
        Returns an XlCalculationState constant that indicates the calculation state of the application, for any calculations that are being performed in Microsoft Excel.

        [MSDN documentation for Application.CalculationState](https://learn.microsoft.com/en-us/office/vba/api/excel.application.calculationstate).
        */
        XlCalculationState GetCalculationState();

        /**
        Returns a number whose rightmost four digits are the minor calculation engine version number, and whose other digits (on the left) are the major version of Microsoft Excel.

        [MSDN documentation for Application.CalculationVersion](https://learn.microsoft.com/en-us/office/vba/api/excel.application.calculationversion).
        */
        long GetCalculationVersion();

        /**
        Returns a String value that represents the name that appears in the title bar of the main Microsoft Excel window.

        [MSDN documentation for Application.Caption](https://learn.microsoft.com/en-us/office/vba/api/excel.application.caption).
        */
        wxString GetCaption();
        /**
        Sets a String value that represents the name that appears in the title bar of the main Microsoft Excel window.

        [MSDN documentation for Application.Caption](https://learn.microsoft.com/en-us/office/vba/api/excel.application.caption).
        */
        void SetCaption(const wxString& caption);

        /**
        True if dragging and dropping cells is enabled.

        [MSDN documentation for Application.CellDragAndDrop](https://learn.microsoft.com/en-us/office/vba/api/excel.application.celldraganddrop).
        */
        bool GetCellDragAndDrop();
        /**
        True if dragging and dropping cells is enabled.

        [MSDN documentation for Application.CellDragAndDrop](https://learn.microsoft.com/en-us/office/vba/api/excel.application.celldraganddrop).
        */
        void SetCellDragAndDrop(bool cellDragAndDrop);

        /**
        Returns the formats that are currently on the Clipboard, as an array of numeric values. To determine whether a particular format is on the Clipboard, compare each element in the array with the appropriate XlClipboardFormat constant.

        [MSDN documentation for Application.ClipboardFormats](https://learn.microsoft.com/en-us/office/vba/api/excel.application.clipboardformats).
        */
        wxArrayLong GetClipboardFormats();

        /**
        True if handwriting recognition is limited to numbers and punctuation only. This property is available only if you're using Microsoft Windows for Pen Computing.

        [MSDN documentation for Application.ConstrainNumeric](https://learn.microsoft.com/en-us/office/vba/api/excel.application.constrainnumeric).
        */
        bool GetConstrainNumeric();
        /**
        True if handwriting recognition is limited to numbers and punctuation only. This property is available only if you're using Microsoft Windows for Pen Computing.

        [MSDN documentation for Application.ConstrainNumeric](https://learn.microsoft.com/en-us/office/vba/api/excel.application.constrainnumeric).
        */
        void SetConstrainNumeric(bool constrainNumeric);

        /**
        True if Microsoft Excel displays control characters for right-to-left languages.

        [MSDN documentation for Application.ControlCharacters](https://learn.microsoft.com/en-us/office/vba/api/excel.application.controlcharacters).
        */
        bool GetControlCharacters();
        /**
        True if Microsoft Excel displays control characters for right-to-left languages. This property can be set only when right to left language support has been installed and selected.

        [MSDN documentation for Application.ControlCharacters](https://learn.microsoft.com/en-us/office/vba/api/excel.application.controlcharacters).
        */
        void SetControlCharacters(bool controlCharacters);

        /**
        True if objects are cut, copied, extracted, and sorted with cells.

        [MSDN documentation for Application.CopyObjectsWithCells](https://learn.microsoft.com/en-us/office/vba/api/excel.application.copyobjectswithcells).
        */
        bool GetCopyObjectsWithCells();
        /**
        True if objects are cut, copied, extracted, and sorted with cells.

        [MSDN documentation for Application.CopyObjectsWithCells](https://learn.microsoft.com/en-us/office/vba/api/excel.application.copyobjectswithcells).
        */
        void SetCopyObjectsWithCells(bool copyObjectsWithCells);

        /**
        Returns the appearance of the mouse pointer in Microsoft Excel.

        [MSDN documentation for Application.Cursor](https://learn.microsoft.com/en-us/office/vba/api/excel.application.cursor).
        */
        XlMousePointer GetCursor();
        /**
        Sets the appearance of the mouse pointer in Microsoft Excel.

        [MSDN documentation for Application.Cursor](https://learn.microsoft.com/en-us/office/vba/api/excel.application.cursor).
        */
        void SetCursor(XlMousePointer cursor);

        /**
        Returns a value that indicates whether a visual cursor or a logical cursor is used. Can be one of the following constants: xlVisualCursor or xlLogicalCursor.

        [MSDN documentation for Application.CursorMovement](https://learn.microsoft.com/en-us/office/vba/api/excel.application.cursormovement).
        */
        long GetCursorMovement();
        /**
        Returns the number of defined custom lists (including built-in lists).

        [MSDN documentation for Application.CustomListCount](https://learn.microsoft.com/en-us/office/vba/api/excel.application.customlistcount).
        */
        long GetCustomListCount();
        /**
        Returns the status of Cut or Copy mode. Can be 0 ( for false) or an XlCutCopyMode constant.

        [MSDN documentation for Application.CutCopyMode](https://learn.microsoft.com/en-us/office/vba/api/excel.application.cutcopymode).
        */
        long GetCutCopyMode();

        /**
        Cancels Cut or Copy mode and removes the moving border.

        [MSDN documentation for Application.CutCopyMode](https://learn.microsoft.com/en-us/office/vba/api/excel.application.cutcopymode).
        */
        void SetCutCopyMode();

        /**
        Returns Data Entry mode (xlOn, xlOff or xlStrict). When in Data Entry mode, you can enter data only in the unlocked cells in the currently selected range.

        [MSDN documentation for Application.DataEntryMode](https://learn.microsoft.com/en-us/office/vba/api/excel.application.dataentrymode).
        */
        long GetDataEntryMode();
        /**
        Sets Data Entry mode (xlOn, xlOff or xlStrict). When in Data Entry mode, you can enter data only in the unlocked cells in the currently selected range.

        [MSDN documentation for Application.DataEntryMode](https://learn.microsoft.com/en-us/office/vba/api/excel.application.dataentrymode).
        */
        void SetDataEntryMode(long dataEntryMode);

        /**
        Returns the application-specific DDE return code that was contained in the last DDE acknowledge message received by Microsoft Excel.

        [MSDN documentation for Application.DDEAppReturnCode](https://learn.microsoft.com/en-us/office/vba/api/excel.application.ddeappreturncode).
        */
        long GetDDEAppReturnCode();

        /**
        Returns the character used for the decimal separator.

        [MSDN documentation for Application.DecimalSeparator](https://learn.microsoft.com/en-us/office/vba/api/excel.application.decimalseparator).
        */
        wxString GetDecimalSeparator();
        /**
        Sets the character used for the decimal separator.

        [MSDN documentation for Application.DecimalSeparator](https://learn.microsoft.com/en-us/office/vba/api/excel.application.decimalseparator).
        */
        void SetDecimalSeparator(const wxString& decimalSeparator);

        /**
        Returns the default path that Microsoft Excel uses when it opens files.

        [MSDN documentation for Application.DefaultFilePath](https://learn.microsoft.com/en-us/office/vba/api/excel.application.defaultfilepath).
        */
        wxString GetDefaultFilePath();
        /**
        Sets the default path that Microsoft Excel uses when it opens files.

        [MSDN documentation for Application.DefaultFilePath](https://learn.microsoft.com/en-us/office/vba/api/excel.application.defaultfilepath).
        */
        void SetDefaultFilePath(const wxString& defaultFilePath);

        /**
        Returns the default format for saving files.

        [MSDN documentation for Application.DefaultSaveFormat](https://learn.microsoft.com/en-us/office/vba/api/excel.application.defaultsaveformat).
        */
        XlFileFormat GetDefaultSaveFormat();
        /**
        Sets the default format for saving files.

        [MSDN documentation for Application.DefaultSaveFormat](https://learn.microsoft.com/en-us/office/vba/api/excel.application.defaultsaveformat).
        */
        void SetDefaultSaveFormat(XlFileFormat defaultSaveFormat);

        /**
        Returns the default direction in which Microsoft Excel displays new windows and worksheets. Can be one of the following constants: xlRTL (right to left) or xlLTR (left to right).

        [MSDN documentation for Application.DefaultSheetDirection](https://learn.microsoft.com/en-us/office/vba/api/excel.application.defaultsheetdirection).
        */
        long GetDefaultSheetDirection();
        /**
        Sets the default direction in which Microsoft Excel displays new windows and worksheets. Can be one of the following constants: xlRTL (right to left) or xlLTR (left to right).

        [MSDN documentation for Application.DefaultSheetDirection](https://learn.microsoft.com/en-us/office/vba/api/excel.application.defaultsheetdirection).
        */
        void SetDefaultSheetDirection(long defaultSheetDirection);

        /**
        True if Microsoft Excel displays certain alerts and messages.

        [MSDN documentation for Application.DisplayAlerts](https://learn.microsoft.com/en-us/office/vba/api/excel.application.displayalerts).
        */
        bool GetDisplayAlerts();
        /**
        True if Microsoft Excel displays certain alerts and messages.

        [MSDN documentation for Application.DisplayAlerts](https://learn.microsoft.com/en-us/office/vba/api/excel.application.displayalerts).
        */
        void SetDisplayAlerts(bool displayAlerts);

        /**
        Returns the way cells display comments and indicators. Can be one of the XlCommentDisplayMode constants.

        [MSDN documentation for Application.DisplayCommentIndicator](https://learn.microsoft.com/en-us/office/vba/api/excel.application.displaycommentindicator).
        */
        XlCommentDisplayMode GetDisplayCommentIndicator();
        /**
        Sets the way cells display comments and indicators. Can be one of the XlCommentDisplayMode constants.

        [MSDN documentation for Application.DisplayCommentIndicator](https://learn.microsoft.com/en-us/office/vba/api/excel.application.displaycommentindicator).
        */
        void SetDisplayCommentIndicator(XlCommentDisplayMode displayCommentIndicator);

        /**
        Gets or sets whether to show a list of relevant functions and defined names when building cell formulas. Since MS Excel 2007.

        [MSDN documentation for Application.DisplayFormulaAutoComplete](https://learn.microsoft.com/en-us/office/vba/api/excel.application.displayformulaautocomplete).
        */
        bool GetDisplayFormulaAutoComplete();
        /**
        Gets or sets whether to show a list of relevant functions and defined names when building cell formulas. Since MS Excel 2007.

        [MSDN documentation for Application.DisplayFormulaAutoComplete](https://learn.microsoft.com/en-us/office/vba/api/excel.application.displayformulaautocomplete).
        */
        void SetDisplayFormulaAutoComplete(bool displayFormulaAutoComplete);

        /**
        True if the formula bar is displayed.

        [MSDN documentation for Application.DisplayFormulaBar](https://learn.microsoft.com/en-us/office/vba/api/excel.application.displayformulabar).
        */
        bool GetDisplayFormulaBar();
        /**
        True if the formula bar is to be displayed.

        [MSDN documentation for Application.DisplayFormulaBar](https://learn.microsoft.com/en-us/office/vba/api/excel.application.displayformulabar).
        */
        void SetDisplayFormulaBar(bool displayFormulaBar);

        /**
        True if Microsoft Excel is in full-screen mode.

        [MSDN documentation for Application.DisplayFullScreen](https://learn.microsoft.com/en-us/office/vba/api/excel.application.displayfullscreen).
        */
        bool GetDisplayFullScreen();
        /**
        True if Microsoft Excel should switch to a full-screen mode.

        [MSDN documentation for Application.DisplayFullScreen](https://learn.microsoft.com/en-us/office/vba/api/excel.application.displayfullscreen).
        */
        void SetDisplayFullScreen(bool displayFullScreen);

        /**
        True if function ToolTips can be displayed.

        [MSDN documentation for Application.DisplayFunctionToolTips](https://learn.microsoft.com/en-us/office/vba/api/excel.application.displayfunctiontooltips).
        */
        bool GetDisplayFunctionToolTips();
        /**
        True if function ToolTips are to be displayed.

        [MSDN documentation for Application.DisplayFunctionToolTips](https://learn.microsoft.com/en-us/office/vba/api/excel.application.displayfunctiontooltips).
        */
        void SetDisplayFunctionToolTips(bool displayFunctionToolTips);

        /**
        True if cells containing notes display cell tips and contain note indicators (small dots in their upper-right corners).

        [MSDN documentation for Application.DisplayNoteIndicator](https://learn.microsoft.com/en-us/office/vba/api/excel.application.displaynoteindicator).
        */
        bool GetDisplayNoteIndicator();
        /**
        True if cells containing notes display cell tips and contain note indicators (small dots in their upper-right corners).

        [MSDN documentation for Application.DisplayNoteIndicator](https://learn.microsoft.com/en-us/office/vba/api/excel.application.displaynoteindicator).
        */
        void SetDisplayNoteIndicator(bool displayNoteIndicator);

        /**
        True if scroll bars are visible for all workbooks.

        [MSDN documentation for Application.DisplayScrollBars](https://learn.microsoft.com/en-us/office/vba/api/excel.application.displayscrollbars).
        */
        bool GetDisplayScrollBars();
        /**
        True if scroll bars are to be visible for all workbooks.

        [MSDN documentation for Application.DisplayScrollBars](https://learn.microsoft.com/en-us/office/vba/api/excel.application.displayscrollbars).
        */
        void SetDisplayScrollBars(bool displayScrollBars);

        /**
        True if the status bar is displayed.

        [MSDN documentation for Application.DisplayStatusBar](https://learn.microsoft.com/en-us/office/vba/api/excel.application.displaystatusbar).
        */
        bool GetDisplayStatusBar();
        /**
        True if the status bar is to be displayed.

        [MSDN documentation for Application.DisplayStatusBar](https://learn.microsoft.com/en-us/office/vba/api/excel.application.displaystatusbar).
        */
        void SetDisplayStatusBar(bool displayStatusBar);

        /**
        True if Microsoft Excel allows editing in cells.

        [MSDN documentation for Application.EditDirectlyInCell](https://learn.microsoft.com/en-us/office/vba/api/excel.application.editdirectlyincell).
        */
        bool GetEditDirectlyInCell();
        /**
        True if Microsoft Excel should allow editing in cells.

        [MSDN documentation for Application.EditDirectlyInCell](https://learn.microsoft.com/en-us/office/vba/api/excel.application.editdirectlyincell).
        */
        void SetEditDirectlyInCell(bool editDirectlyInCell);

        /**
        True if animated insertion and deletion is enabled.

        [MSDN documentation for Application.EnableAnimations](https://learn.microsoft.com/en-us/office/vba/api/excel.application.enableanimations).
        */
        bool GetEnableAnimations();
        /**
        True if animated insertion and deletion is to be enabled.

        [MSDN documentation for Application.EnableAnimations](https://learn.microsoft.com/en-us/office/vba/api/excel.application.enableanimations).
        */
        void SetEnableAnimations(bool enableAnimations);

        /**
        True if the AutoComplete feature is enabled.

        [MSDN documentation for Application.EnableAutoComplete](https://learn.microsoft.com/en-us/office/vba/api/excel.application.enableautocomplete).
        */
        bool GetEnableAutoComplete();
        /**
        True if the AutoComplete feature is to be enabled.

        [MSDN documentation for Application.EnableAutoComplete](https://learn.microsoft.com/en-us/office/vba/api/excel.application.enableautocomplete).
        */
        void SetEnableAutoComplete(bool enableAutoComplete);

        /**
        Controls how Microsoft Excel handles CTRL+BREAK (or ESC or COMMAND+PERIOD) user interruptions to the running procedure.

        [MSDN documentation for Application.EnableCancelKey](https://learn.microsoft.com/en-us/office/vba/api/excel.application.enablecancelkey).
        */
        XlEnableCancelKey GetEnableCancelKey();
        /**
        Controls how Microsoft Excel handles CTRL+BREAK (or ESC or COMMAND+PERIOD) user interruptions to the running procedure.

        [MSDN documentation for Application.EnableCancelKey](https://learn.microsoft.com/en-us/office/vba/api/excel.application.enablecancelkey).
        */
        void SetEnableCancelKey(XlEnableCancelKey enableCancelKey);

        /**
        Returns a Boolean that represents whether to display an alert message when a user attempts to perform an operation that affects a larger number of cells than is specified in the Office center UI. Since MS Excel 2007.

        [MSDN documentation for Application.EnableLargeOperationAlert](https://learn.microsoft.com/en-us/office/vba/api/excel.application.enablelargeoperationalert).
        */
        bool GetEnableLargeOperationAlert();
        /**
        Sets a Boolean that represents whether to display an alert message when a user attempts to perform an operation that affects a larger number of cells than is specified in the Office center UI. Since MS Excel 2007.

        [MSDN documentation for Application.EnableLargeOperationAlert](https://learn.microsoft.com/en-us/office/vba/api/excel.application.enablelargeoperationalert).
        */
        void SetEnableLargeOperationAlert(bool enableLargeOperationAlert);

        /**
        Returns a Boolean that represents whether to show or hide gallery previews that appear when using galleries that support previewing. Setting this property to True shows a preview of your workbook before applying the command. Since MS Excel 2007.

        [MSDN documentation for Application.EnableLivePreview](https://learn.microsoft.com/en-us/office/vba/api/excel.application.enablelivepreview).
        */
        bool GetEnableLivePreview();
        /**
        Sets a Boolean that represents whether to show or hide gallery previews that appear when using galleries that support previewing. Setting this property to True shows a preview of your workbook before applying the command. Since MS Excel 2007.

        [MSDN documentation for Application.EnableLivePreview](https://learn.microsoft.com/en-us/office/vba/api/excel.application.enablelivepreview).
        */
        void SetEnableLivePreview(bool enableLivePreview);

        /**
        Controls whether macro animations are enabled.

        [Excel VBA documentation for Application.EnableMacroAnimations](https://learn.microsoft.com/en-us/office/vba/api/excel.application.enablemacroanimations)
        */
        bool GetEnableMacroAnimations();

        /**
        Controls whether macro animations are enabled.

        [Excel VBA documentation for Application.EnableMacroAnimations](https://learn.microsoft.com/en-us/office/vba/api/excel.application.enablemacroanimations)
        */
        void SetEnableMacroAnimations(bool enableMacroAnimations);

        /**
        True if sound is enabled for Microsoft Office.

        [MSDN documentation for Application.EnableSound](https://learn.microsoft.com/en-us/office/vba/api/excel.application.enablesound).
        */
        bool GetEnableSound();
        /**
        True if sound is enabled for Microsoft Office.

        [MSDN documentation for Application.EnableSound](https://learn.microsoft.com/en-us/office/vba/api/excel.application.enablesound).
        */
        void SetEnableSound(bool enableSound);

        /**
        Returns an ErrorCheckingOptions object, which represents the error checking options for an application.

        [MSDN documentation for Application.ErrorCheckingOptions](https://learn.microsoft.com/en-us/office/vba/api/excel.application.errorcheckingoptions).
        */
        wxExcelErrorCheckingOptions GetErrorCheckingOptions();

        /**
        True if Microsoft Excel automatically extends formatting and formulas to new data that is added to a list.

        [MSDN documentation for Application.ExtendList](https://learn.microsoft.com/en-us/office/vba/api/excel.application.extendlist).
        */
        bool GetExtendList();
        /**
        True if Microsoft Excel automatically extends formatting and formulas to new data that is added to a list.

        [MSDN documentation for Application.ExtendList](https://learn.microsoft.com/en-us/office/vba/api/excel.application.extendlist).
        */
        void SetExtendList(bool extendList);

        /**
        Returns a value (constant) that specifies how Microsoft Excel handles calls to methods and properties that require features that aren't yet installed. Can be one of the MsoFeatureInstall constants listed in the following table.

        [MSDN documentation for Application.FeatureInstall](https://learn.microsoft.com/en-us/office/vba/api/excel.application.featureinstall).
        */
        MsoFeatureInstall GetFeatureInstall();
        /**
        Sets a value (constant) that specifies how Microsoft Excel handles calls to methods and properties that require features that aren't yet installed. Can be one of the MsoFeatureInstall constants listed in the following table.

        [MSDN documentation for Application.FeatureInstall](https://learn.microsoft.com/en-us/office/vba/api/excel.application.featureinstall).
        */
        void SetFeatureInstall(MsoFeatureInstall featureInstall);

        /**
        Returns a FileExportConverters collection that represents all the file converters for saving files available to Microsoft Excel.

        [MSDN documentation for Application.FileExportConverters](https://learn.microsoft.com/en-us/office/vba/api/excel.application.fileexportconverters).
        */
        wxExcelFileExportConverters GetFileExportConverters();


        /**
        Returns how Excel will validate files before opening them. Since Excel 2010.

        [MSDN documentation for Application.FileValidation](https://learn.microsoft.com/en-us/office/vba/api/excel.application.filevalidation).
        */
        MsoFileValidationMode GetFileValidation();

        /**
        Sets how Excel will validate files before opening them. Since Excel 2010.

        [MSDN documentation for Application.FileValidation](https://learn.microsoft.com/en-us/office/vba/api/excel.application.filevalidation).
        */
        void SetFileValidation(MsoFileValidationMode fileValidation);

        /**
        All data entered after this property is set to True will be formatted with the number of fixed decimal places set by the FixedDecimalPlaces property.

        [MSDN documentation for Application.FixedDecimal](https://learn.microsoft.com/en-us/office/vba/api/excel.application.fixeddecimal).
        */
        bool GetFixedDecimal();
        /**
        All data entered after this property is set to True will be formatted with the number of fixed decimal places set by the FixedDecimalPlaces property.

        [MSDN documentation for Application.FixedDecimal](https://learn.microsoft.com/en-us/office/vba/api/excel.application.fixeddecimal).
        */
        void SetFixedDecimal(bool fixedDecimal);

        /**
        Returns the number of fixed decimal places used when the FixedDecimal property is set to True.

        [MSDN documentation for Application.FixedDecimalPlaces](https://learn.microsoft.com/en-us/office/vba/api/excel.application.fixeddecimalplaces).
        */
        long GetFixedDecimalPlaces();
        /**
        Sets the number of fixed decimal places used when the FixedDecimal property is set to True.

        [MSDN documentation for Application.FixedDecimalPlaces](https://learn.microsoft.com/en-us/office/vba/api/excel.application.fixeddecimalplaces).
        */
        void SetFixedDecimalPlaces(long fixedDecimalPlaces);

         /**
        Controls whether Flash Fill feature has been enabled and active.

        [Excel VBA documentation for Application.FlashFill](https://learn.microsoft.com/en-us/office/vba/api/excel.application.flashfill)
        */
        bool GetFlashFill();

        /**
        Controls whether Flash Fill feature has been enabled and active.

        [Excel VBA documentation for Application.FlashFill](https://learn.microsoft.com/en-us/office/vba/api/excel.application.flashfill)
        */
        void SetFlashFill(bool flashFill);

        /**
        Controls whether Flash Fill feature is enabled

        [Excel VBA documentation for Application.FlashFillMode](https://learn.microsoft.com/en-us/office/vba/api/excel.application.flashfillmode)
        */
        bool GetFlashFillMode();

        /**
        Controls whether Flash Fill feature is enabled

        [Excel VBA documentation for Application.FlashFillMode](https://learn.microsoft.com/en-us/office/vba/api/excel.application.flashfillmode)
        */
        void SetFlashFillMode(bool flashFillMode);

        /**
        Allows the user to specify the height of the formula bar in lines. Since MS Excel 2007.

        [MSDN documentation for Application.FormulaBarHeight](https://learn.microsoft.com/en-us/office/vba/api/excel.application.formulabarheight).
        */
        long GetFormulaBarHeight();
        /**
        Allows the user to specify the height of the formula bar in lines. Since MS Excel 2007.

        [MSDN documentation for Application.FormulaBarHeight](https://learn.microsoft.com/en-us/office/vba/api/excel.application.formulabarheight).
        */
        void SetFormulaBarHeight(long formulaBarHeight);

        /**
        Returns True when Microsoft Excel can get PivotTable report data.

        [MSDN documentation for Application.GenerateGetPivotData](https://learn.microsoft.com/en-us/office/vba/api/excel.application.generategetpivotdata).
        */
        bool GetGenerateGetPivotData();
        /**
        Returns True when Microsoft Excel can get PivotTable report data.

        [MSDN documentation for Application.GenerateGetPivotData](https://learn.microsoft.com/en-us/office/vba/api/excel.application.generategetpivotdata).
        */
        void SetGenerateGetPivotData(bool generateGetPivotData);

        /**
        The GenerateTableRefs property determines whether the traditional notation method or the new structured referencing notation method is used for referencing tables in formulas. Since MS Excel 2007.

        [MSDN documentation for Application.GenerateTableRefs](https://learn.microsoft.com/en-us/office/vba/api/excel.application.generatetablerefs).
        */
        XlGenerateTableRefs GetGenerateTableRefs();
        /**
        The GenerateTableRefs property determines whether the traditional notation method or the new structured referencing notation method is used for referencing tables in formulas. Since MS Excel 2007.

        [MSDN documentation for Application.GenerateTableRefs](https://learn.microsoft.com/en-us/office/vba/api/excel.application.generatetablerefs).
        */
        void SetGenerateTableRefs(XlGenerateTableRefs generateTableRefs);

        /**
        Returns a Double value that represents the height in points of the main application window.

        [MSDN documentation for Application.Height](https://learn.microsoft.com/en-us/office/vba/api/excel.application.height).
        */
        double GetHeight();
        /**
        Sets a Double value that represents the height in points of the main application window.

        [MSDN documentation for Application.Height](https://learn.microsoft.com/en-us/office/vba/api/excel.application.height).
        */
        void SetHeight(double height);

        /**
        Returns whether Excel uses high quality mode to print graphics.

        [MSDN documentation for Application.HighQualityModeForGraphics](https://learn.microsoft.com/en-us/office/vba/api/excel.application.highqualitymodeforgraphics).
        */
        bool GetHighQualityModeForGraphics();

        /**
        Sets whether Excel uses high quality mode to print graphics.

        [MSDN documentation for Application.HighQualityModeForGraphics](https://learn.microsoft.com/en-us/office/vba/api/excel.application.highqualitymodeforgraphics).
        */
        void SetHighQualityModeForGraphics(bool highQualityModeForGraphics);


        /**
        Returns the instance Microsoft Excel application.

        [MSDN documentation for Application.Hinstance](https://learn.microsoft.com/en-us/office/vba/api/excel.application.hinstance).
        */
        long GetHinstance();

        /**
        Returns the instance Microsoft Excel application. Since Excel 2010.

        [MSDN documentation for Application.HinstancePtr](https://learn.microsoft.com/en-us/office/vba/api/excel.application.hinstanceptr).
        */
        wxAEHinstPtr GetHinstancePtr();

        /**
        Returns a Long indicating the top-level window handle of the Microsoft Excel window.

        [MSDN documentation for Application.Hwnd](https://learn.microsoft.com/en-us/office/vba/api/excel.application.hwnd).
        */
        long GetHwnd();

        /**
        True if remote DDE requests are ignored.

        [MSDN documentation for Application.IgnoreRemoteRequests](https://learn.microsoft.com/en-us/office/vba/api/excel.application.ignoreremoterequests).
        */
        bool GetIgnoreRemoteRequests();

        /**
        True if remote DDE requests are ignored.

        [MSDN documentation for Application.IgnoreRemoteRequests](https://learn.microsoft.com/en-us/office/vba/api/excel.application.ignoreremoterequests).
        */
        void SetIgnoreRemoteRequests(bool ignoreRemoteRequests);

        /**
        True if Microsoft Excel is in interactive mode.

        [MSDN documentation for Application.Interactive](https://learn.microsoft.com/en-us/office/vba/api/excel.application.interactive).
        */
        bool GetInteractive();

        /**
        True to set Microsoft Excel into interactive mode.

        [MSDN documentation for Application.Interactive](https://learn.microsoft.com/en-us/office/vba/api/excel.application.interactive).
        */
        void SetInteractive(bool interactive);

        /**
        Returns information about the current country/region and international settings.

        When called without a parameter, all the values are returned as a list.

        [MSDN documentation for Application.International](https://learn.microsoft.com/en-us/office/vba/api/excel.application.international).
        */
        wxVariant GetInternational(XlApplicationInternational* index = NULL);

        /**
        True if Microsoft Excel will use iteration to resolve circular references.

        [MSDN documentation for Application.Iteration](https://learn.microsoft.com/en-us/office/vba/api/excel.application.iteration).
        */
        bool GetIteration();
        /**
        True if Microsoft Excel will use iteration to resolve circular references.

        [MSDN documentation for Application.Iteration](https://learn.microsoft.com/en-us/office/vba/api/excel.application.iteration).
        */
        void SetIteration(bool iteration);

        /**
        Returns true if the specified workbook is open in a Protected View window. Since Excel 2010.

        [MSDN documentation for Application.IsSandboxed](https://learn.microsoft.com/en-us/office/vba/api/excel.application.issandboxed).
        */
        bool GetIsSandboxed();

        /**
        Returns the LanguageSettings object, which contains information about the language settings in Microsoft Excel.

        [MSDN documentation for Application.LanguageSettings](https://learn.microsoft.com/en-us/office/vba/api/excel.application.languagesettings).
        */
        wxExcelLanguageSettings GetLanguageSettings();

        /**
        Returns the maximum number of cells needed in an operation beyond which an alert is triggered. Since MS Excel 2007.

        [MSDN documentation for Application.LargeOperationCellThousandCount](https://learn.microsoft.com/en-us/office/vba/api/excel.application.largeoperationcellthousandcount).
        */
        long GetLargeOperationCellThousandCount();
        /**
        Sets the maximum number of cells needed in an operation beyond which an alert is triggered. Since MS Excel 2007.

        [MSDN documentation for Application.LargeOperationCellThousandCount](https://learn.microsoft.com/en-us/office/vba/api/excel.application.largeoperationcellthousandcount).
        */
        void SetLargeOperationCellThousandCount(long largeOperationCellThousandCount);

        /**
        The distance, in points, from the left edge of the screen to the left edge of the main Microsoft Excel window.

        [MSDN documentation for Application.Left](https://learn.microsoft.com/en-us/office/vba/api/excel.application.left).
        */
        double GetLeft();
        /**
        The distance, in points, from the left edge of the screen to the left edge of the main Microsoft Excel window.

        [MSDN documentation for Application.Left](https://learn.microsoft.com/en-us/office/vba/api/excel.application.left).
        */
        void SetLeft(double left);

        /**
        Returns the path to the Library folder, but without the final separator.

        [MSDN documentation for Application.LibraryPath](https://learn.microsoft.com/en-us/office/vba/api/excel.application.librarypath).
        */
        wxString GetLibraryPath();
        /**
        Returns the MAPI mail session number as a hexadecimal string (if there's an active session), or an empty string if there's no session.

        [MSDN documentation for Application.MailSession](https://learn.microsoft.com/en-us/office/vba/api/excel.application.mailsession).
        */
        wxString GetMailSession();
        /**
        Returns the mail system that's installed on the host machine.

        [MSDN documentation for Application.MailSystem](https://learn.microsoft.com/en-us/office/vba/api/excel.application.mailsystem).
        */
        XlMailSystem GetMailSystem();

        /**
        True if documents formatted for the standard paper size of another country/region (for example, A4) are automatically adjusted so that they're printed correctly on the standard paper size (for example, Letter) of your country/region.

        [MSDN documentation for Application.MapPaperSize](https://learn.microsoft.com/en-us/office/vba/api/excel.application.mappapersize).
        */
        bool GetMapPaperSize();
        /**
        True if documents formatted for the standard paper size of another country/region (for example, A4) are automatically adjusted so that they're printed correctly on the standard paper size (for example, Letter) of your country/region.

        [MSDN documentation for Application.MapPaperSize](https://learn.microsoft.com/en-us/office/vba/api/excel.application.mappapersize).
        */
        void SetMapPaperSize(bool mapPaperSize);

        /**
        True if a math coprocessor is available.

        [MSDN documentation for Application.MathCoprocessorAvailable](https://learn.microsoft.com/en-us/office/vba/api/excel.application.mathcoprocessoravailable).
        */
        bool GetMathCoprocessorAvailable();

        /**
        Returns the maximum amount of change between each iteration as Microsoft Excel resolves circular references.

        [MSDN documentation for Application.MaxChange](https://learn.microsoft.com/en-us/office/vba/api/excel.application.maxchange).
        */
        double GetMaxChange();
        /**
        Sets the maximum amount of change between each iteration as Microsoft Excel resolves circular references.

        [MSDN documentation for Application.MaxChange](https://learn.microsoft.com/en-us/office/vba/api/excel.application.maxchange).
        */
        void SetMaxChange(double maxChange);

        /**
        Returns the maximum number of iterations that Microsoft Excel can use to resolve a circular reference.

        [MSDN documentation for Application.MaxIterations](https://learn.microsoft.com/en-us/office/vba/api/excel.application.maxiterations).
        */
        long GetMaxIterations();

        /**
        Sets the maximum number of iterations that Microsoft Excel can use to resolve a circular reference.

        [MSDN documentation for Application.MaxIterations](https://learn.microsoft.com/en-us/office/vba/api/excel.application.maxiterations).
        */
        void SetMaxIterations(long iterations);

        /**
        Controls whether to merge multiple instances of the application into a single one.

        [Excel VBA documentation for Application.MergeInstances](https://learn.microsoft.com/en-us/office/vba/api/excel.application.mergeinstances)
        */
        bool GetMergeInstances();

        /**
        Controls whether to merge multiple instances of the application into a single one.

        [Excel VBA documentation for Application.MergeInstances](https://learn.microsoft.com/en-us/office/vba/api/excel.application.mergeinstances)
        */
        void SetMergeInstances(bool mergeInstances);

        /**
        Specifies the measurement unit used in the application. Since MS Excel 2007.

        [MSDN documentation for Application.MeasurementUnit](https://learn.microsoft.com/en-us/office/vba/api/excel.application.measurementunit).
        */
        XlMeasurementUnits GetMeasurementUnit();
        /**
        Specifies the measurement unit used in the application. Since MS Excel 2007.

        [MSDN documentation for Application.MeasurementUnit](https://learn.microsoft.com/en-us/office/vba/api/excel.application.measurementunit).
        */
        void SetMeasurementUnit(XlMeasurementUnits measurementUnit);

        /**
        True if a mouse is available.

        [MSDN documentation for Application.MouseAvailable](https://learn.microsoft.com/en-us/office/vba/api/excel.application.mouseavailable).
        */
        bool GetMouseAvailable();

        /**
        True if the active cell will be moved as soon as the ENTER (RETURN) key is pressed.

        [MSDN documentation for Application.MoveAfterReturn](https://learn.microsoft.com/en-us/office/vba/api/excel.application.moveafterreturn).
        */
        bool GetMoveAfterReturn();
        /**
        True if the active cell will be moved as soon as the ENTER (RETURN) key is pressed.

        [MSDN documentation for Application.MoveAfterReturn](https://learn.microsoft.com/en-us/office/vba/api/excel.application.moveafterreturn).
        */
        void SetMoveAfterReturn(bool moveAfterReturn);

        /**
        Returns the direction in which the active cell is moved when the user presses ENTER.

        [MSDN documentation for Application.MoveAfterReturnDirection](https://learn.microsoft.com/en-us/office/vba/api/excel.application.moveafterreturndirection).
        */
        XlDirection GetMoveAfterReturnDirection();
        /**
        Sets the direction in which the active cell is moved when the user presses ENTER.

        [MSDN documentation for Application.MoveAfterReturnDirection](https://learn.microsoft.com/en-us/office/vba/api/excel.application.moveafterreturndirection).
        */
        void SetMoveAfterReturnDirection(XlDirection moveAfterReturnDirection);

        /**
        Returns a MultiThreadedCalculation object that controls the multi-threaded recalculation settings.

        [MSDN documentation for Application.MultiThreadedCalculation](https://learn.microsoft.com/en-us/office/vba/api/excel.application.multithreadedcalculation)
        */
        wxExcelMultiThreadedCalculation GetMultiThreadedCalculation();

        /**
        Returns the name of the object.

        [MSDN documentation for Application.Name](https://learn.microsoft.com/en-us/office/vba/api/excel.application.name).
        */
        wxString GetName();

        /**
        Returns a Names collection that represents all the names in the active workbook.

        [MSDN documentation for Application.Names](https://learn.microsoft.com/en-us/office/vba/api/excel.application.names).
        */
        wxExcelNames GetNames();

        /**
        Returns the network path where templates are stored. If the network path doesn't exist, this property returns an empty string.

        [MSDN documentation for Application.NetworkTemplatesPath](https://learn.microsoft.com/en-us/office/vba/api/excel.application.networktemplatespath).
        */
        wxString GetNetworkTemplatesPath();

        /**
        Returns the ODBC query time limit, in seconds. The default value is 45 seconds.

        [MSDN documentation for Application.ODBCTimeout](https://learn.microsoft.com/en-us/office/vba/api/excel.application.odbctimeout).
        */
        long GetODBCTimeout();
        /**
        Sets the ODBC query time limit, in seconds. The default value is 45 seconds.

        [MSDN documentation for Application.ODBCTimeout](https://learn.microsoft.com/en-us/office/vba/api/excel.application.odbctimeout).
        */
        void SetODBCTimeout(long oDBCTimeout);

        /**
        Returns the name and version number of the current operating system - for example, "Windows (32-bit) 4.00" or "Macintosh 7.00".

        [MSDN documentation for Application.OperatingSystem](https://learn.microsoft.com/en-us/office/vba/api/excel.application.operatingsystem).
        */
        wxString GetOperatingSystem();

        /**
        Returns the registered organization name.

        [MSDN documentation for Application.OrganizationName](https://learn.microsoft.com/en-us/office/vba/api/excel.application.organizationname).
        */
        wxString GetOrganizationName();

        /**
        Returns a String value that represents the complete path to the application, excluding the final separator and name of the application.

        [MSDN documentation for Application.Path](https://learn.microsoft.com/en-us/office/vba/api/excel.application.path).
        */
        wxString GetPath();

        /**
        Returns the path separator character ("\").

        [MSDN documentation for Application.PathSeparator](https://learn.microsoft.com/en-us/office/vba/api/excel.application.pathseparator).
        */
        wxString GetPathSeparator();
        /**
        True if PivotTable reports use structured selection.

        [MSDN documentation for Application.PivotTableSelection](https://learn.microsoft.com/en-us/office/vba/api/excel.application.pivottableselection).
        */
        bool GetPivotTableSelection();
        /**
        True if PivotTable reports use structured selection.

        [MSDN documentation for Application.PivotTableSelection](https://learn.microsoft.com/en-us/office/vba/api/excel.application.pivottableselection).
        */
        void SetPivotTableSelection(bool pivotTableSelection);

        /**
        Specifies whether communication with the printer is turned on. Since Excel 2010.

        [MSDN documentation for Application.PrintCommunication](https://learn.microsoft.com/en-us/office/vba/api/excel.application.printcommunication).
        */
        bool GetPrintCommunication();

        /**
        Specifies whether communication with the printer is turned on. Since Excel 2010.

        [MSDN documentation for Application.PrintCommunication](https://learn.microsoft.com/en-us/office/vba/api/excel.application.printcommunication).
        */
        void SetPrintCommunication(bool printCommunication);

        /**
        Returns the globally unique identifier (GUID) for Microsoft Excel.

        [MSDN documentation for Application.ProductCode](https://learn.microsoft.com/en-us/office/vba/api/excel.application.productcode).
        */
        wxString GetProductCode();

        /**
        True if Microsoft Excel asks for summary information when files are first saved.

        [MSDN documentation for Application.PromptForSummaryInfo](https://learn.microsoft.com/en-us/office/vba/api/excel.application.promptforsummaryinfo).
        */
        bool GetPromptForSummaryInfo();
        /**
        True if Microsoft Excel asks for summary information when files are first saved.

        [MSDN documentation for Application.PromptForSummaryInfo](https://learn.microsoft.com/en-us/office/vba/api/excel.application.promptforsummaryinfo).
        */
        void SetPromptForSummaryInfo(bool promptForSummaryInfo);

        /**
        Returns True when the Microsoft Excel application is ready

        [MSDN documentation for Application.Ready](https://learn.microsoft.com/en-us/office/vba/api/excel.application.ready).
        */
        bool GetReady();

        /**
        Returns a RecentFiles collection that represents the list of recently used files.

        [MSDN documentation for Application.RecentFiles](https://learn.microsoft.com/en-us/office/vba/api/excel.application.recentfiles).
        */
        wxExcelRecentFiles GetRecentFiles();

        /**
        Returns how Microsoft Excel displays cell references and row and column headings in either A1 or R1C1 reference style.

        [MSDN documentation for Application.ReferenceStyle](https://learn.microsoft.com/en-us/office/vba/api/excel.application.referencestyle).
        */
        XlReferenceStyle GetReferenceStyle();
        /**
        Sets how Microsoft Excel displays cell references and row and column headings in either A1 or R1C1 reference style.

        [MSDN documentation for Application.ReferenceStyle](https://learn.microsoft.com/en-us/office/vba/api/excel.application.referencestyle).
        */
        void SetReferenceStyle(XlReferenceStyle referenceStyle);

        /**
        True if the IntelliMouse zooms instead of scrolling.

        [MSDN documentation for Application.RollZoom](https://learn.microsoft.com/en-us/office/vba/api/excel.application.rollzoom).
        */
        bool GetRollZoom();
        /**
        True if the IntelliMouse zooms instead of scrolling.

        [MSDN documentation for Application.RollZoom](https://learn.microsoft.com/en-us/office/vba/api/excel.application.rollzoom).
        */
        void SetRollZoom(bool rollZoom);

        /**
        True if screen updating is turned on.

        [MSDN documentation for Application.ScreenUpdating](https://learn.microsoft.com/en-us/office/vba/api/excel.application.screenupdating).
        */
        bool GetScreenUpdating();
        /**
        True if screen updating is turned on.

        [MSDN documentation for Application.ScreenUpdating](https://learn.microsoft.com/en-us/office/vba/api/excel.application.screenupdating).
        */
        void SetScreenUpdating(bool screenUpdating);

        /**
        Returns a Sheets collection that represents all the sheets in the active workbook.

        [MSDN documentation for Application.Sheets](https://learn.microsoft.com/en-us/office/vba/api/excel.application.sheets).
        */
        wxExcelSheets GetSheets();

        /**
        Returns the number of sheets that Microsoft Excel automatically inserts into new workbooks.

        [MSDN documentation for Application.SheetsInNewWorkbook](https://learn.microsoft.com/en-us/office/vba/api/excel.application.sheetsinnewworkbook).
        */
        long GetSheetsInNewWorkbook();
        /**
        Sets the number of sheets that Microsoft Excel automatically inserts into new workbooks.

        [MSDN documentation for Application.SheetsInNewWorkbook](https://learn.microsoft.com/en-us/office/vba/api/excel.application.sheetsinnewworkbook).
        */
        void SetSheetsInNewWorkbook(long sheetsInNewWorkbook);

        /**
        Returns a Boolean that represents whether the Developer tab is displayed in the Ribbon.

        [MSDN documentation for Application.ShowDevTools](https://learn.microsoft.com/en-us/office/vba/api/excel.application.showdevtools). Since MS Excel 2007.
        */
        bool GetShowDevTools();
        /**
        Sets a Boolean that represents whether the Developer tab is displayed in the Ribbon. Since MS Excel 2007.

        [MSDN documentation for Application.ShowDevTools](https://learn.microsoft.com/en-us/office/vba/api/excel.application.showdevtools).
        */
        void SetShowDevTools(bool showDevTools);

        /**
        Returns a Boolean that represents whether to display Mini toolbars when the user right-clicks in the workbook window. Since MS Excel 2007.

        [MSDN documentation for Application.ShowMenuFloaties](https://learn.microsoft.com/en-us/office/vba/api/excel.application.showmenufloaties).
        */
        bool GetShowMenuFloaties();
        /**
        Sets a Boolean that represents whether to display Mini toolbars when the user right-clicks in the workbook window. Since MS Excel 2007.

        [MSDN documentation for Application.ShowMenuFloaties](https://learn.microsoft.com/en-us/office/vba/api/excel.application.showmenufloaties).
        */
        void SetShowMenuFloaties(bool showMenuFloaties);

        /**
        Controls whether the Quick Analysis contextual user interface is displayed on selection.

        [Excel VBA documentation for Application.ShowQuickAnalysis](https://learn.microsoft.com/en-us/office/vba/api/excel.application.showquickanalysis)
        */
        bool GetShowQuickAnalysis();

        /**
        Controls whether the Quick Analysis contextual user interface is displayed on selection.

        [Excel VBA documentation for Application.ShowQuickAnalysis](https://learn.microsoft.com/en-us/office/vba/api/excel.application.showquickanalysis)
        */
        void SetShowQuickAnalysis(bool showQuickAnalysis);

        /**
        Returns a Boolean that represents whether Mini toolbars displays when a user selects text. Since MS Excel 2007.

        [MSDN documentation for Application.ShowSelectionFloaties](https://learn.microsoft.com/en-us/office/vba/api/excel.application.showselectionfloaties).
        */
        bool GetShowSelectionFloaties();
        /**
        Sets a Boolean that represents whether Mini toolbars displays when a user selects text. Since MS Excel 2007.

        [MSDN documentation for Application.ShowSelectionFloaties](https://learn.microsoft.com/en-us/office/vba/api/excel.application.showselectionfloaties).
        */
        void SetShowSelectionFloaties(bool showSelectionFloaties);

        /**
        Returns True (default is False) when the New Workbook task pane appears for a Microsoft Excel application.

        [MSDN documentation for Application.ShowStartupDialog](https://learn.microsoft.com/en-us/office/vba/api/excel.application.showstartupdialog).
        */
        bool GetShowStartupDialog();
        /**
        Returns True (default is False) when the New Workbook task pane appears for a Microsoft Excel application.

        [MSDN documentation for Application.ShowStartupDialog](https://learn.microsoft.com/en-us/office/vba/api/excel.application.showstartupdialog).
        */
        void SetShowStartupDialog(bool showStartupDialog);

        /**
        True if ToolTips are turned on.

        [MSDN documentation for Application.ShowToolTips](https://learn.microsoft.com/en-us/office/vba/api/excel.application.showtooltips).
        */
        bool GetShowToolTips();
        /**
        True if ToolTips are turned on.

        [MSDN documentation for Application.ShowToolTips](https://learn.microsoft.com/en-us/office/vba/api/excel.application.showtooltips).
        */
        void SetShowToolTips(bool showToolTips);

        /**
        True if there's a separate Windows taskbar button for each open workbook. The default value is True.

        [MSDN documentation for Application.ShowWindowsInTaskbar](https://learn.microsoft.com/en-us/previous-versions/office/developer/office-2007/bb209238%28v=office.12%29).
        */
        bool GetShowWindowsInTaskbar();
        /**
        True if there's a separate Windows taskbar button for each open workbook. The default value is True.

        [MSDN documentation for Application.ShowWindowsInTaskbar](https://learn.microsoft.com/en-us/previous-versions/office/developer/office-2007/bb209238%28v=office.12%29).
        */
        void SetShowWindowsInTaskbar(bool showWindowsInTaskbar);

        /**
        Returns the name of the standard font.

        [MSDN documentation for Application.StandardFont](https://learn.microsoft.com/en-us/office/vba/api/excel.application.standardfont).
        */
        wxString GetStandardFont();
        /**
        Sets the name of the standard font. The change doesn't take effect until you restart Microsoft Excel.

        [MSDN documentation for Application.StandardFont](https://learn.microsoft.com/en-us/office/vba/api/excel.application.standardfont).
        */
        void SetStandardFont(const wxString& standardFont);

        /**
        Returns the standard font size, in points.

        [MSDN documentation for Application.StandardFontSize](https://learn.microsoft.com/en-us/office/vba/api/excel.application.standardfontsize).
        */
        double GetStandardFontSize();
        /**
        Sets the standard font size, in points. The change doesn't take effect until you restart Microsoft Excel.

        [MSDN documentation for Application.StandardFontSize](https://learn.microsoft.com/en-us/office/vba/api/excel.application.standardfontsize).
        */
        void SetStandardFontSize(double standardFontSize);

        /**
        Returns the complete path of the startup folder, excluding the final separator.

        [MSDN documentation for Application.StartupPath](https://learn.microsoft.com/en-us/office/vba/api/excel.application.startuppath).
        */
        wxString GetStartupPath();

        /**
        Returns the text in the status bar.

        [MSDN documentation for Application.StatusBar](https://learn.microsoft.com/en-us/office/vba/api/excel.application.statusbar).
        */
        wxString GetStatusBar();
        //@{
        /**
        Sets the text in the status bar. The second version restores the default status bar text.

        [MSDN documentation for Application.StatusBar](https://learn.microsoft.com/en-us/office/vba/api/excel.application.statusbar).
        */
        void SetStatusBar(const wxString& statusBar);
        void SetStatusBar();
        //@}

        /**
        Returns the local path where templates are stored.

        [MSDN documentation for Application.TemplatesPath](https://learn.microsoft.com/en-us/office/vba/api/excel.application.templatespath).
        */
        wxString GetTemplatesPath();

        /**
        Returns the character used for the thousands separator as a String.

        [MSDN documentation for Application.ThousandsSeparator](https://learn.microsoft.com/en-us/office/vba/api/excel.application.thousandsseparator).
        */
        wxString GetThousandsSeparator();
        /**
        Sets the character used for the thousands separator as a String.

        [MSDN documentation for Application.ThousandsSeparator](https://learn.microsoft.com/en-us/office/vba/api/excel.application.thousandsseparator).
        */
        void SetThousandsSeparator(const wxString& thousandsSeparator);

        /**
        Returns a Double value that represents the distance, in points, from the top edge of the screen to the top edge of the main Microsoft Excel window.

        [MSDN documentation for Application.Top](https://learn.microsoft.com/en-us/office/vba/api/excel.application.top).
        */
        double GetTop();
        /**
        Sets a Double value that represents the distance, in points, from the top edge of the screen to the top edge of the main Microsoft Excel window.

        [MSDN documentation for Application.Top](https://learn.microsoft.com/en-us/office/vba/api/excel.application.top).
        */
        void SetTop(double top);

        /**
        Returns the Microsoft Excel menu or help key, which is usually "/".

        [MSDN documentation for Application.TransitionMenuKey](https://learn.microsoft.com/en-us/office/vba/api/excel.application.transitionmenukey).
        */
        wxString GetTransitionMenuKey();
        /**
        Sets the Microsoft Excel menu or help key, which is usually "/".

        [MSDN documentation for Application.TransitionMenuKey](https://learn.microsoft.com/en-us/office/vba/api/excel.application.transitionmenukey).
        */
        void SetTransitionMenuKey(const wxString& transitionMenuKey);

        /**
        Returns the action taken when the Microsoft Excel menu key is pressed. Can be either xlExcelMenus or xlLotusHelp.

        [MSDN documentation for Application.TransitionMenuKeyAction](https://learn.microsoft.com/en-us/office/vba/api/excel.application.transitionmenukeyaction).
        */
        long GetTransitionMenuKeyAction();
        /**
        Sets the action taken when the Microsoft Excel menu key is pressed. Can be either xlExcelMenus or xlLotusHelp.

        [MSDN documentation for Application.TransitionMenuKeyAction](https://learn.microsoft.com/en-us/office/vba/api/excel.application.transitionmenukeyaction).
        */
        void SetTransitionMenuKeyAction(long transitionMenuKeyAction);

        /**
        True if transition navigation keys are active.

        [MSDN documentation for Application.TransitionNavigKeys](https://learn.microsoft.com/en-us/office/vba/api/excel.application.transitionnavigkeys).
        */
        bool GetTransitionNavigKeys();
        /**
        True if transition navigation keys are active.

        [MSDN documentation for Application.TransitionNavigKeys](https://learn.microsoft.com/en-us/office/vba/api/excel.application.transitionnavigkeys).
        */
        void SetTransitionNavigKeys(bool transitionNavigKeys);

        /**
        Returns the maximum height of the space that a window can occupy in the application window area, in points.

        [MSDN documentation for Application.UsableHeight](https://learn.microsoft.com/en-us/office/vba/api/excel.application.usableheight).
        */
        double GetUsableHeight();

        /**
        Returns the maximum width of the space that a window can occupy in the application window area, in points.

        [MSDN documentation for Application.UsableWidth](https://learn.microsoft.com/en-us/office/vba/api/excel.application.usablewidth).
        */
        double GetUsableWidth();

        /**
        True if the application is visible or if it was created or started by the user. False if you created or started the application programmatically by using the CreateObject or GetObject functions, and the application is hidden.

        [MSDN documentation for Application.UserControl](https://learn.microsoft.com/en-us/office/vba/api/excel.application.usercontrol).
        */
        bool GetUserControl();
        /**
        True if the application is visible or if it was created or started by the user. False if you created or started the application programmatically by using the CreateObject or GetObject functions, and the application is hidden.

        [MSDN documentation for Application.UserControl](https://learn.microsoft.com/en-us/office/vba/api/excel.application.usercontrol).
        */
        void SetUserControl(bool userControl);

        /**
        Returns the path to the location on the user's computer where the COM add-ins are installed.

        [MSDN documentation for Application.UserLibraryPath](https://learn.microsoft.com/en-us/office/vba/api/excel.application.userlibrarypath).
        */
        wxString GetUserLibraryPath();
        /**
        Returns the name of the current user.

        [MSDN documentation for Application.UserName](https://learn.microsoft.com/en-us/office/vba/api/excel.application.username).
        */
        wxString GetUsername();
        /**
        Sets the name of the current user.

        [MSDN documentation for Application.UserName](https://learn.microsoft.com/en-us/office/vba/api/excel.application.username).
        */
        void SetUsername(const wxString& userName);

        /**
        True (default) if the system separators of Microsoft Excel are enabled.

        [MSDN documentation for Application.UseSystemSeparators](https://learn.microsoft.com/en-us/office/vba/api/excel.application.usesystemseparators).
        */
        bool GetUseSystemSeparators();
        /**
        True (default) if the system separators of Microsoft Excel are enabled.

        [MSDN documentation for Application.UseSystemSeparators](https://learn.microsoft.com/en-us/office/vba/api/excel.application.usesystemseparators).
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
        Excel 2016 = 16

        [MSDN documentation for Application.Version](https://learn.microsoft.com/en-us/office/vba/api/excel.application.version).

        See also GetVersionAsDouble_(), GetVersionAsEnum_(), and IsVersionAtLeast_() methods.
        */
        wxString GetVersion();

        /**
        Returns a Boolean value that determines whether the object is visible.

        [MSDN documentation for Application.Visible](https://learn.microsoft.com/en-us/office/vba/api/excel.application.visible).
        */
        bool GetVisible();
        /**
        Sets a Boolean value that determines whether the object is visible.

        [MSDN documentation for Application.Visible](https://learn.microsoft.com/en-us/office/vba/api/excel.application.visible).
        */
        void SetVisible(bool visible);

        /**
        Returns a Double value that represents the distance, in points, from the left edge of the application window to its right edge.

        [MSDN documentation for Application.Width](https://learn.microsoft.com/en-us/office/vba/api/excel.application.width).
        */
        double GetWidth();
        /**
        Sets a Double value that represents the distance, in points, from the left edge of the application window to its right edge.

        [MSDN documentation for Application.Width](https://learn.microsoft.com/en-us/office/vba/api/excel.application.width).
        */
        void SetWidth(double width);

        /**
        Returns a Windows collection that represents all the windows in all the workbooks.

        [MSDN documentation for Application.Windows](https://learn.microsoft.com/en-us/office/vba/api/excel.application.windows).
        */
        wxExcelWindows GetWindows();

        /**
        True if the computer is running under Microsoft Windows for Pen Computing.

        [MSDN documentation for Application.WindowsForPens](https://learn.microsoft.com/en-us/office/vba/api/excel.application.windowsforpens).
        */
        bool GetWindowsForPens();

        /**
        Returns the state of the window.

        [MSDN documentation for Application.WindowState](https://learn.microsoft.com/en-us/office/vba/api/excel.application.windowstate).
        */
        XlWindowState GetWindowState();
        /**
        Sets the state of the window.

        [MSDN documentation for Application.WindowState](https://learn.microsoft.com/en-us/office/vba/api/excel.application.windowstate).
        */
        void SetWindowState(XlWindowState windowState);

        /**
        Returns a Workbooks collection that represents all the open workbooks.

        [MSDN documentation for Application.Workbooks](https://learn.microsoft.com/en-us/office/vba/api/excel.application.workbooks).
        */
        wxExcelWorkbooks GetWorkbooks();

        /**
        For an Application object, returns a Sheets collection that represents all the worksheets in the active workbook. For a Workbook object, returns a Sheets collection that represents all the worksheets in the specified workbook.

        [MSDN documentation for Application.Worksheets](https://learn.microsoft.com/en-us/office/vba/api/excel.application.worksheets).
        */
        wxExcelWorksheets GetWorksheets();

        /**
            Returns MS Excel version as a double matching Application.Version.
        */
        bool GetVersionAsDouble_(double& version);
        /**
        Returns MS Excel version as an ExcelForWindowsVersions enum value.
        */
        bool GetVersionAsEnum_(ExcelForWindowsVersions& version);
        /**
            Returns true if the instance is @a version or newer.
        */
        bool IsVersionAtLeast_(ExcelForWindowsVersions version);

        /**
            Returns the list separator of the Windows regional settings, i.e. GetInternational()
            for xlListSeparator, or an empty string if it could not be obtained.

            MS Excel uses this separator in the addresses of ranges with several areas,
            e.g. "A1:B2,D4:E5", passed to wxExcelRangeOwner::GetRange() or another method
            or property taking an address, regardless of the automation LCID.
            Therefore, build such addresses with this separator, or combine the ranges
            of the areas with Union() instead:
            @code
            wxString address = "A1:B2" + app.GetListSeparator_() + "D4:E5";
            @endcode
            wxExcelRange::GetAddressLocal() returns addresses with this separator, while the
            separator in wxExcelRange::GetAddress() depends on how the range was obtained.
        */
        wxString GetListSeparator_();


        /**
        Returns "Application".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("Application"); }
    private:
        bool RangesToVariants(const wxExcelRangeVector& ranges, wxVariantVector& variants);



    };



} // namespace wxAutoExcel

#endif // _WXAUTOEXCEL_APPLICATION_H

