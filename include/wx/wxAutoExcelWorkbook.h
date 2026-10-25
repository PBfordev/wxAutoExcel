/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// License:     MIT license
/////////////////////////////////////////////////////////////////////////////


#ifndef _WXAUTOEXCEL_WORKBOOK_H
#define _WXAUTOEXCEL_WORKBOOK_H

#include "wx/wxAutoExcel_defs.h"

#include "wx/wxAutoExcel_object.h"
#include "wx/wxAutoExcel_enums.h"

class WXDLLIMPEXP_FWD_CORE wxColour;

namespace wxAutoExcel {


    /**
    @brief Represents Microsoft Excel Workbook.
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelWorkbook : public wxExcelObject
    {
    public:
        // ***** METHODS *****

        /**
        Accepts all changes in the specified shared workbook.

        [MSDN documentation for Workbook.AcceptAllChanges](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.acceptallchanges).
        */
        void AcceptAllChanges(XlHighlightChangesTime* when = NULL, const wxString& who = wxEmptyString, const wxString& where = wxEmptyString);

        /**
        Activates the first window associated with the workbook.

        [MSDN documentation for Workbook.Activate](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.activate%28method%29)).
        */
        bool Activate();

        /**
        Adds a shortcut to the workbook or hyperlink to the Favorites folder.

        [MSDN documentation for Workbook.AddToFavorites](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.addtofavorites).
        */
        void AddToFavorites();

        /**
        Applies the specified theme to the current workbook.

        [MSDN documentation for Workbook.ApplyTheme](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.applytheme).
        */
        void ApplyTheme(const wxString& fileName);

        /**
        Converts formulas linked to other Microsoft Excel sources or OLE sources to values.

        [MSDN documentation for Workbook.BreakLink](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.breaklink).
        */
        void BreakLink(const wxString& name, XlLinkType type);

        /**
        True if Microsoft Excel can check in a specified workbook to a server.

        [MSDN documentation for Workbook.CanCheckIn](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.cancheckin).
        */
        bool CanCheckIn();

        /**
        Changes the access permissions for the workbook. This may require an updated version to be loaded from the disk.

        [MSDN documentation for Workbook.ChangeFileAccess](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.changefileaccess).
        */
        void ChangeFileAccess(XlFileAccess mode, const wxString& writePassword, wxXlTribool notify);

        /**
        Changes a link from one document to another.

        [MSDN documentation for Workbook.ChangeLink](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.changelink).
        */
        void ChangeLink(const wxString& name, const wxString& newName, XlLinkType* type);

        /**
        Returns a workbook from a local computer to a server, and sets the local workbook to read-only so that it cannot be edited locally. Calling this method will also close the workbook.

        [MSDN documentation for Workbook.CheckIn](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.checkin).
        */
        void CheckIn(wxXlTribool saveChanges = wxDefaultXlTribool,
            wxXlTribool comments = wxDefaultXlTribool,
            wxXlTribool makePublic = wxDefaultXlTribool);

        /**
        Saves a workbook to a server from a local computer, and sets the local workbook to read-only so that it cannot be edited locally. Since Excel 2010

        [MSDN documentation for Workbook.CheckInWithVersion](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.checkinwithversion).
        */
        void CheckInWithVersion(wxXlTribool saveChanges = wxDefaultXlTribool,
            wxXlTribool comments = wxDefaultXlTribool, wxXlTribool makePublic = wxDefaultXlTribool,
            XlCheckInVersionType* versionType = NULL);

        /**

        Closes the object.

        [MSDN documentation for Workbook.Close](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.close).
        */
        bool Close(wxXlTribool saveChanges = wxDefaultXlTribool, const wxString& fileName = wxEmptyString, wxXlTribool routeWorkbook = wxDefaultXlTribool);

        /**
        Deletes a custom number format from the workbook.

        [MSDN documentation for Workbook.DeleteNumberFormat](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.deletenumberformat).
        */
        void DeleteNumberFormat(const wxString& numberFormat);

        /**
        The EnableConnections method allows developers to programmatically enable data connections within the workbook for the user.

        [MSDN documentation for Workbook.EnableConnections](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.enableconnections).
        */
        void EnableConnections();

        /**
        Terminates a review of a file that has been sent for review using the SendForReview method.

        [MSDN documentation for Workbook.EndReview](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.endreview).
        */
        void EndReview();

        /**
        Assigns the current user exclusive access to the workbook that's open as a shared list.

        [MSDN documentation for Workbook.ExclusiveAccess](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.exclusiveaccess).
        */
        bool ExclusiveAccess();

        //@{
        /**
        The ExportAsFixedFormat method is used to publish a workbook to either the PDF or XPS format.

        [MSDN documentation for Workbook.ExportAsFixedFormat](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.exportasfixedformat).
        */
        void ExportAsFixedFormat(XlFixedFormatType type, const wxString& fileName = wxEmptyString,
            XlFixedFormatQuality* quality = NULL, wxXlTribool includeDocProperties = wxDefaultXlTribool,
            wxXlTribool ignorePrintAreas = wxDefaultXlTribool,
            long* from = NULL, long* to = NULL, wxXlTribool openAfterPublish = wxDefaultXlTribool);

        void ExportAsFixedFormat(XlFixedFormatType type, const wxVariantVector& optionalArgs);
        //@}

        /**
        Controls how changes are shown in a shared workbook.

        [MSDN documentation for Workbook.HighlightChangesOptions](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.highlightchangesoptions).
        */
        void HighlightChangesOptions(XlHighlightChangesTime* when = NULL, const wxString& who = wxEmptyString, const wxString& where = wxEmptyString);

        /**
        Locks the workbook on the server to prevent modification.

        [MSDN documentation for Workbook.LockServerFile](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.lockserverfile).
        */
        void LockServerFile();

        /**
        Merges changes from one workbook into an open workbook.

        [MSDN documentation for Workbook.MergeWorkbook](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.mergeworkbook).
        */
        void MergeWorkbook(const wxString& fileName);

        /**
        Creates a new window or a copy of the specified window.

        [MSDN documentation for Workbook.NewWindow](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.newwindow).
        */
        wxExcelWindow NewWindow();

        /**
        Posts the specified workbook to a public folder. This method works only with a Microsoft Exchange client connected to a Microsoft Exchange server.

        [MSDN documentation for Workbook.Post](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.post).
        */
        void Post();

        //@{
        /**
        Prints the object.

        [MSDN documentation for Workbook.PrintOut](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.printout).
        */
        bool PrintOut(long* from = NULL, long* to = NULL, long* copies = NULL, wxXlTribool preview = wxDefaultXlTribool,
            const wxString& activePrinter = wxEmptyString, wxXlTribool printToFile = wxDefaultXlTribool,
            wxXlTribool collate = wxDefaultXlTribool, const wxString& prToFileName= wxEmptyString, wxXlTribool ignorePrintAreas = wxDefaultXlTribool);
        bool PrintOut(const wxVariantVector& args);
        //@}

        /**
        Shows a preview of the object as it would look when printed.

        [MSDN documentation for Workbook.PrintPreview](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.printpreview).
        */
        bool PrintPreview(wxXlTribool enableChanges = wxDefaultXlTribool);

        /**
        Protects a workbook so that it cannot be modified.

        [MSDN documentation for Workbook.Protect](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.protect).
        */
        void Protect(const wxString& password = wxEmptyString, wxXlTribool structure = wxDefaultXlTribool, wxXlTribool windows = wxDefaultXlTribool);

        //@{
        /**
        Saves the workbook and protects it for sharing.

        [MSDN documentation for Workbook.ProtectSharing](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.protectsharing).
        */
        void ProtectSharing(const wxString& fileName = wxEmptyString, const wxString& password = wxEmptyString,
            const wxString& writeResPassword = wxEmptyString,
            wxXlTribool readOnlyRecommended = wxDefaultXlTribool, wxXlTribool createBackup = wxDefaultXlTribool,
            const wxString& sharingPassword = wxEmptyString);

        void ProtectSharing(const wxVariantVector& optionalArgs);
        //@}

        /**
        Removes entries from the change log for the specified workbook.

        [MSDN documentation for Workbook.PurgeChangeHistoryNow](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.purgechangehistorynow).
        */
        void PurgeChangeHistoryNow(long days, const wxString& sharingPassword = wxEmptyString);

        /**
        Refreshes all external data ranges and PivotTable reports in the specified workbook.

        [MSDN documentation for Workbook.RefreshAll](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.refreshall).
        */
        void RefreshAll();

        /**
        Rejects all changes in the specified shared workbook.

        [MSDN documentation for Workbook.RejectAllChanges](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.rejectallchanges).
        */
        void RejectAllChanges(XlHighlightChangesTime* when = NULL, const wxString& who = wxEmptyString, const wxString& where = wxEmptyString);

        /**
        Reloads a workbook based on an HTML document, using the specified document encoding.

        [MSDN documentation for Workbook.ReloadAs](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.reloadas).
        */
        void ReloadAs(MsoEncoding encoding);

        /**
        Removes all information of the specified type from the workbook.

        [MSDN documentation for Workbook.RemoveDocumentInformation](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.removedocumentinformation).
        */
        void RemoveDocumentInformation(XlRemoveDocInfoType removeDocInfoType);

        /**
        Disconnects the specified user from the shared workbook.

        [MSDN documentation for Workbook.RemoveUser](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.removeuser).
        */
        void RemoveUser(long index);

        /**
        Sends an e-mail message to the author of a workbook that has been sent out for review, notifying them that a reviewer has completed review of the workbook.

        [MSDN documentation for Workbook.ReplyWithChanges](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.replywithchanges).
        */
        void ReplyWithChanges(wxXlTribool showMessage = wxDefaultXlTribool);

        /**
        Resets the color palette to the default colors.

        [MSDN documentation for Workbook.ResetColors](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.resetcolors).
        */
        void ResetColors();

        /**
        Runs the Auto_Open, Auto_Close, Auto_Activate, or Auto_Deactivate macro attached to the workbook.

        [MSDN documentation for Workbook.RunAutoMacros](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.runautomacros).
        */
        void RunAutoMacros(XlRunAutoMacro which);

        /**
        Saves changes to the specified workbook.

        [MSDN documentation for Workbook.Save](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.save).
        */
        bool Save();

        //@{
        /**
        Saves changes to the workbook in a different file.

        [MSDN documentation for Workbook.SaveAs](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.saveas).
        */
        bool SaveAs(const wxString& fileName = wxEmptyString, XlFileFormat* fileFormat = NULL,
            const wxString& password = wxEmptyString, const wxString& writeResPassword = wxEmptyString,
            wxXlTribool readOnlyRecommended = wxDefaultXlTribool, wxXlTribool createBackup = wxDefaultXlTribool,
            XlSaveAsAccessMode* accessMode = NULL, XlSaveConflictResolution* conflictResolution = NULL,
            wxXlTribool addToMru = wxDefaultXlTribool, wxXlTribool local = wxDefaultXlTribool);

        bool SaveAs(const wxVariantVector& optionalArgs);
        //@}

        /**
        Saves a copy of the workbook to a file but doesn't modify the open workbook in memory.

        [MSDN documentation for Workbook.SaveCopyAs](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.savecopyas).
        */
        bool SaveCopyAs(const wxString& fileName);

        /**
        Sends a worksheet as a fax to the specfied recipients.

        [MSDN documentation for Workbook.SendFaxOverInternet](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.sendfaxoverinternet).
        */
        void SendFaxOverInternet(const wxString& recipients = wxEmptyString, const wxString& subject= wxEmptyString,
            wxXlTribool showMessage = wxDefaultXlTribool);

        /**
        Sends a workbook in an e-mail message for review to the specified recipients.

        [MSDN documentation for Workbook.SendForReview](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.sendforreview).
        */
        void SendForReview(const wxString& recipients = wxEmptyString, const wxString& subject= wxEmptyString,
            wxXlTribool showMessage = wxDefaultXlTribool, wxXlTribool includeAttachment = wxDefaultXlTribool);

        /**
        Sends the workbook by using the installed mail system.

        [MSDN documentation for Workbook.SendMail](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.sendmail).
        */
        void SendMail(const wxString& recipients = wxEmptyString, const wxString& subject= wxEmptyString,
            wxXlTribool returnReceipt = wxDefaultXlTribool);

        /**
        Sets the options for encrypting workbooks using passwords.

        [MSDN documentation for Workbook.SetPasswordEncryptionOptions](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.setpasswordencryptionoptions).
        */
        void SetPasswordEncryptionOptions(const wxString& passwordEncryptionProvider = wxEmptyString,
            const wxString& passwordEncryptionAlgorithm  = wxEmptyString,
            long* passwordEncryptionKeyLength = NULL,
            wxXlTribool passwordEncryptionFileProperties = wxDefaultXlTribool);

        /**
        Removes protection from a sheet or workbook. This method has no effect if the sheet or workbook isn't protected.

        [MSDN documentation for Workbook.Unprotect](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.unprotect).
        */
        bool Unprotect(const wxString& password = wxEmptyString);

        /**
        Turns off protection for sharing and saves the workbook.

        [MSDN documentation for Workbook.UnprotectSharing](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.unprotectsharing).
        */
        void UnprotectSharing(const wxString& sharingPassword = wxEmptyString);

        /**
        Updates a read-only workbook from the saved disk version of the workbook if the disk version is more recent than the copy of the workbook that is loaded in memory. If the disk copy hasn't changed since the workbook was loaded, the in-memory copy of the workbook isn't reloaded.

        [MSDN documentation for Workbook.UpdateFromFile](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.updatefromfile).
        */
        void UpdateFromFile();

        /**
        Displays a preview of the specified workbook as it would look if saved as a Web page.

        [MSDN documentation for Workbook.WebPagePreview](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.webpagepreview).
        */
        void WebPagePreview();

        // ***** PROPERTIES *****

#if WXAUTOEXCEL_USE_CHARTS
        /**
        Returns a Chart object that represents the active chart (either an embedded chart or a chart sheet). An embedded chart is considered active when it's either selected or activated. When no chart is active, this property returns Nothing.

        [MSDN documentation for Workbook.ActiveChart]().
        */
        wxExcelChart GetActiveChart();
#endif // #if WXAUTOEXCEL_USE_CHARTS

        /**
        Returns the active sheet (the sheet on top) in the active workbook or in the specified window or workbook.

        [MSDN documentation for Workbook.ActiveSheet](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.activesheet).
        */
        wxExcelSheet GetActiveSheet();

        /**
        Specifies whether certain worksheet functions use the latest accuracy algorithms to calculate their results. Since Excel 2010.

        [MSDN documentation for Workbook.AccuracyVersion](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.accuracyversion).
        */
        long GetAccuracyVersion();

        /**
        Specifies whether certain worksheet functions use the latest accuracy algorithms to calculate their results. Since Excel 2010.

        [MSDN documentation for Workbook.AccuracyVersion](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.accuracyversion).
        */
        void SetAccuracyVersion(long accuracyVersion);

        /**
        Returns the number of minutes between automatic updates to the shared workbook.

        [MSDN documentation for Workbook.AutoUpdateFrequency](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.autoupdatefrequency).
        */
        long GetAutoUpdateFrequency();
        /**
        Sets the number of minutes between automatic updates to the shared workbook.

        [MSDN documentation for Workbook.AutoUpdateFrequency](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.autoupdatefrequency).
        */
        void SetAutoUpdateFrequency(long autoUpdateFrequency);

        /**
        True if current changes to the shared workbook are posted to other users whenever the workbook is automatically updated. False if changes aren't posted (this workbook is still synchronized with changes made by other users). The default value is True.

        [MSDN documentation for Workbook.AutoUpdateSaveChanges](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.autoupdatesavechanges).
        */
        bool GetAutoUpdateSaveChanges();
        /**
        True if current changes to the shared workbook are posted to other users whenever the workbook is automatically updated. False if changes aren't posted (this workbook is still synchronized with changes made by other users). The default value is True.

        [MSDN documentation for Workbook.AutoUpdateSaveChanges](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.autoupdatesavechanges).
        */
        void SetAutoUpdateSaveChanges(bool autoUpdateSaveChanges);


        /**
        Returns a DocumentProperties collection that represents all the built-in document properties.

        [MSDN documentation for Workbook.BuiltinDocumentProperties](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.builtindocumentproperties).
        */
        wxExcelDocumentProperties GetBuiltinDocumentProperties();

        /**
        Returns the information about the version of Excel that the workbook was last fully recalculated by.

        [MSDN documentation for Workbook.CalculationVersion](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.calculationversion).
        */
        long GetCalculationVersion();
        /**
        Returns the number of days shown in the shared workbook's change history.

        [MSDN documentation for Workbook.ChangeHistoryDuration](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.changehistoryduration).
        */
        long GetChangeHistoryDuration();
        /**
        Sets the number of days shown in the shared workbook's change history.

        [MSDN documentation for Workbook.ChangeHistoryDuration](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.changehistoryduration).
        */
        void SetChangeHistoryDuration(long changeHistoryDuration);

#if WXAUTOEXCEL_USE_CHARTS
        /**
        Returns a Sheets collection that represents all the chart sheets in the specified workbook.

        [MSDN documentation for Workbook.Charts](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.charts).
        */
        wxExcelCharts GetCharts();
#endif // #if WXAUTOEXCEL_USE_CHARTS

        /**
        Controls whether or not the compatibility checker is run automatically when the workbook is saved. Since MS Excel 2007.

        [MSDN documentation for Workbook.CheckCompatibility](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.checkcompatibility).
        */
        bool GetCheckCompatibility();
        /**
        Controls whether or not the compatibility checker is run automatically when the workbook is saved. Since MS Excel 2007.

        [MSDN documentation for Workbook.CheckCompatibility](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.checkcompatibility).
        */
        void SetCheckCompatibility(bool checkCompatibility);

        //@{
        /**
        Returns colors in the palette for the workbook. The palette has 56 entries, each represented by an RGB value.

        [MSDN documentation for Workbook.Colors](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.colors).
        */
        wxColour GetColors(long index);
        wxVector<wxColour> GetColors();
        //@}

        //@{
        /**
        Sets colors in the palette for the workbook. The palette has 56 entries, each represented by an RGB value.

        [MSDN documentation for Workbook.Colors](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.colors).
        */
        void SetColors(long index, const wxColour& color);
        void SetColors(const wxVector<wxColour>& colors);
        //@}

        /**
        [MSDN documentation for Workbook.ConflictResolution](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.conflictresolution).
        */
        XlSaveConflictResolution GetConflictResolution();
        /**
        Sets the way conflicts are to be resolved whenever a shared workbook is updated. Read/write XlSaveConflictResolution.

        [MSDN documentation for Workbook.ConflictResolution](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.conflictresolution).
        */
        void SetConflictResolution(XlSaveConflictResolution conflictResolution);

        /**
        Disables the external connections or links in the workbook. Since MS Excel 2007.

        [MSDN documentation for Workbook.ConnectionsDisabled](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.connectionsdisabled).
        */
        bool GetConnectionsDisabled();

        /**
        True if a backup file is created when this file is saved.

        [MSDN documentation for Workbook.CreateBackup](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.createbackup).
        */
        bool GetCreateBackup();

         /**
        Returns a DocumentProperties collection that represents all the custom document properties.

        [MSDN documentation for Workbook.CustomDocumentProperties](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.customdocumentproperties).
        */
        wxExcelDocumentProperties GetCustomDocumentProperties();

        /**
        True if the workbook uses the 1904 date system.

        [MSDN documentation for Workbook.Date1904](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.date1904).
        */
        bool GetDate1904();
        /**
        True if the workbook uses the 1904 date system.

        [MSDN documentation for Workbook.Date1904](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.date1904).
        */
        void SetDate1904(bool date1904);

        /**
        A Boolean value that determines whether ink comments are displayed in the workbook.

        [MSDN documentation for Workbook.DisplayInkComments](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.displayinkcomments).
        */
        bool GetDisplayInkComments();
        /**
        A Boolean value that determines whether ink comments are displayed in the workbook.

        [MSDN documentation for Workbook.DisplayInkComments](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.displayinkcomments).
        */
        void SetDisplayInkComments(bool displayInkComments);

        /**
        Returns if the user should be prompted to convert the workbook if the workbook contains features that are not supported by versions of Excel earlier than Excel 2007. Since MS Excel 2007.

        [MSDN documentation for Workbook.DoNotPromptForConvert](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.donotpromptforconvert).
        */
        bool GetDoNotPromptForConvert();
        /**
        Sets if the user should be prompted to convert the workbook if the workbook contains features that are not supported by versions of Excel earlier than Excel 2007. Since MS Excel 2007.

        [MSDN documentation for Workbook.DoNotPromptForConvert](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.donotpromptforconvert).
        */
        void SetDoNotPromptForConvert(bool doNotPromptForConvert);

        /**
        Saves changed files, of all formats, on a timed interval.

        [MSDN documentation for Workbook.EnableAutoRecover](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.enableautorecover).
        */
        bool GetEnableAutoRecover();
        /**
        Saves changed files, of all formats, on a timed interval.

        [MSDN documentation for Workbook.EnableAutoRecover](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.enableautorecover).
        */
        void SetEnableAutoRecover(bool enableAutoRecover);

        /**
        Returns a string specifying the name of the algorithm encryption provider that Excel uses when encrypting documents. Since MS Excel 2007.

        [MSDN documentation for Workbook.EncryptionProvider](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.encryptionprovider).
        */
        wxString GetEncryptionProvider();
        /**
        Sets the name of the algorithm encryption provider that Excel uses when encrypting documents. Since MS Excel 2007.

        [MSDN documentation for Workbook.EncryptionProvider](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.encryptionprovider).
        */
        void SetEncryptionProvider(const wxString& encryptionProvider);

        /**
        True if the e-mail composition header and the envelope toolbar are both visible.

        [MSDN documentation for Workbook.EnvelopeVisible](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.envelopevisible).
        */
        bool GetEnvelopeVisible();
        /**
        True if the e-mail composition header and the envelope toolbar are both visible.

        [MSDN documentation for Workbook.EnvelopeVisible](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.envelopevisible).
        */
        void SetEnvelopeVisible(bool envelopeVisible);

        /**
        The Excel8CompatibilityMode property provides developers with a way to check if the workbook is in compatibility mode. Since MS Excel 2007.

        [MSDN documentation for Workbook.Excel8CompatibilityMode](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.excel8compatibilitymode).
        */
        bool GetExcel8CompatibilityMode();

        /**
        Returns the file format and/or type of the workbook.

        [MSDN documentation for Workbook.FileFormat](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.fileformat).
        */
        XlFileFormat GetFileFormat();

        /**
        Returns a Boolean that indicates whether a workbook is final. Since MS Excel 2007.

        [MSDN documentation for Workbook.Final](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.final).
        */
        bool GetFinal();
        /**
        Sets a Boolean that indicates whether a workbook is final.

        [MSDN documentation for Workbook.Final](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.final).
        */
        void SetFinal(bool final);

        /**
        Forces a full calculation of a workbook. Since MS Excel 2007.

        [MSDN documentation for Workbook.ForceFullCalculation](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.forcefullcalculation).
        */
        bool GetForceFullCalculation();
        /**
        Forces a full calculation of a workbook. Since MS Excel 2007.

        [MSDN documentation for Workbook.ForceFullCalculation](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.forcefullcalculation).
        */
        void SetForceFullCalculation(bool forceFullCalculation);

        /**
        Returns the workbook name including its path on disk.

        [MSDN documentation for Workbook.FullName](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.fullname).
        */
        wxString GetFullName();

        /**
        Returns the workbook name including its path on disk.

        [MSDN documentation for Workbook.FullNameURLEncoded](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.fullnameurlencoded).
        */
        wxString GetFullNameURLEncoded();

        /**
        True if the workbook has a protection password.

        [MSDN documentation for Workbook.HasPassword](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.haspassword).
        */
        bool GetHasPassword();

        /**
        True if the workbook has an attached Microsoft Visual Basic for Applications project. Since MS Excel 2007.

        [MSDN documentation for Workbook.HasVBProject](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.hasvbproject).
        */
        bool GetHasVBProject();

        /**
        True if changes to the shared workbook are highlighted on-screen.

        [MSDN documentation for Workbook.HighlightChangesOnScreen](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.highlightchangesonscreen).
        */
        bool GetHighlightChangesOnScreen();
        /**
        True if changes to the shared workbook are highlighted on-screen.

        [MSDN documentation for Workbook.HighlightChangesOnScreen](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.highlightchangesonscreen).
        */
        void SetHighlightChangesOnScreen(bool highlightChangesOnScreen);

#if WXAUTOEXCEL_USE_CONDFORMAT
        //@{
        /**
        Returns the IconSets collection of all the icon sets, or the icon set given by @a index,
        e.g., to set it to an icon set conditional format. Since Excel 2007.

        [MSDN documentation for Workbook.IconSets](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.iconsets).
        */
        wxExcelIconSets GetIconSets();
        wxExcelIconSet GetIconSets(XlIconSet index);
        //@}
#endif // #if WXAUTOEXCEL_USE_CONDFORMAT

        /**
        True if the list borders are visible when a list is not active.

        [MSDN documentation for Workbook.InactiveListBorderVisible](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.inactivelistbordervisible).
        */
        bool GetInactiveListBorderVisible();
        /**
        True if the list borders are visible when a list is not active.

        [MSDN documentation for Workbook.InactiveListBorderVisible](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.inactivelistbordervisible).
        */
        void SetInactiveListBorderVisible(bool inactiveListBorderVisible);

        /**
        True if the workbook is running as an add-in.

        [MSDN documentation for Workbook.IsAddin](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.isaddin).
        */
        bool GetIsAddin();
        /**
        True if the workbook is running as an add-in.

        [MSDN documentation for Workbook.IsAddin](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.isaddin).
        */
        void SetIsAddin(bool isAddin);

        /**
        True if the specified workbook is being edited in place. False if the workbook has been opened in Microsoft Excel for editing.

        [MSDN documentation for Workbook.IsInplace](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.isinplace).
        */
        bool GetIsInplace();

        /**
        True if change tracking is enabled for the shared workbook.

        [MSDN documentation for Workbook.KeepChangeHistory](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.keepchangehistory).
        */
        bool GetKeepChangeHistory();
        /**
        True if change tracking is enabled for the shared workbook.

        [MSDN documentation for Workbook.KeepChangeHistory](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.keepchangehistory).
        */
        void SetKeepChangeHistory(bool keepChangeHistory);

        /**
        True if changes to the shared workbook are shown on a separate worksheet.

        [MSDN documentation for Workbook.ListChangesOnNewSheet](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.listchangesonnewsheet).
        */
        bool GetListChangesOnNewSheet();
        /**
        True if changes to the shared workbook are shown on a separate worksheet.

        [MSDN documentation for Workbook.ListChangesOnNewSheet](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.listchangesonnewsheet).
        */
        void SetListChangesOnNewSheet(bool listChangesOnNewSheet);

        /**
        True if the workbook is open as a shared list.

        [MSDN documentation for Workbook.MultiUserEditing](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.multiuserediting).
        */
        bool GetMultiUserEditing();

        /**
        Returns the name of the object.

        [MSDN documentation for Workbook.Name](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.name).
        */
        wxString GetName();

        /**
        Returns a Names collection that represents all the names in the specified workbook (including all worksheet-specific names).

        [MSDN documentation for Workbook.Names](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.names).
        */
        wxExcelNames GetNames();

        /**
        Returns the password that must be supplied to open the specified workbook.

        [MSDN documentation for Workbook.Password](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.password).
        */
        wxString GetPassword();
        /**
        Sets the password that must be supplied to open the specified workbook.

        [MSDN documentation for Workbook.Password](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.password).
        */
        void SetPassword(const wxString& password);

        /**
        Returns a string indicating the algorithm Microsoft Excel uses to encrypt passwords for the specified workbook.

        [MSDN documentation for Workbook.PasswordEncryptionAlgorithm](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.passwordencryptionalgorithm).
        */
        wxString GetPasswordEncryptionAlgorithm();
        /**
        True if Microsoft Excel encrypts file properties for the specified password-protected workbook.

        [MSDN documentation for Workbook.PasswordEncryptionFileProperties](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.passwordencryptionfileproperties).
        */
        bool GetPasswordEncryptionFileProperties();

        /**
        Returns the key length of the algorithm Microsoft Excel uses when encrypting passwords for the specified workbook.

        [MSDN documentation for Workbook.PasswordEncryptionKeyLength](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.passwordencryptionkeylength).
        */
        long GetPasswordEncryptionKeyLength();

        /**
        Returns the name of the algorithm encryption provider that Microsoft Excel uses when encrypting passwords for the specified workbook.

        [MSDN documentation for Workbook.PasswordEncryptionProvider](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.passwordencryptionprovider).
        */
        wxString GetPasswordEncryptionProvider();

        /**
        Returns the complete path to the workbook/file that this workbook object respresents.

        [MSDN documentation for Workbook.Path](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.path).
        */
        wxString GetPath();

        /**
        True if filter and sort settings for lists are included in the user's personal view of the shared workbook.

        [MSDN documentation for Workbook.PersonalViewListSettings](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.personalviewlistsettings).
        */
        bool GetPersonalViewListSettings();
        /**
        True if filter and sort settings for lists are included in the user's personal view of the shared workbook.

        [MSDN documentation for Workbook.PersonalViewListSettings](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.personalviewlistsettings).
        */
        void SetPersonalViewListSettings(bool personalViewListSettings);

        /**
        True if print settings are included in the user's personal view of the shared workbook.

        [MSDN documentation for Workbook.PersonalViewPrintSettings](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.personalviewprintsettings).
        */
        bool GetPersonalViewPrintSettings();
        /**
        True if print settings are included in the user's personal view of the shared workbook.

        [MSDN documentation for Workbook.PersonalViewPrintSettings](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.personalviewprintsettings).
        */
        void SetPersonalViewPrintSettings(bool personalViewPrintSettings);

        /**
        True if calculations in this workbook will be done using only the precision of the numbers as they're displayed.

        [MSDN documentation for Workbook.PrecisionAsDisplayed](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.precisionasdisplayed).
        */
        bool GetPrecisionAsDisplayed();
        /**
        True if calculations in this workbook will be done using only the precision of the numbers as they're displayed.

        [MSDN documentation for Workbook.PrecisionAsDisplayed](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.precisionasdisplayed).
        */
        void SetPrecisionAsDisplayed(bool precisionAsDisplayed);

        /**
        True if the order of the sheets in the workbook is protected.

        [MSDN documentation for Workbook.ProtectStructure](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.protectstructure).
        */
        bool GetProtectStructure();

        /**
        True if the windows of the workbook are protected.

        [MSDN documentation for Workbook.ProtectWindows](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.protectwindows).
        */
        bool GetProtectWindows();

        /**
        True if the object has been opened as read-only.

        [MSDN documentation for Workbook.ReadOnly](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.readonly).
        */
        bool GetReadOnly();
        /**
        True if the workbook was saved as read-only recommended.

        [MSDN documentation for Workbook.ReadOnlyRecommended](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.readonlyrecommended).
        */
        bool GetReadOnlyRecommended();
        /**
        True if personal information can be removed from the specified workbook. The default value is False.

        [MSDN documentation for Workbook.RemovePersonalInformation](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.removepersonalinformation).
        */
        bool GetRemovePersonalInformation();
        /**
        True if personal information can be removed from the specified workbook.

        [MSDN documentation for Workbook.RemovePersonalInformation](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.removepersonalinformation).
        */
        void SetRemovePersonalInformation(bool removePersonalInformation);

        /**
        Returns the number of times the workbook has been saved while open as a shared list. If the workbook is open in exclusive mode, this property returns 0 (zero).

        [MSDN documentation for Workbook.RevisionNumber](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.revisionnumber).
        */
        long GetRevisionNumber();

        /**
        True if no changes have been made to the specified workbook since it was last saved.

        [MSDN documentation for Workbook.Saved](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.saved).
        */
        bool GetSaved();
        /**
        True if no changes have been made to the specified workbook since it was last saved.

        [MSDN documentation for Workbook.Saved](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.saved).
        */
        void SetSaved(bool saved);

        /**
        True if Microsoft Excel saves external link values with the workbook.

        [MSDN documentation for Workbook.SaveLinkValues](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.savelinkvalues).
        */
        bool GetSaveLinkValues();
        /**
        True if Microsoft Excel saves external link values with the workbook.

        [MSDN documentation for Workbook.SaveLinkValues](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.savelinkvalues).
        */
        void SetSaveLinkValues(bool saveLinkValues);

        /**
        Returns wxExcelSheets that represents all the sheets in the specified workbook.

        [MSDN documentation for Workbook.Sheets](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.sheets).
        */
        wxExcelSheets GetSheets();

        /**
        True if the Conflict History worksheet is visible in the workbook that's open as a shared list.

        [MSDN documentation for Workbook.ShowConflictHistory](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.showconflicthistory).
        */
        bool GetShowConflictHistory();
        /**
        True if the Conflict History worksheet is visible in the workbook that's open as a shared list.

        [MSDN documentation for Workbook.ShowConflictHistory](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.showconflicthistory).
        */
        void SetShowConflictHistory(bool showConflictHistory);

        /**
        Returns a Styles collection that represents all the styles in the specified workbook.

        [MSDN documentation for Workbook.Styles](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.styles).
        */
        wxExcelStyles GetStyles();

        /**
        Returns a TableStyles collection that refers to the table styles used in the workbook.

        [MSDN documentation for Workbook.TableStyles](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.tablestyles)
        */
        wxExcelTableStyles GetTableStyles();

        /**
        True if external data references are removed when the workbook is saved as a template.

        [MSDN documentation for Workbook.TemplateRemoveExtData](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.templateremoveextdata).
        */
        bool GetTemplateRemoveExtData();
        /**
        True if external data references are removed when the workbook is saved as a template.

        [MSDN documentation for Workbook.TemplateRemoveExtData](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.templateremoveextdata).
        */
        void SetTemplateRemoveExtData(bool templateRemoveExtData);

        /**
        Returns an XlUpdateLink constant indicating a workbook's setting for updating embedded OLE links.

        [MSDN documentation for Workbook.UpdateLinks](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.updatelinks).
        */
        XlUpdateLinks GetUpdateLinks();
        /**
        Sets an XlUpdateLink constant indicating a workbook's setting for updating embedded OLE links.

        [MSDN documentation for Workbook.UpdateLinks](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.updatelinks).
        */
        void SetUpdateLinks(XlUpdateLinks updateLinks);

        /**
        True if Microsoft Excel updates remote references in for the workbook.

        [MSDN documentation for Workbook.UpdateRemoteReferences](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.updateremotereferences).
        */
        bool GetUpdateRemoteReferences();
        /**
        True if Microsoft Excel updates remote references in for the workbook.

        [MSDN documentation for Workbook.UpdateRemoteReferences](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.updateremotereferences).
        */
        void SetUpdateRemoteReferences(bool updateRemoteReferences);

        /**
        True if the Visual Basic for Applications project for the specified workbook has been digitally signed.

        [MSDN documentation for Workbook.VBASigned](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.vbasigned).
        */
        bool GetVBASigned();

        /**
        Returns a Windows collection that represents all the windows in the specified workbook.

        [MSDN documentation for Workbook.Windows](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.windows).
        */
        wxExcelWindows GetWindows();
        /**
        Returns a Worksheets collection that represents all the worksheets in the specified workbook.

        [MSDN documentation for Workbook.Worksheets](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.worksheets).
        */
        wxExcelWorksheets GetWorksheets();
        /**
        Returns a String for the write password of a workbook.

        [MSDN documentation for Workbook.WritePassword](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.writepassword).
        */
        wxString GetWritePassword();
        /**
        Sets a String for the write password of a workbook.

        [MSDN documentation for Workbook.WritePassword](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.writepassword).
        */
        void SetWritePassword(const wxString& writePassword);

        /**
        True if the workbook is write-reserved.

        [MSDN documentation for Workbook.WriteReserved](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.writereserved).
        */
        bool GetWriteReserved();
        /**
        Returns the name of the user who currently has write permission for the workbook.

        [MSDN documentation for Workbook.WriteReservedBy](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.writereservedby).
        */
        wxString GetWriteReservedBy();

        /**
        Returns "Workbook".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("Workbook"); }
    };


} // namespace wxAutoExcel

#endif //_WXAUTOEXCEL_WORKBOOK_H
