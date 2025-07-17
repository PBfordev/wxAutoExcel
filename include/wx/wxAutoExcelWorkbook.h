/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
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

        [MSDN documentation for Workbook.AcceptAllChanges](http://msdn.microsoft.com/en-us/library/bb209526.aspx).
        */
        void AcceptAllChanges(XlHighlightChangesTime* when = NULL, const wxString& who = wxEmptyString, const wxString& where = wxEmptyString);

        /**
        Activates the first window associated with the workbook.

        [MSDN documentation for Workbook.Activate](http://msdn.microsoft.com/en-us/library/bb179144.aspx).
        */
        bool Activate();

        /**
        Adds a shortcut to the workbook or hyperlink to the Favorites folder.

        [MSDN documentation for Workbook.AddToFavorites](http://msdn.microsoft.com/en-us/library/bb179150.aspx).
        */
        void AddToFavorites();

        /**
        Applies the specified theme to the current workbook.

        [MSDN documentation for Workbook.ApplyTheme](http://msdn.microsoft.com/en-us/library/bb238900.aspx).
        */
        void ApplyTheme(const wxString& fileName);

        /**
        Converts formulas linked to other Microsoft Excel sources or OLE sources to values.

        [MSDN documentation for Workbook.BreakLink](http://msdn.microsoft.com/en-us/library/bb209716.aspx).
        */
        void BreakLink(const wxString& name, XlLinkType type);

        /**
        True if Microsoft Excel can check in a specified workbook to a server.

        [MSDN documentation for Workbook.CanCheckIn](http://msdn.microsoft.com/en-us/library/bb223216.aspx).
        */
        bool CanCheckIn();

        /**
        Changes the access permissions for the workbook. This may require an updated version to be loaded from the disk.

        [MSDN documentation for Workbook.ChangeFileAccess](http://msdn.microsoft.com/en-us/library/bb223225.aspx).
        */
        void ChangeFileAccess(XlFileAccess mode, const wxString& writePassword, wxXlTribool notify);

        /**
        Changes a link from one document to another.

        [MSDN documentation for Workbook.ChangeLink](http://msdn.microsoft.com/en-us/library/bb223228.aspx).
        */
        void ChangeLink(const wxString& name, const wxString& newName, XlLinkType* type);

        /**
        Returns a workbook from a local computer to a server, and sets the local workbook to read-only so that it cannot be edited locally. Calling this method will also close the workbook.

        [MSDN documentation for Workbook.CheckIn](http://msdn.microsoft.com/en-us/library/bb223246.aspx).
        */
        void CheckIn(wxXlTribool saveChanges = wxDefaultXlTribool,
            wxXlTribool comments = wxDefaultXlTribool,
            wxXlTribool makePublic = wxDefaultXlTribool);

        /**
        Saves a workbook to a server from a local computer, and sets the local workbook to read-only so that it cannot be edited locally. Since Excel 2010

        [MSDN documentation for Workbook.CheckInWithVersion](http://msdn.microsoft.com/en-us/library/office/ff196878%28v=office.14%29.aspx).
        */
        void CheckInWithVersion(wxXlTribool saveChanges = wxDefaultXlTribool,
            wxXlTribool comments = wxDefaultXlTribool, wxXlTribool makePublic = wxDefaultXlTribool,
            XlCheckInVersionType* versionType = NULL);

        /**

        Closes the object.

        [MSDN documentation for Workbook.Close](http://msdn.microsoft.com/en-us/library/bb179153.aspx).
        */
        bool Close(wxXlTribool saveChanges = wxDefaultXlTribool, const wxString& fileName = wxEmptyString, wxXlTribool routeWorkbook = wxDefaultXlTribool);

        /**
        Deletes a custom number format from the workbook.

        [MSDN documentation for Workbook.DeleteNumberFormat](http://msdn.microsoft.com/en-us/library/bb223329.aspx).
        */
        void DeleteNumberFormat(const wxString& numberFormat);

        /**
        The EnableConnections method allows developers to programmatically enable data connections within the workbook for the user.

        [MSDN documentation for Workbook.EnableConnections](http://msdn.microsoft.com/en-us/library/bb238905.aspx).
        */
        void EnableConnections();

        /**
        Terminates a review of a file that has been sent for review using the SendForReview method.

        [MSDN documentation for Workbook.EndReview](http://msdn.microsoft.com/en-us/library/bb209810.aspx).
        */
        void EndReview();

        /**
        Assigns the current user exclusive access to the workbook that's open as a shared list.

        [MSDN documentation for Workbook.ExclusiveAccess](http://msdn.microsoft.com/en-us/library/bb209820.aspx).
        */
        bool ExclusiveAccess();

        //@{
        /**
        The ExportAsFixedFormat method is used to publish a workbook to either the PDF or XPS format.

        [MSDN documentation for Workbook.ExportAsFixedFormat](http://msdn.microsoft.com/en-us/library/bb238907.aspx).
        */
        void ExportAsFixedFormat(XlFixedFormatType type, const wxString& fileName = wxEmptyString,
            XlFixedFormatQuality* quality = NULL, wxXlTribool includeDocProperties = wxDefaultXlTribool,
            wxXlTribool ignorePrintAreas = wxDefaultXlTribool,
            long* from = NULL, long* to = NULL, wxXlTribool openAfterPublish = wxDefaultXlTribool);

        void ExportAsFixedFormat(XlFixedFormatType type, const wxVariantVector& optionalArgs);
        //@}

        /**
        Controls how changes are shown in a shared workbook.

        [MSDN documentation for Workbook.HighlightChangesOptions](http://msdn.microsoft.com/en-us/library/bb209922.aspx).
        */
        void HighlightChangesOptions(XlHighlightChangesTime* when = NULL, const wxString& who = wxEmptyString, const wxString& where = wxEmptyString);

        /**
        Locks the workbook on the server to prevent modification.

        [MSDN documentation for Workbook.LockServerFile](http://msdn.microsoft.com/en-us/library/bb238915.aspx).
        */
        void LockServerFile();

        /**
        Merges changes from one workbook into an open workbook.

        [MSDN documentation for Workbook.MergeWorkbook](http://msdn.microsoft.com/en-us/library/bb223481.aspx).
        */
        void MergeWorkbook(const wxString& fileName);

        /**
        Creates a new window or a copy of the specified window.

        [MSDN documentation for Workbook.NewWindow](http://msdn.microsoft.com/en-us/library/bb179154.aspx).
        */
        wxExcelWindow NewWindow();

        /**
        Posts the specified workbook to a public folder. This method works only with a Microsoft Exchange client connected to a Microsoft Exchange server.

        [MSDN documentation for Workbook.Post](http://msdn.microsoft.com/en-us/library/bb223546.aspx).
        */
        void Post();

        //@{
        /**
        Prints the object.

        [MSDN documentation for Workbook.PrintOut](http://msdn.microsoft.com/en-us/library/bb179158.aspx).
        */
        bool PrintOut(long* from = NULL, long* to = NULL, long* copies = NULL, wxXlTribool preview = wxDefaultXlTribool,
            const wxString& activePrinter = wxEmptyString, wxXlTribool printToFile = wxDefaultXlTribool,
            wxXlTribool collate = wxDefaultXlTribool, const wxString& prToFileName= wxEmptyString, wxXlTribool ignorePrintAreas = wxDefaultXlTribool);
        bool PrintOut(const wxVariantVector& args);
        //@}

        /**
        Shows a preview of the object as it would look when printed.

        [MSDN documentation for Workbook.PrintPreview](http://msdn.microsoft.com/en-us/library/bb179160.aspx).
        */
        bool PrintPreview(wxXlTribool enableChanges = wxDefaultXlTribool);

        /**
        Protects a workbook so that it cannot be modified.

        [MSDN documentation for Workbook.Protect](http://msdn.microsoft.com/en-us/library/bb179162.aspx).
        */
        void Protect(const wxString& password = wxEmptyString, wxXlTribool structure = wxDefaultXlTribool, wxXlTribool windows = wxDefaultXlTribool);

        //@{
        /**
        Saves the workbook and protects it for sharing.

        [MSDN documentation for Workbook.ProtectSharing](http://msdn.microsoft.com/en-us/library/bb223552.aspx).
        */
        void ProtectSharing(const wxString& fileName = wxEmptyString, const wxString& password = wxEmptyString,
            const wxString& writeResPassword = wxEmptyString,
            wxXlTribool readOnlyRecommended = wxDefaultXlTribool, wxXlTribool createBackup = wxDefaultXlTribool,
            const wxString& sharingPassword = wxEmptyString);

        void ProtectSharing(const wxVariantVector& optionalArgs);
        //@}

        /**
        Removes entries from the change log for the specified workbook.

        [MSDN documentation for Workbook.PurgeChangeHistoryNow](http://msdn.microsoft.com/en-us/library/bb223556.aspx).
        */
        void PurgeChangeHistoryNow(long days, const wxString& sharingPassword = wxEmptyString);

        /**
        Refreshes all external data ranges and PivotTable reports in the specified workbook.

        [MSDN documentation for Workbook.RefreshAll](http://msdn.microsoft.com/en-us/library/bb223574.aspx).
        */
        void RefreshAll();

        /**
        Rejects all changes in the specified shared workbook.

        [MSDN documentation for Workbook.RejectAllChanges](http://msdn.microsoft.com/en-us/library/bb223585.aspx).
        */
        void RejectAllChanges(XlHighlightChangesTime* when = NULL, const wxString& who = wxEmptyString, const wxString& where = wxEmptyString);

        /**
        Reloads a workbook based on an HTML document, using the specified document encoding.

        [MSDN documentation for Workbook.ReloadAs](http://msdn.microsoft.com/en-us/library/bb223588.aspx).
        */
        void ReloadAs(MsoEncoding encoding);

        /**
        Removes all information of the specified type from the workbook.

        [MSDN documentation for Workbook.RemoveDocumentInformation](http://msdn.microsoft.com/en-us/library/bb242484.aspx).
        */
        void RemoveDocumentInformation(XlRemoveDocInfoType removeDocInfoType);

        /**
        Disconnects the specified user from the shared workbook.

        [MSDN documentation for Workbook.RemoveUser](http://msdn.microsoft.com/en-us/library/bb223600.aspx).
        */
        void RemoveUser(long index);

        /**
        Sends an e-mail message to the author of a workbook that has been sent out for review, notifying them that a reviewer has completed review of the workbook.

        [MSDN documentation for Workbook.ReplyWithChanges](http://msdn.microsoft.com/en-us/library/bb177952.aspx).
        */
        void ReplyWithChanges(wxXlTribool showMessage = wxDefaultXlTribool);

        /**
        Resets the color palette to the default colors.

        [MSDN documentation for Workbook.ResetColors](http://msdn.microsoft.com/en-us/library/bb177963.aspx).
        */
        void ResetColors();

        /**
        Runs the Auto_Open, Auto_Close, Auto_Activate, or Auto_Deactivate macro attached to the workbook.

        [MSDN documentation for Workbook.RunAutoMacros](http://msdn.microsoft.com/en-us/library/bb177988.aspx).
        */
        void RunAutoMacros(XlRunAutoMacro which);

        /**
        Saves changes to the specified workbook.

        [MSDN documentation for Workbook.Save](http://msdn.microsoft.com/en-us/library/bb177993.aspx).
        */
        bool Save();

        //@{
        /**
        Saves changes to the workbook in a different file.

        [MSDN documentation for Workbook.SaveAs](http://msdn.microsoft.com/en-us/library/bb214129.aspx).
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

        [MSDN documentation for Workbook.SaveCopyAs](http://msdn.microsoft.com/en-us/library/bb178003.aspx).
        */
        bool SaveCopyAs(const wxString& fileName);

        /**
        Sends a worksheet as a fax to the specfied recipients.

        [MSDN documentation for Workbook.SendFaxOverInternet](http://msdn.microsoft.com/en-us/library/bb178019.aspx).
        */
        void SendFaxOverInternet(const wxString& recipients = wxEmptyString, const wxString& subject= wxEmptyString,
            wxXlTribool showMessage = wxDefaultXlTribool);

        /**
        Sends a workbook in an e-mail message for review to the specified recipients.

        [MSDN documentation for Workbook.SendForReview](http://msdn.microsoft.com/en-us/library/bb178022.aspx).
        */
        void SendForReview(const wxString& recipients = wxEmptyString, const wxString& subject= wxEmptyString,
            wxXlTribool showMessage = wxDefaultXlTribool, wxXlTribool includeAttachment = wxDefaultXlTribool);

        /**
        Sends the workbook by using the installed mail system.

        [MSDN documentation for Workbook.SendMail](http://msdn.microsoft.com/en-us/library/bb178034.aspx).
        */
        void SendMail(const wxString& recipients = wxEmptyString, const wxString& subject= wxEmptyString,
            wxXlTribool returnReceipt = wxDefaultXlTribool);

        /**
        Sets the options for encrypting workbooks using passwords.

        [MSDN documentation for Workbook.SetPasswordEncryptionOptions](http://msdn.microsoft.com/en-us/library/bb178065.aspx).
        */
        void SetPasswordEncryptionOptions(const wxString& passwordEncryptionProvider = wxEmptyString,
            const wxString& passwordEncryptionAlgorithm  = wxEmptyString,
            long* passwordEncryptionKeyLength = NULL,
            wxXlTribool passwordEncryptionFileProperties = wxDefaultXlTribool);

        /**
        Removes protection from a sheet or workbook. This method has no effect if the sheet or workbook isn't protected.

        [MSDN documentation for Workbook.Unprotect](http://msdn.microsoft.com/en-us/library/bb214137.aspx).
        */
        bool Unprotect(const wxString& password = wxEmptyString);

        /**
        Turns off protection for sharing and saves the workbook.

        [MSDN documentation for Workbook.UnprotectSharing](http://msdn.microsoft.com/en-us/library/bb210015.aspx).
        */
        void UnprotectSharing(const wxString& sharingPassword = wxEmptyString);

        /**
        Updates a read-only workbook from the saved disk version of the workbook if the disk version is more recent than the copy of the workbook that is loaded in memory. If the disk copy hasn't changed since the workbook was loaded, the in-memory copy of the workbook isn't reloaded.

        [MSDN documentation for Workbook.UpdateFromFile](http://msdn.microsoft.com/en-us/library/bb210018.aspx).
        */
        void UpdateFromFile();

        /**
        Displays a preview of the specified workbook as it would look if saved as a Web page.

        [MSDN documentation for Workbook.WebPagePreview](http://msdn.microsoft.com/en-us/library/bb210042.aspx).
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

        [MSDN documentation for Workbook.ActiveSheet](http://msdn.microsoft.com/en-us/library/bb148768.aspx).
        */
        wxExcelSheet GetActiveSheet();

        /**
        Specifies whether certain worksheet functions use the latest accuracy algorithms to calculate their results. Since Excel 2010.

        [MSDN documentation for Worksheet.AccuracyVersion](http://msdn.microsoft.com/en-us/library/bb148863.aspx).
        */
        long GetAccuracyVersion();

        /**
        Specifies whether certain worksheet functions use the latest accuracy algorithms to calculate their results. Since Excel 2010.

        [MSDN documentation for Worksheet.AccuracyVersion](http://msdn.microsoft.com/en-us/library/bb148863.aspx).
        */
        void SetAccuracyVersion(long accuracyVersion);

        /**
        Returns the number of minutes between automatic updates to the shared workbook.

        [MSDN documentation for Workbook.AutoUpdateFrequency](http://msdn.microsoft.com/en-us/library/bb220869.aspx).
        */
        long GetAutoUpdateFrequency();
        /**
        Sets the number of minutes between automatic updates to the shared workbook.

        [MSDN documentation for Workbook.AutoUpdateFrequency](http://msdn.microsoft.com/en-us/library/bb220869.aspx).
        */
        void SetAutoUpdateFrequency(long autoUpdateFrequency);

        /**
        True if current changes to the shared workbook are posted to other users whenever the workbook is automatically updated. False if changes aren't posted (this workbook is still synchronized with changes made by other users). The default value is True.

        [MSDN documentation for Workbook.AutoUpdateSaveChanges](http://msdn.microsoft.com/en-us/library/bb220872.aspx).
        */
        bool GetAutoUpdateSaveChanges();
        /**
        True if current changes to the shared workbook are posted to other users whenever the workbook is automatically updated. False if changes aren't posted (this workbook is still synchronized with changes made by other users). The default value is True.

        [MSDN documentation for Workbook.AutoUpdateSaveChanges](http://msdn.microsoft.com/en-us/library/bb220872.aspx).
        */
        void SetAutoUpdateSaveChanges(bool autoUpdateSaveChanges);


        /**
        Returns a DocumentProperties collection that represents all the built-in document properties.

        [MSDN documentation for Workbook.BuiltinDocumentProperties](http://msdn.microsoft.com/en-us/library/bb220896.aspx).
        */
        wxExcelDocumentProperties GetBuiltinDocumentProperties();

        /**
        Returns the information about the version of Excel that the workbook was last fully recalculated by.

        [MSDN documentation for Workbook.CalculationVersion](http://msdn.microsoft.com/en-us/library/bb148771.aspx).
        */
        long GetCalculationVersion();
        /**
        Returns the number of days shown in the shared workbook's change history.

        [MSDN documentation for Workbook.ChangeHistoryDuration](http://msdn.microsoft.com/en-us/library/bb177339.aspx).
        */
        long GetChangeHistoryDuration();
        /**
        Sets the number of days shown in the shared workbook's change history.

        [MSDN documentation for Workbook.ChangeHistoryDuration](http://msdn.microsoft.com/en-us/library/bb177339.aspx).
        */
        void SetChangeHistoryDuration(long changeHistoryDuration);

#if WXAUTOEXCEL_USE_CHARTS
        /**
        Returns a Sheets collection that represents all the chart sheets in the specified workbook.

        [MSDN documentation for Workbook.Charts](http://msdn.microsoft.com/en-us/library/bb148780.aspx).
        */
        wxExcelCharts GetCharts();
#endif // #if WXAUTOEXCEL_USE_CHARTS

        /**
        Controls whether or not the compatibility checker is run automatically when the workbook is saved. Since MS Excel 2007.

        [MSDN documentation for Workbook.CheckCompatibility](http://msdn.microsoft.com/en-us/library/bb216059.aspx).
        */
        bool GetCheckCompatibility();
        /**
        Controls whether or not the compatibility checker is run automatically when the workbook is saved. Since MS Excel 2007.

        [MSDN documentation for Workbook.CheckCompatibility](http://msdn.microsoft.com/en-us/library/bb216059.aspx).
        */
        void SetCheckCompatibility(bool checkCompatibility);

        //@{
        /**
        Returns colors in the palette for the workbook. The palette has 56 entries, each represented by an RGB value.

        [MSDN documentation for Workbook.Colors](http://msdn.microsoft.com/en-us/library/bb177361.aspx).
        */
        wxColour GetColors(long index);
        wxVector<wxColour> GetColors();
        //@}

        //@{
        /**
        Sets colors in the palette for the workbook. The palette has 56 entries, each represented by an RGB value.

        [MSDN documentation for Workbook.Colors](http://msdn.microsoft.com/en-us/library/bb177361.aspx).
        */
        void SetColors(long index, const wxColour& color);
        void SetColors(const wxVector<wxColour>& colors);
        //@}

        /**
        [MSDN documentation for Workbook.ConflictResolution](http://msdn.microsoft.com/en-us/library/bb177382.aspx).
        */
        XlSaveConflictResolution GetConflictResolution();
        /**
        Sets the way conflicts are to be resolved whenever a shared workbook is updated. Read/write XlSaveConflictResolution.

        [MSDN documentation for Workbook.ConflictResolution](http://msdn.microsoft.com/en-us/library/bb177382.aspx).
        */
        void SetConflictResolution(XlSaveConflictResolution conflictResolution);

        /**
        Disables the external connections or links in the workbook. Since MS Excel 2007.

        [MSDN documentation for Workbook.ConnectionsDisabled](http://msdn.microsoft.com/en-us/library/bb257126.aspx).
        */
        bool GetConnectionsDisabled();

        /**
        True if a backup file is created when this file is saved.

        [MSDN documentation for Workbook.CreateBackup](http://msdn.microsoft.com/en-us/library/bb177400.aspx).
        */
        bool GetCreateBackup();

         /**
        Returns a DocumentProperties collection that represents all the custom document properties.

        [MSDN documentation for Workbook.CustomDocumentProperties](http://msdn.microsoft.com/en-us/library/bb177423.aspx).
        */
        wxExcelDocumentProperties GetCustomDocumentProperties();

        /**
        True if the workbook uses the 1904 date system.

        [MSDN documentation for Workbook.Date1904](http://msdn.microsoft.com/en-us/library/bb177448.aspx).
        */
        bool GetDate1904();
        /**
        True if the workbook uses the 1904 date system.

        [MSDN documentation for Workbook.Date1904](http://msdn.microsoft.com/en-us/library/bb177448.aspx).
        */
        void SetDate1904(bool date1904);

        /**
        A Boolean value that determines whether ink comments are displayed in the workbook.

        [MSDN documentation for Workbook.DisplayInkComments](http://msdn.microsoft.com/en-us/library/bb177510.aspx).
        */
        bool GetDisplayInkComments();
        /**
        A Boolean value that determines whether ink comments are displayed in the workbook.

        [MSDN documentation for Workbook.DisplayInkComments](http://msdn.microsoft.com/en-us/library/bb177510.aspx).
        */
        void SetDisplayInkComments(bool displayInkComments);

        /**
        Returns if the user should be prompted to convert the workbook if the workbook contains features that are not supported by versions of Excel earlier than Excel 2007. Since MS Excel 2007.

        [MSDN documentation for Workbook.DoNotPromptForConvert](http://msdn.microsoft.com/en-us/library/bb226049.aspx).
        */
        bool GetDoNotPromptForConvert();
        /**
        Sets if the user should be prompted to convert the workbook if the workbook contains features that are not supported by versions of Excel earlier than Excel 2007. Since MS Excel 2007.

        [MSDN documentation for Workbook.DoNotPromptForConvert](http://msdn.microsoft.com/en-us/library/bb226049.aspx).
        */
        void SetDoNotPromptForConvert(bool doNotPromptForConvert);

        /**
        Saves changed files, of all formats, on a timed interval.

        [MSDN documentation for Workbook.EnableAutoRecover](http://msdn.microsoft.com/en-us/library/bb221131.aspx).
        */
        bool GetEnableAutoRecover();
        /**
        Saves changed files, of all formats, on a timed interval.

        [MSDN documentation for Workbook.EnableAutoRecover](http://msdn.microsoft.com/en-us/library/bb221131.aspx).
        */
        void SetEnableAutoRecover(bool enableAutoRecover);

        /**
        Returns a string specifying the name of the algorithm encryption provider that Excel uses when encrypting documents. Since MS Excel 2007.

        [MSDN documentation for Workbook.EncryptionProvider](http://msdn.microsoft.com/en-us/library/bb242647.aspx).
        */
        wxString GetEncryptionProvider();
        /**
        Sets the name of the algorithm encryption provider that Excel uses when encrypting documents. Since MS Excel 2007.

        [MSDN documentation for Workbook.EncryptionProvider](http://msdn.microsoft.com/en-us/library/bb242647.aspx).
        */
        void SetEncryptionProvider(const wxString& encryptionProvider);

        /**
        True if the e-mail composition header and the envelope toolbar are both visible.

        [MSDN documentation for Workbook.EnvelopeVisible](http://msdn.microsoft.com/en-us/library/bb208468.aspx).
        */
        bool GetEnvelopeVisible();
        /**
        True if the e-mail composition header and the envelope toolbar are both visible.

        [MSDN documentation for Workbook.EnvelopeVisible](http://msdn.microsoft.com/en-us/library/bb208468.aspx).
        */
        void SetEnvelopeVisible(bool envelopeVisible);

        /**
        The Excel8CompatibilityMode property provides developers with a way to check if the workbook is in compatibility mode. Since MS Excel 2007.

        [MSDN documentation for Workbook.Excel8CompatibilityMode](http://msdn.microsoft.com/en-us/library/bb216076.aspx).
        */
        bool GetExcel8CompatibilityMode();

        /**
        Returns the file format and/or type of the workbook.

        [MSDN documentation for Workbook.FileFormat](http://msdn.microsoft.com/en-us/library/bb208506.aspx).
        */
        XlFileFormat GetFileFormat();

        /**
        Returns a Boolean that indicates whether a workbook is final. Since MS Excel 2007.

        [MSDN documentation for Workbook.Final](http://msdn.microsoft.com/en-us/library/bb216084.aspx).
        */
        bool GetFinal();
        /**
        Sets a Boolean that indicates whether a workbook is final.

        [MSDN documentation for Workbook.Final](http://msdn.microsoft.com/en-us/library/bb216084.aspx).
        */
        void SetFinal(bool final);

        /**
        Forces a full calculation of a workbook. Since MS Excel 2007.

        [MSDN documentation for Workbook.ForceFullCalculation](http://msdn.microsoft.com/en-us/library/bb257030.aspx).
        */
        bool GetForceFullCalculation();
        /**
        Forces a full calculation of a workbook. Since MS Excel 2007.

        [MSDN documentation for Workbook.ForceFullCalculation](http://msdn.microsoft.com/en-us/library/bb257030.aspx).
        */
        void SetForceFullCalculation(bool forceFullCalculation);

        /**
        Returns the workbook name including its path on disk.

        [MSDN documentation for Workbook.FullName](http://msdn.microsoft.com/en-us/library/bb148805.aspx).
        */
        wxString GetFullName();

        /**
        Returns the workbook name including its path on disk.

        [MSDN documentation for Workbook.FullNameURLEncoded](http://msdn.microsoft.com/en-us/library/bb208565.aspx).
        */
        wxString GetFullNameURLEncoded();

        /**
        True if the workbook has a protection password.

        [MSDN documentation for Workbook.HasPassword](http://msdn.microsoft.com/en-us/library/bb208650.aspx).
        */
        bool GetHasPassword();

        /**
        True if the workbook has an attached Microsoft Visual Basic for Applications project. Since MS Excel 2007.

        [MSDN documentation for Workbook.HasVBProject](http://msdn.microsoft.com/en-us/library/bb216182.aspx).
        */
        bool GetHasVBProject();

        /**
        True if changes to the shared workbook are highlighted on-screen.

        [MSDN documentation for Workbook.HighlightChangesOnScreen](http://msdn.microsoft.com/en-us/library/bb208694.aspx).
        */
        bool GetHighlightChangesOnScreen();
        /**
        True if changes to the shared workbook are highlighted on-screen.

        [MSDN documentation for Workbook.HighlightChangesOnScreen](http://msdn.microsoft.com/en-us/library/bb208694.aspx).
        */
        void SetHighlightChangesOnScreen(bool highlightChangesOnScreen);

#if WXAUTOEXCEL_USE_CONDFORMAT
        /**
        This property is used to filter data in a workbook based on a cell icon from the IconSet collection.

        [MSDN documentation for Workbook.IconSets](http://msdn.microsoft.com/en-us/library/bb216187.aspx).
        */
        wxExcelIconSets GetIconSets(XlIconSet index);
#endif // #if WXAUTOEXCEL_USE_CONDFORMAT

        /**
        True if the list borders are visible when a list is not active.

        [MSDN documentation for Workbook.InactiveListBorderVisible](http://msdn.microsoft.com/en-us/library/bb177606.aspx).
        */
        bool GetInactiveListBorderVisible();
        /**
        True if the list borders are visible when a list is not active.

        [MSDN documentation for Workbook.InactiveListBorderVisible](http://msdn.microsoft.com/en-us/library/bb177606.aspx).
        */
        void SetInactiveListBorderVisible(bool inactiveListBorderVisible);

        /**
        True if the workbook is running as an add-in.

        [MSDN documentation for Workbook.IsAddin](http://msdn.microsoft.com/en-us/library/bb177681.aspx).
        */
        bool GetIsAddin();
        /**
        True if the workbook is running as an add-in.

        [MSDN documentation for Workbook.IsAddin](http://msdn.microsoft.com/en-us/library/bb177681.aspx).
        */
        void SetIsAddin(bool isAddin);

        /**
        True if the specified workbook is being edited in place. False if the workbook has been opened in Microsoft Excel for editing.

        [MSDN documentation for Workbook.IsInplace](http://msdn.microsoft.com/en-us/library/bb177690.aspx).
        */
        bool GetIsInplace();

        /**
        True if change tracking is enabled for the shared workbook.

        [MSDN documentation for Workbook.KeepChangeHistory](http://msdn.microsoft.com/en-us/library/bb177811.aspx).
        */
        bool GetKeepChangeHistory();
        /**
        True if change tracking is enabled for the shared workbook.

        [MSDN documentation for Workbook.KeepChangeHistory](http://msdn.microsoft.com/en-us/library/bb177811.aspx).
        */
        void SetKeepChangeHistory(bool keepChangeHistory);

        /**
        True if changes to the shared workbook are shown on a separate worksheet.

        [MSDN documentation for Workbook.ListChangesOnNewSheet](http://msdn.microsoft.com/en-us/library/bb177912.aspx).
        */
        bool GetListChangesOnNewSheet();
        /**
        True if changes to the shared workbook are shown on a separate worksheet.

        [MSDN documentation for Workbook.ListChangesOnNewSheet](http://msdn.microsoft.com/en-us/library/bb177912.aspx).
        */
        void SetListChangesOnNewSheet(bool listChangesOnNewSheet);

        /**
        True if the workbook is open as a shared list.

        [MSDN documentation for Workbook.MultiUserEditing](http://msdn.microsoft.com/en-us/library/bb208810.aspx).
        */
        bool GetMultiUserEditing();

        /**
        Returns the name of the object.

        [MSDN documentation for Workbook.Name](https://msdn.microsoft.com/en-us/library/office/ff820899.aspx).
        */
        wxString GetName();

        /**
        Returns a Names collection that represents all the names in the specified workbook (including all worksheet-specific names).

        [MSDN documentation for Workbook.Names](https://msdn.microsoft.com/en-us/library/office/ff195422.aspx).
        */
        wxExcelNames GetNames();

        /**
        Returns the password that must be supplied to open the specified workbook.

        [MSDN documentation for Workbook.Password](http://msdn.microsoft.com/en-us/library/bb221191.aspx).
        */
        wxString GetPassword();
        /**
        Sets the password that must be supplied to open the specified workbook.

        [MSDN documentation for Workbook.Password](http://msdn.microsoft.com/en-us/library/bb221191.aspx).
        */
        void SetPassword(const wxString& password);

        /**
        Returns a string indicating the algorithm Microsoft Excel uses to encrypt passwords for the specified workbook.

        [MSDN documentation for Workbook.PasswordEncryptionAlgorithm](http://msdn.microsoft.com/en-us/library/bb221197.aspx).
        */
        wxString GetPasswordEncryptionAlgorithm();
        /**
        True if Microsoft Excel encrypts file properties for the specified password-protected workbook.

        [MSDN documentation for Workbook.PasswordEncryptionFileProperties](http://msdn.microsoft.com/en-us/library/bb221202.aspx).
        */
        bool GetPasswordEncryptionFileProperties();

        /**
        Returns the key length of the algorithm Microsoft Excel uses when encrypting passwords for the specified workbook.

        [MSDN documentation for Workbook.PasswordEncryptionKeyLength](http://msdn.microsoft.com/en-us/library/bb221207.aspx).
        */
        long GetPasswordEncryptionKeyLength();

        /**
        Returns the name of the algorithm encryption provider that Microsoft Excel uses when encrypting passwords for the specified workbook.

        [MSDN documentation for Workbook.PasswordEncryptionProvider](http://msdn.microsoft.com/en-us/library/bb221213.aspx).
        */
        wxString GetPasswordEncryptionProvider();

        /**
        Returns the complete path to the workbook/file that this workbook object respresents.

        [MSDN documentation for Workbook.Path](http://msdn.microsoft.com/en-us/library/bb148817.aspx).
        */
        wxString GetPath();

        /**
        True if filter and sort settings for lists are included in the user's personal view of the shared workbook.

        [MSDN documentation for Workbook.PersonalViewListSettings](http://msdn.microsoft.com/en-us/library/bb221417.aspx).
        */
        bool GetPersonalViewListSettings();
        /**
        True if filter and sort settings for lists are included in the user's personal view of the shared workbook.

        [MSDN documentation for Workbook.PersonalViewListSettings](http://msdn.microsoft.com/en-us/library/bb221417.aspx).
        */
        void SetPersonalViewListSettings(bool personalViewListSettings);

        /**
        True if print settings are included in the user's personal view of the shared workbook.

        [MSDN documentation for Workbook.PersonalViewPrintSettings](http://msdn.microsoft.com/en-us/library/bb221418.aspx).
        */
        bool GetPersonalViewPrintSettings();
        /**
        True if print settings are included in the user's personal view of the shared workbook.

        [MSDN documentation for Workbook.PersonalViewPrintSettings](http://msdn.microsoft.com/en-us/library/bb221418.aspx).
        */
        void SetPersonalViewPrintSettings(bool personalViewPrintSettings);

        /**
        True if calculations in this workbook will be done using only the precision of the numbers as they're displayed.

        [MSDN documentation for Workbook.PrecisionAsDisplayed](http://msdn.microsoft.com/en-us/library/bb208965.aspx).
        */
        bool GetPrecisionAsDisplayed();
        /**
        True if calculations in this workbook will be done using only the precision of the numbers as they're displayed.

        [MSDN documentation for Workbook.PrecisionAsDisplayed](http://msdn.microsoft.com/en-us/library/bb208965.aspx).
        */
        void SetPrecisionAsDisplayed(bool precisionAsDisplayed);

        /**
        True if the order of the sheets in the workbook is protected.

        [MSDN documentation for Workbook.ProtectStructure](http://msdn.microsoft.com/en-us/library/bb209045.aspx).
        */
        bool GetProtectStructure();

        /**
        True if the windows of the workbook are protected.

        [MSDN documentation for Workbook.ProtectWindows](http://msdn.microsoft.com/en-us/library/bb209046.aspx).
        */
        bool GetProtectWindows();

        /**
        True if the object has been opened as read-only.

        [MSDN documentation for Workbook.ReadOnly](http://msdn.microsoft.com/en-us/library/bb148821.aspx).
        */
        bool GetReadOnly();
        /**
        True if the workbook was saved as read-only recommended.

        [MSDN documentation for Workbook.ReadOnlyRecommended](http://msdn.microsoft.com/en-us/library/bb209057.aspx).
        */
        bool GetReadOnlyRecommended();
        /**
        True if personal information can be removed from the specified workbook. The default value is False.

        [MSDN documentation for Workbook.RemovePersonalInformation](http://msdn.microsoft.com/en-us/library/bb209103.aspx).
        */
        bool GetRemovePersonalInformation();
        /**
        True if personal information can be removed from the specified workbook.

        [MSDN documentation for Workbook.RemovePersonalInformation](http://msdn.microsoft.com/en-us/library/bb209103.aspx).
        */
        void SetRemovePersonalInformation(bool removePersonalInformation);

        /**
        Returns the number of times the workbook has been saved while open as a shared list. If the workbook is open in exclusive mode, this property returns 0 (zero).

        [MSDN documentation for Workbook.RevisionNumber](http://msdn.microsoft.com/en-us/library/bb209156.aspx).
        */
        long GetRevisionNumber();

        /**
        True if no changes have been made to the specified workbook since it was last saved.

        [MSDN documentation for Workbook.Saved](http://msdn.microsoft.com/en-us/library/bb221578.aspx).
        */
        bool GetSaved();
        /**
        True if no changes have been made to the specified workbook since it was last saved.

        [MSDN documentation for Workbook.Saved](http://msdn.microsoft.com/en-us/library/bb221578.aspx).
        */
        void SetSaved(bool saved);

        /**
        True if Microsoft Excel saves external link values with the workbook.

        [MSDN documentation for Workbook.SaveLinkValues](http://msdn.microsoft.com/en-us/library/bb221590.aspx).
        */
        bool GetSaveLinkValues();
        /**
        True if Microsoft Excel saves external link values with the workbook.

        [MSDN documentation for Workbook.SaveLinkValues](http://msdn.microsoft.com/en-us/library/bb221590.aspx).
        */
        void SetSaveLinkValues(bool saveLinkValues);

        /**
        Returns wxExcelSheets that represents all the sheets in the specified workbook.

        [MSDN documentation for Workbook.Sheets](http://msdn.microsoft.com/en-us/library/bb215242.aspx).
        */
        wxExcelSheets GetSheets();

        /**
        True if the Conflict History worksheet is visible in the workbook that's open as a shared list.

        [MSDN documentation for Workbook.ShowConflictHistory](http://msdn.microsoft.com/en-us/library/bb221685.aspx).
        */
        bool GetShowConflictHistory();
        /**
        True if the Conflict History worksheet is visible in the workbook that's open as a shared list.

        [MSDN documentation for Workbook.ShowConflictHistory](http://msdn.microsoft.com/en-us/library/bb221685.aspx).
        */
        void SetShowConflictHistory(bool showConflictHistory);

        /**
        Returns a Styles collection that represents all the styles in the specified workbook.

        [MSDN documentation for Workbook.Styles](http://msdn.microsoft.com/en-us/library/bb209305.aspx).
        */
        wxExcelStyles GetStyles();

        /**
        Returns a TableStyles collection that refers to the table styles used in the workbook.

        [MSDN documentation for Workbook.TableStyles](https://docs.microsoft.com/en-us/office/vba/api/excel.workbook.tablestyles)
        */
        wxExcelTableStyles GetTableStyles();

        /**
        True if external data references are removed when the workbook is saved as a template.

        [MSDN documentation for Workbook.TemplateRemoveExtData](http://msdn.microsoft.com/en-us/library/bb221717.aspx).
        */
        bool GetTemplateRemoveExtData();
        /**
        True if external data references are removed when the workbook is saved as a template.

        [MSDN documentation for Workbook.TemplateRemoveExtData](http://msdn.microsoft.com/en-us/library/bb221717.aspx).
        */
        void SetTemplateRemoveExtData(bool templateRemoveExtData);

        /**
        Returns an XlUpdateLink constant indicating a workbook's setting for updating embedded OLE links.

        [MSDN documentation for Workbook.UpdateLinks](http://msdn.microsoft.com/en-us/library/bb221949.aspx).
        */
        XlUpdateLinks GetUpdateLinks();
        /**
        Sets an XlUpdateLink constant indicating a workbook's setting for updating embedded OLE links.

        [MSDN documentation for Workbook.UpdateLinks](http://msdn.microsoft.com/en-us/library/bb221949.aspx).
        */
        void SetUpdateLinks(XlUpdateLinks updateLinks);

        /**
        True if Microsoft Excel updates remote references in for the workbook.

        [MSDN documentation for Workbook.UpdateRemoteReferences](http://msdn.microsoft.com/en-us/library/bb221958.aspx).
        */
        bool GetUpdateRemoteReferences();
        /**
        True if Microsoft Excel updates remote references in for the workbook.

        [MSDN documentation for Workbook.UpdateRemoteReferences](http://msdn.microsoft.com/en-us/library/bb221958.aspx).
        */
        void SetUpdateRemoteReferences(bool updateRemoteReferences);

        /**
        True if the Visual Basic for Applications project for the specified workbook has been digitally signed.

        [MSDN documentation for Workbook.VBASigned](http://msdn.microsoft.com/en-us/library/bb223012.aspx).
        */
        bool GetVBASigned();

        /**
        Returns a Windows collection that represents all the windows in the specified workbook.

        [MSDN documentation for Workbook.Windows](http://msdn.microsoft.com/en-us/library/bb215244.aspx).
        */
        wxExcelWindows GetWindows();
        /**
        Returns a Worksheets collection that represents all the worksheets in the specified workbook.

        [MSDN documentation for Workbook.Worksheets](http://msdn.microsoft.com/en-us/library/bb215247.aspx).
        */
        wxExcelWorksheets GetWorksheets();
        /**
        Returns a String for the write password of a workbook.

        [MSDN documentation for Workbook.WritePassword](http://msdn.microsoft.com/en-us/library/bb223071.aspx).
        */
        wxString GetWritePassword();
        /**
        Sets a String for the write password of a workbook.

        [MSDN documentation for Workbook.WritePassword](http://msdn.microsoft.com/en-us/library/bb223071.aspx).
        */
        void SetWritePassword(const wxString& writePassword);

        /**
        True if the workbook is write-reserved.

        [MSDN documentation for Workbook.WriteReserved](http://msdn.microsoft.com/en-us/library/bb223074.aspx).
        */
        bool GetWriteReserved();
        /**
        Returns the name of the user who currently has write permission for the workbook.

        [MSDN documentation for Workbook.WriteReservedBy](http://msdn.microsoft.com/en-us/library/bb223076.aspx).
        */
        wxString GetWriteReservedBy();

        /**
        Returns "Workbook".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("Workbook"); }
    };


} // namespace wxAutoExcel

#endif //_WXAUTOEXCEL_WORKBOOK_H
