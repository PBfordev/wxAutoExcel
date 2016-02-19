/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Modified by:
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////


#include "wx/wxAutoExcel_prec.h"

#include "wx/wxAutoExcelWorkbook.h"

#include <wx/colour.h>

#include "wx/wxAutoExcel_object.h"
#include "wx/wxAutoExcelApplication.h"
#include "wx/wxAutoExcelChart.h"
#include "wx/wxAutoExcelCharts.h"
#include "wx/wxAutoExcelWindows.h"
#include "wx/wxAutoExcelWindows.h"
#include "wx/wxAutoExcelSheets.h"
#include "wx/wxAutoExcelSheet.h"
#include "wx/wxAutoExcelWorksheets.h"
#include "wx/wxAutoExcelWorksheet.h"
#include "wx/wxAutoExcelRange.h"
#include "wx/wxAutoExcelStyles.h"
#include "wx/wxAutoExcelIconSets.h"
#include "wx/wxAutoExcelDocumentProperties.h"

#include "wx/wxAutoExcel_private.h"

namespace wxAutoExcel {

// ***** class wxAutoExcelWorkbook METHODS *****

void wxExcelWorkbook::AcceptAllChanges(XlHighlightChangesTime* when, const wxString& who, const wxString& where)
{
    WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT_NAME(When, ((long*)when));
    WXAUTOEXCEL_OPTIONALCPPSTR_TO_OPTIONALVARIANT_NAME(Who, who);
    WXAUTOEXCEL_OPTIONALCPPSTR_TO_OPTIONALVARIANT_NAME(Where, where);

    WXAUTOEXCEL_CALL_METHOD3_RET("AcceptAllChanges", vWhen, vWho, vWhere, "null");
}

bool wxExcelWorkbook::Activate()
{
     WXAUTOEXCEL_CALL_METHOD0_BOOL("Activate");
}

void wxExcelWorkbook::AddToFavorites()
{
     WXAUTOEXCEL_CALL_METHOD0_RET("AddToFavorites", "null");
}

void wxExcelWorkbook::ApplyTheme(const wxString& fileName)
{
     WXAUTOEXCEL_CALL_METHOD1_RET("ApplyTheme", fileName, "null");
}

void wxExcelWorkbook::BreakLink(const wxString& name, XlLinkType type)
{
     WXAUTOEXCEL_CALL_METHOD2_RET("BreakLink", name, (long)type, "null");
}

bool wxExcelWorkbook::CanCheckIn()
{
    WXAUTOEXCEL_CALL_METHOD0_BOOL("CanCheckIn");
}

void wxExcelWorkbook::ChangeFileAccess(XlFileAccess mode, const wxString& writePassword, wxXlTribool notify)
{
    WXAUTOEXCEL_OPTIONALCPPSTR_TO_OPTIONALVARIANT_NAME(WritePassword, writePassword);
    WXAUTOEXCEL_OPTIONALCPPTBOOL_TO_OPTIONALVARIANT_NAME(Notify, notify);

    WXAUTOEXCEL_CALL_METHOD3_RET("ChangeFileAccess", wxVariant((long)mode), vWritePassword, vNotify, "null");
}

void wxExcelWorkbook::ChangeLink(const wxString& name, const wxString& newName, XlLinkType* type)
{
    WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT(Type, ((long*)type));

    WXAUTOEXCEL_CALL_METHOD3_RET("ChangeLink", name, newName, vType, "null");
}

wxExcelWorkbook wxExcelWorkbook::CheckIn(wxXlTribool saveChanges, wxXlTribool comments, wxXlTribool makePublic)
{
    wxExcelWorkbook workbook;

    WXAUTOEXCEL_OPTIONALCPPTBOOL_TO_OPTIONALVARIANT_NAME(SaveChanges, saveChanges);
    WXAUTOEXCEL_OPTIONALCPPTBOOL_TO_OPTIONALVARIANT_NAME(Comments, comments);
    WXAUTOEXCEL_OPTIONALCPPTBOOL_TO_OPTIONALVARIANT_NAME(MakePublic, makePublic);

    WXAUTOEXCEL_CALL_METHOD3("CheckIn", vSaveChanges, vComments, vMakePublic, "void*", workbook);
    VariantToObject(vResult, &workbook);
    return workbook;
}

void wxExcelWorkbook::CheckInWithVersion(wxXlTribool saveChanges, wxXlTribool comments, 
                                         wxXlTribool makePublic, XlCheckInVersionType* versionType)
{ 

    WXAUTOEXCEL_OPTIONALCPPTBOOL_TO_OPTIONALVARIANT_NAME(SaveChanges, saveChanges);
    WXAUTOEXCEL_OPTIONALCPPTBOOL_TO_OPTIONALVARIANT_NAME(Comments, comments);
    WXAUTOEXCEL_OPTIONALCPPTBOOL_TO_OPTIONALVARIANT_NAME(MakePublic, makePublic);
    WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT(VersionType, ((long*)versionType));

    WXAUTOEXCEL_CALL_METHOD4_RET("CheckInWithVersion", vSaveChanges, vComments, vMakePublic, vVersionType, "null");    
}

bool wxExcelWorkbook::Close(wxXlTribool saveChanges, const wxString& fileName, wxXlTribool routeWorkbook)
{
    WXAUTOEXCEL_OPTIONALCPPTBOOL_TO_OPTIONALVARIANT_NAME(SaveChanges, saveChanges);
    WXAUTOEXCEL_OPTIONALCPPSTR_TO_OPTIONALVARIANT_NAME(FileName, fileName);
    WXAUTOEXCEL_OPTIONALCPPTBOOL_TO_OPTIONALVARIANT_NAME(RouteWorkbook, routeWorkbook);

    WXAUTOEXCEL_CALL_METHOD3("Close", vSaveChanges, vFileName, vRouteWorkbook, "bool", false);
    return vResult.GetBool();
}

void wxExcelWorkbook::DeleteNumberFormat(const wxString& numberFormat)
{
     WXAUTOEXCEL_CALL_METHOD1_RET("DeleteNumberFormat", numberFormat, "null");
}

void wxExcelWorkbook::EnableConnections()
{
     WXAUTOEXCEL_CALL_METHOD0_RET("EnableConnections", "null");
}

void wxExcelWorkbook::EndReview()
{
     WXAUTOEXCEL_CALL_METHOD0_RET("EndReview", "null");
}

bool wxExcelWorkbook::ExclusiveAccess()
{
     WXAUTOEXCEL_CALL_METHOD0_BOOL("ExclusiveAccess");
}

void wxExcelWorkbook::ExportAsFixedFormat(XlFixedFormatType type, const wxString& fileName,
                                          XlFixedFormatQuality* quality, wxXlTribool includeDocProperties,
                                          wxXlTribool ignorePrintAreas, long* from, long* to, wxXlTribool openAfterPublish)
{

    wxVariantVector args;

    WXAUTOEXCEL_OPTIONALCPPSTR_TO_OPTIONALVARIANT_NAME_VECTOR(FileName, fileName, args);
    WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT_NAME_VECTOR(Quality, ((long*)quality), args);
    WXAUTOEXCEL_OPTIONALCPPTBOOL_TO_OPTIONALVARIANT_NAME_VECTOR(IncludeDocProperties, includeDocProperties, args);
    WXAUTOEXCEL_OPTIONALCPPTBOOL_TO_OPTIONALVARIANT_NAME_VECTOR(IgnorePrintAreas, ignorePrintAreas, args);
    WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT_NAME_VECTOR(From, from, args);
    WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT_NAME_VECTOR(To, to, args);
    WXAUTOEXCEL_OPTIONALCPPTBOOL_TO_OPTIONALVARIANT_NAME_VECTOR(OpenAfterPublish, openAfterPublish, args);

    ExportAsFixedFormat(type, args);
}

void wxExcelWorkbook::ExportAsFixedFormat(XlFixedFormatType type, const wxVariantVector& optionalArgs)
{
    wxVariantVector args(optionalArgs);

    args.push_back(wxVariant((long)type, wxS("Type")));

    WXAUTOEXCEL_CALL_METHODARR_RET("ExportAsFixedFormat", args, "null");
}

void wxExcelWorkbook::HighlightChangesOptions(XlHighlightChangesTime* when, const wxString& who, const wxString& where)
{
    WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT_NAME(When, ((long*)when));
    WXAUTOEXCEL_OPTIONALCPPSTR_TO_OPTIONALVARIANT_NAME(Who, who);
    WXAUTOEXCEL_OPTIONALCPPSTR_TO_OPTIONALVARIANT_NAME(Where, where);

    WXAUTOEXCEL_CALL_METHOD3_RET("HighlightChangesOptions", vWhen, vWho, vWhere, "null");
}


void wxExcelWorkbook::LockServerFile()
{
    WXAUTOEXCEL_CALL_METHOD0_RET("LockServerFile", "null");
}

void wxExcelWorkbook::MergeWorkbook(const wxString& fileName)
{
    WXAUTOEXCEL_CALL_METHOD1_RET("MergeWorkbook", fileName, "null");
}

wxExcelWindow wxExcelWorkbook::NewWindow()
{
    wxExcelWindow window;

    WXAUTOEXCEL_CALL_METHOD0_OBJECT("NewWindow", window);    
}


void wxExcelWorkbook::Post()
{
     WXAUTOEXCEL_CALL_METHOD0_RET("Post", "null");
}

bool wxExcelWorkbook::PrintOut(long* from, long* to, long* copies, wxXlTribool preview, const wxString& activePrinter,
                                 wxXlTribool printToFile, wxXlTribool collate, const wxString& prToFileName, wxXlTribool ignorePrintAreas)
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
    WXAUTOEXCEL_OPTIONALCPPTBOOL_TO_OPTIONALVARIANT_NAME_VECTOR(IgnorePrintAreas, ignorePrintAreas, args);

    return PrintOut(args);
}

bool wxExcelWorkbook::PrintOut(const wxVariantVector& args)
{
    WXAUTOEXCEL_CALL_METHODARR("PrintOut", args, "bool", false);
    return vResult.GetBool();
}

bool wxExcelWorkbook::PrintPreview(wxXlTribool enableChanges)
{
    WXAUTOEXCEL_OPTIONALCPPTBOOL_TO_OPTIONALVARIANT_NAME(EnableChanges, enableChanges);

    WXAUTOEXCEL_CALL_METHOD1_BOOL("PrintPreview", vEnableChanges);
}

void wxExcelWorkbook::Protect(const wxString& password, wxXlTribool structure, wxXlTribool windows)
{
    WXAUTOEXCEL_OPTIONALCPPSTR_TO_OPTIONALVARIANT_NAME(Password, password);
    WXAUTOEXCEL_OPTIONALCPPTBOOL_TO_OPTIONALVARIANT_NAME(Structure, structure);
    WXAUTOEXCEL_OPTIONALCPPTBOOL_TO_OPTIONALVARIANT_NAME(Windows, windows);

    WXAUTOEXCEL_CALL_METHOD3_RET("Protect", vPassword, vStructure, vWindows, "null");
}

void wxExcelWorkbook::ProtectSharing(const wxString& fileName, const wxString& password,
                                       const wxString& writeResPassword ,
                                       wxXlTribool readOnlyRecommended, wxXlTribool createBackup,
                                       const wxString& sharingPassword)
{
    wxVariantVector args;

    WXAUTOEXCEL_OPTIONALCPPSTR_TO_OPTIONALVARIANT_NAME_VECTOR(FileName, fileName, args);
    WXAUTOEXCEL_OPTIONALCPPSTR_TO_OPTIONALVARIANT_NAME_VECTOR(Password, password, args);
    WXAUTOEXCEL_OPTIONALCPPSTR_TO_OPTIONALVARIANT_NAME_VECTOR(WriteResPassword, writeResPassword, args);
    WXAUTOEXCEL_OPTIONALCPPTBOOL_TO_OPTIONALVARIANT_NAME_VECTOR(ReadOnlyRecommended, readOnlyRecommended, args);
    WXAUTOEXCEL_OPTIONALCPPTBOOL_TO_OPTIONALVARIANT_NAME_VECTOR(CreateBackup, createBackup, args);
    WXAUTOEXCEL_OPTIONALCPPSTR_TO_OPTIONALVARIANT_NAME_VECTOR(SharingPassword, sharingPassword, args);

    ProtectSharing(args);
}

void wxExcelWorkbook::ProtectSharing(const wxVariantVector& optionalArgs)
{
    WXAUTOEXCEL_CALL_METHODARR_RET("ProtectSharing", optionalArgs, "null");
}

void wxExcelWorkbook::PurgeChangeHistoryNow(long days, const wxString& sharingPassword)
{
    WXAUTOEXCEL_OPTIONALCPPSTR_TO_OPTIONALVARIANT_NAME(SharingPassword, sharingPassword);
    WXAUTOEXCEL_CALL_METHOD2_RET("PurgeChangeHistoryNow", wxVariant(days, wxS("Days")), vSharingPassword, "null");
}


void wxExcelWorkbook::RefreshAll()
{
     WXAUTOEXCEL_CALL_METHOD0_RET("RefreshAll", "null");
}

void wxExcelWorkbook::RejectAllChanges(XlHighlightChangesTime* when, const wxString& who, const wxString& where)
{
    WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT_NAME(When, ((long*)when));
    WXAUTOEXCEL_OPTIONALCPPSTR_TO_OPTIONALVARIANT_NAME(Who, who);
    WXAUTOEXCEL_OPTIONALCPPSTR_TO_OPTIONALVARIANT_NAME(Where, where);

    WXAUTOEXCEL_CALL_METHOD3_RET("RejectAllChanges", vWhen, vWho, vWhere, "null");
}

void wxExcelWorkbook::ReloadAs(MsoEncoding encoding)
{
    WXAUTOEXCEL_CALL_METHOD1_RET("ReloadAs", wxVariant((long)encoding), "null");
}

void wxExcelWorkbook::RemoveDocumentInformation(XlRemoveDocInfoType removeDocInfoType)
{
     WXAUTOEXCEL_CALL_METHOD1_RET("RemoveDocumentInformation", wxVariant((long)removeDocInfoType), "null");
}

void wxExcelWorkbook::RemoveUser(long index)
{
     WXAUTOEXCEL_CALL_METHOD1_RET("RemoveUser", index, "null");
}


void wxExcelWorkbook::ReplyWithChanges(wxXlTribool showMessage)
{
    WXAUTOEXCEL_OPTIONALCPPTBOOL_TO_OPTIONALVARIANT_NAME(ShowMessage, showMessage);
    WXAUTOEXCEL_CALL_METHOD1_RET("ReplyWithChanges", vShowMessage, "null");
}

void wxExcelWorkbook::ResetColors()
{
    WXAUTOEXCEL_CALL_METHOD0_RET("ResetColors", "null");
}

void wxExcelWorkbook::RunAutoMacros(XlRunAutoMacro which)
{
    WXAUTOEXCEL_CALL_METHOD1_RET("RunAutoMacros", wxVariant((long)which), "null");
}

bool wxExcelWorkbook::Save()
{
     WXAUTOEXCEL_CALL_METHOD0_BOOL("Save");
}

bool wxExcelWorkbook::SaveAs(const wxString& fileName, XlFileFormat* fileFormat,
                const wxString& password, const wxString& writeResPassword,
                wxXlTribool readOnlyRecommended, wxXlTribool createBackup,
                XlSaveAsAccessMode* accessMode, XlSaveConflictResolution* conflictResolution,
                wxXlTribool addToMru, wxXlTribool local)
{
    wxVariantVector args;

    WXAUTOEXCEL_OPTIONALCPPSTR_TO_OPTIONALVARIANT_NAME_VECTOR(FileName, fileName, args);
    WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT_NAME_VECTOR(FileFormat, ((long*)fileFormat), args);
    WXAUTOEXCEL_OPTIONALCPPSTR_TO_OPTIONALVARIANT_NAME_VECTOR(Password, password, args);
    WXAUTOEXCEL_OPTIONALCPPSTR_TO_OPTIONALVARIANT_NAME_VECTOR(WriteResPassword, writeResPassword, args);
    WXAUTOEXCEL_OPTIONALCPPTBOOL_TO_OPTIONALVARIANT_NAME_VECTOR(ReadOnlyRecommended, readOnlyRecommended, args);
    WXAUTOEXCEL_OPTIONALCPPTBOOL_TO_OPTIONALVARIANT_NAME_VECTOR(CreateBackup, createBackup, args);
    WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT_NAME_VECTOR(AccessMode, ((long*)accessMode), args);
    WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT_NAME_VECTOR(ConflictResolution, ((long*)conflictResolution), args);
    WXAUTOEXCEL_OPTIONALCPPTBOOL_TO_OPTIONALVARIANT_NAME_VECTOR(AddToMru, addToMru, args);
    WXAUTOEXCEL_OPTIONALCPPTBOOL_TO_OPTIONALVARIANT_NAME_VECTOR(Local, local, args);

    return SaveAs(args);
}

bool wxExcelWorkbook::SaveAs(const wxVariantVector& optionalArgs)
{
    WXAUTOEXCEL_CALL_METHODARR("SaveAs", optionalArgs, "bool", false);
    return vResult.GetBool();
}



bool wxExcelWorkbook::SaveCopyAs(const wxString& fileName)
{
     WXAUTOEXCEL_CALL_METHOD1_BOOL("SaveCopyAs", fileName);
}

void wxExcelWorkbook::SendFaxOverInternet(const wxString& recipients, const wxString& subject, wxXlTribool showMessage)
{
    WXAUTOEXCEL_OPTIONALCPPSTR_TO_OPTIONALVARIANT_NAME(Recipients, recipients);
    WXAUTOEXCEL_OPTIONALCPPSTR_TO_OPTIONALVARIANT_NAME(Subject, subject);
    WXAUTOEXCEL_OPTIONALCPPTBOOL_TO_OPTIONALVARIANT_NAME(ShowMessage, showMessage);

    WXAUTOEXCEL_CALL_METHOD3_RET("SendFaxOverInternet", vRecipients, vSubject, vShowMessage, "null");
}

void wxExcelWorkbook::SendForReview(const wxString& recipients, const wxString& subject,
                                      wxXlTribool showMessage, wxXlTribool includeAttachment)
{
    WXAUTOEXCEL_OPTIONALCPPSTR_TO_OPTIONALVARIANT_NAME(Recipients, recipients);
    WXAUTOEXCEL_OPTIONALCPPSTR_TO_OPTIONALVARIANT_NAME(Subject, subject);
    WXAUTOEXCEL_OPTIONALCPPTBOOL_TO_OPTIONALVARIANT_NAME(ShowMessage, showMessage);
    WXAUTOEXCEL_OPTIONALCPPTBOOL_TO_OPTIONALVARIANT_NAME(IncludeAttachment, includeAttachment);

    WXAUTOEXCEL_CALL_METHOD4_RET("SendForReview", vRecipients, vSubject, vShowMessage, vIncludeAttachment, "null");
}

void wxExcelWorkbook::SendMail(const wxString& recipients, const wxString& subject, wxXlTribool returnReceipt)
{
    WXAUTOEXCEL_OPTIONALCPPSTR_TO_OPTIONALVARIANT_NAME(Recipients, recipients);
    WXAUTOEXCEL_OPTIONALCPPSTR_TO_OPTIONALVARIANT_NAME(Subject, subject);
    WXAUTOEXCEL_OPTIONALCPPTBOOL_TO_OPTIONALVARIANT_NAME(ReturnReceipt, returnReceipt);

    WXAUTOEXCEL_CALL_METHOD3_RET("SendMail", vRecipients, vSubject, vReturnReceipt, "null");
}


void wxExcelWorkbook::SetPasswordEncryptionOptions(const wxString& passwordEncryptionProvider,
                                                     const wxString& passwordEncryptionAlgorithm,
                                                     long* passwordEncryptionKeyLength,
                                                     wxXlTribool passwordEncryptionFileProperties)
{
    WXAUTOEXCEL_OPTIONALCPPSTR_TO_OPTIONALVARIANT_NAME(PasswordEncryptionProvider, passwordEncryptionProvider);
    WXAUTOEXCEL_OPTIONALCPPSTR_TO_OPTIONALVARIANT_NAME(PasswordEncryptionAlgorithm, passwordEncryptionAlgorithm);
    WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT_NAME(PasswordEncryptionKeyLength, passwordEncryptionKeyLength);
    WXAUTOEXCEL_OPTIONALCPPTBOOL_TO_OPTIONALVARIANT_NAME(PasswordEncryptionFileProperties, passwordEncryptionFileProperties);

    WXAUTOEXCEL_CALL_METHOD4_RET("SetPasswordEncryptionOptions", vPasswordEncryptionProvider, vPasswordEncryptionAlgorithm, vPasswordEncryptionKeyLength, vPasswordEncryptionFileProperties, "null");
}

bool wxExcelWorkbook::Unprotect(const wxString& password)
{
    WXAUTOEXCEL_OPTIONALCPPSTR_TO_OPTIONALVARIANT_NAME(Password, password);
    WXAUTOEXCEL_CALL_METHOD1_BOOL("Unprotect", vPassword);
}

bool wxExcelWorkbook::UnprotectSharing(const wxString& sharingPassword)
{
     WXAUTOEXCEL_OPTIONALCPPSTR_TO_OPTIONALVARIANT_NAME(SharingPassword, sharingPassword);
     WXAUTOEXCEL_CALL_METHOD1_BOOL("UnprotectSharing", vSharingPassword);
}

void wxExcelWorkbook::UpdateFromFile()
{
     WXAUTOEXCEL_CALL_METHOD0_RET("UpdateFromFile", "null");
}

void wxExcelWorkbook::WebPagePreview()
{
    WXAUTOEXCEL_CALL_METHOD0_RET("WebPagePreview", "null");
}


// ***** class wxAutoExcelWorkbook PROPERTIES *****

#if WXAUTOEXCEL_USE_CHARTS

wxExcelChart wxExcelWorkbook::GetActiveChart()
{
    wxExcelChart chart;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("ActiveChart", chart);
}

#endif // #if WXAUTOEXCEL_USE_CHARTS

wxExcelSheet wxExcelWorkbook::GetActiveSheet()
{
    wxExcelSheet sheet;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("ActiveSheet", sheet);
}

long wxExcelWorkbook::GetAccuracyVersion()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("AccuracyVersion");
}

void wxExcelWorkbook::SetAccuracyVersion(long accuracyVersion)
{
    InvokePutProperty("AccuracyVersion", accuracyVersion);
}


long wxExcelWorkbook::GetAutoUpdateFrequency()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("AutoUpdateFrequency");
}

void wxExcelWorkbook::SetAutoUpdateFrequency(long autoUpdateFrequency)
{
    InvokePutProperty("AutoUpdateFrequency", autoUpdateFrequency);
}

bool wxExcelWorkbook::GetAutoUpdateSaveChanges()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("AutoUpdateSaveChanges");
}

void wxExcelWorkbook::SetAutoUpdateSaveChanges(bool autoUpdateSaveChanges)
{
    InvokePutProperty("AutoUpdateSaveChanges", autoUpdateSaveChanges);
}

wxExcelDocumentProperties wxExcelWorkbook::GetBuiltinDocumentProperties()
{
    wxExcelDocumentProperties props;
    
    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("BuiltinDocumentProperties", props);
}


long wxExcelWorkbook::GetCalculationVersion()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("CalculationVersion");
}

long wxExcelWorkbook::GetChangeHistoryDuration()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("ChangeHistoryDuration");
}

void wxExcelWorkbook::SetChangeHistoryDuration(long changeHistoryDuration)
{
    InvokePutProperty("ChangeHistoryDuration", changeHistoryDuration);
}

#if WXAUTOEXCEL_USE_CHARTS

wxExcelCharts wxExcelWorkbook::GetCharts()
{
    wxExcelCharts charts;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("Charts", charts);
}

#endif // #if WXAUTOEXCEL_USE_CHARTS

bool wxExcelWorkbook::GetCheckCompatibility()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("CheckCompatibility");
}

void wxExcelWorkbook::SetCheckCompatibility(bool checkCompatibility)
{
    InvokePutProperty("CheckCompatibility", checkCompatibility);
}


wxColour wxExcelWorkbook::GetColors(long index)
{
    wxColour result;
    wxVariant vResult;

    if ( InvokeGetProperty(wxS("Colors"), vResult, index) )
    {
        WXAUTOEXCEL_CHECK_VARIANT_TYPE(vResult, "double", "Colors", result);
        result.SetRGB(vResult.GetDouble());
    }
    return result;
}

wxVector<wxColour> wxExcelWorkbook::GetColors()
{
    wxVector<wxColour> result;
    wxVariant vResult;

    if ( InvokeGetProperty(wxS("Colors"), vResult) )
    {
        WXAUTOEXCEL_CHECK_VARIANT_TYPE(vResult, "list", "Colors", result);
        result.reserve(vResult.GetCount());
        for ( size_t i = 0; i < vResult.GetCount(); i++ )
            result.push_back(wxColour(vResult[i].GetDouble()));
    }
    return result;
}

void wxExcelWorkbook::SetColors(long index, const wxColour& color)
{
    InvokePutProperty("Colors", index, (long)color.GetRGB());
}

void wxExcelWorkbook::SetColors(const wxVector<wxColour>& colors)
{
    wxASSERT( colors.size() == 56 );

    wxVariant vColors;

    vColors.NullList();
    for (size_t i = 0; i < colors.size(); i++)
        vColors.Append((long)colors[i].GetRGB());
    InvokePutProperty("Colors", vColors);
}

XlSaveConflictResolution wxExcelWorkbook::GetConflictResolution()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("ConflictResolution", XlSaveConflictResolution, (XlSaveConflictResolution)0);
}

void wxExcelWorkbook::SetConflictResolution(XlSaveConflictResolution conflictResolution)
{
    InvokePutProperty("ConflictResolution", (long)conflictResolution);
}

bool wxExcelWorkbook::GetConnectionsDisabled()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("ConnectionsDisabled");
}

bool wxExcelWorkbook::GetCreateBackup()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("CreateBackup");
}

wxExcelDocumentProperties wxExcelWorkbook::GetCustomDocumentProperties()
{
    wxExcelDocumentProperties props;
    
    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("CustomDocumentProperties", props);
}


bool wxExcelWorkbook::GetDate1904()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("Date1904");
}

void wxExcelWorkbook::SetDate1904(bool date1904)
{
    InvokePutProperty("Date1904", date1904);
}


bool wxExcelWorkbook::GetDisplayInkComments()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("DisplayInkComments");
}

void wxExcelWorkbook::SetDisplayInkComments(bool displayInkComments)
{
    InvokePutProperty("DisplayInkComments", displayInkComments);
}

bool wxExcelWorkbook::GetDoNotPromptForConvert()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("DoNotPromptForConvert");
}

void wxExcelWorkbook::SetDoNotPromptForConvert(bool doNotPromptForConvert)
{
    InvokePutProperty("DoNotPromptForConvert", doNotPromptForConvert);
}

bool wxExcelWorkbook::GetEnableAutoRecover()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("EnableAutoRecover");
}

void wxExcelWorkbook::SetEnableAutoRecover(bool enableAutoRecover)
{
    InvokePutProperty("EnableAutoRecover", enableAutoRecover);
}

wxString wxExcelWorkbook::GetEncryptionProvider()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("EncryptionProvider");
}

void wxExcelWorkbook::SetEncryptionProvider(const wxString& encryptionProvider)
{
    InvokePutProperty("EncryptionProvider", encryptionProvider);
}

bool wxExcelWorkbook::GetEnvelopeVisible()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("EnvelopeVisible");
}

void wxExcelWorkbook::SetEnvelopeVisible(bool envelopeVisible)
{
    InvokePutProperty("EnvelopeVisible", envelopeVisible);
}

bool wxExcelWorkbook::GetExcel8CompatibilityMode()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("Excel8CompatibilityMode");
}

XlFileFormat wxExcelWorkbook::GetFileFormat()
{
    XlFileFormat result = xlWorkbookDefault;
    wxVariant vResult;
    
    // MS Excel for some reason returns the value as a double
    if ( InvokeGetProperty("FileFormat", vResult) )
    {
        long l;
        if ( vResult.MakeString().ToCLong(&l) )
            result = XlFileFormat(l);
    }
    return result;
}

bool wxExcelWorkbook::GetFinal()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("Final");
}

void wxExcelWorkbook::SetFinal(bool final)
{
    InvokePutProperty("Final", final);
}

bool wxExcelWorkbook::GetForceFullCalculation()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("ForceFullCalculation");
}

void wxExcelWorkbook::SetForceFullCalculation(bool forceFullCalculation)
{
    InvokePutProperty("ForceFullCalculation", forceFullCalculation);
}

wxString wxExcelWorkbook::GetFullName()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("FullName");
}

wxString wxExcelWorkbook::GetFullNameURLEncoded()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("FullNameURLEncoded");
}

bool wxExcelWorkbook::GetHasPassword()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("HasPassword");
}

bool wxExcelWorkbook::GetHasVBProject()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("HasVBProject");
}

bool wxExcelWorkbook::GetHighlightChangesOnScreen()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("HighlightChangesOnScreen");
}

void wxExcelWorkbook::SetHighlightChangesOnScreen(bool highlightChangesOnScreen)
{
    InvokePutProperty("HighlightChangesOnScreen", highlightChangesOnScreen);
}

#if WXAUTOEXCEL_USE_CONDFORMAT

wxExcelIconSets wxExcelWorkbook::GetIconSets(XlIconSetE index)
{
    wxExcelIconSets sets;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET1("IconSets", index, sets);
}

#endif // #if WXAUTOEXCEL_USE_CONDFORMAT

bool wxExcelWorkbook::GetInactiveListBorderVisible()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("InactiveListBorderVisible");
}

void wxExcelWorkbook::SetInactiveListBorderVisible(bool inactiveListBorderVisible)
{
    InvokePutProperty("InactiveListBorderVisible", inactiveListBorderVisible);
}

bool wxExcelWorkbook::GetIsAddin()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("IsAddin");
}

void wxExcelWorkbook::SetIsAddin(bool isAddin)
{
    InvokePutProperty("IsAddin", isAddin);
}

bool wxExcelWorkbook::GetIsInplace()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("IsInplace");
}

bool wxExcelWorkbook::GetKeepChangeHistory()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("KeepChangeHistory");
}

void wxExcelWorkbook::SetKeepChangeHistory(bool keepChangeHistory)
{
    InvokePutProperty("KeepChangeHistory", keepChangeHistory);
}

bool wxExcelWorkbook::GetListChangesOnNewSheet()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("ListChangesOnNewSheet");
}

void wxExcelWorkbook::SetListChangesOnNewSheet(bool listChangesOnNewSheet)
{
    InvokePutProperty("ListChangesOnNewSheet", listChangesOnNewSheet);
}

bool wxExcelWorkbook::GetMultiUserEditing()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("MultiUserEditing");
}

wxString wxExcelWorkbook::GetName()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("Name");
}

wxString wxExcelWorkbook::GetPassword()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("Password");
}

void wxExcelWorkbook::SetPassword(const wxString& password)
{
    InvokePutProperty("Password", password);
}

wxString wxExcelWorkbook::GetPasswordEncryptionAlgorithm()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("PasswordEncryptionAlgorithm");
}

bool wxExcelWorkbook::GetPasswordEncryptionFileProperties()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("PasswordEncryptionFileProperties");
}

long wxExcelWorkbook::GetPasswordEncryptionKeyLength()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("PasswordEncryptionKeyLength");
}

wxString wxExcelWorkbook::GetPasswordEncryptionProvider()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("PasswordEncryptionProvider");
}

wxString wxExcelWorkbook::GetPath()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("Path");
}


bool wxExcelWorkbook::GetPersonalViewListSettings()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("PersonalViewListSettings");
}

void wxExcelWorkbook::SetPersonalViewListSettings(bool personalViewListSettings)
{
    InvokePutProperty("PersonalViewListSettings", personalViewListSettings);
}

bool wxExcelWorkbook::GetPersonalViewPrintSettings()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("PersonalViewPrintSettings");
}

void wxExcelWorkbook::SetPersonalViewPrintSettings(bool personalViewPrintSettings)
{
    InvokePutProperty("PersonalViewPrintSettings", personalViewPrintSettings);
}

bool wxExcelWorkbook::GetPrecisionAsDisplayed()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("PrecisionAsDisplayed");
}

void wxExcelWorkbook::SetPrecisionAsDisplayed(bool precisionAsDisplayed)
{
    InvokePutProperty("PrecisionAsDisplayed", precisionAsDisplayed);
}

bool wxExcelWorkbook::GetProtectStructure()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("ProtectStructure");
}

bool wxExcelWorkbook::GetProtectWindows()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("ProtectWindows");
}

bool wxExcelWorkbook::GetReadOnly()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("ReadOnly");
}

bool wxExcelWorkbook::GetReadOnlyRecommended()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("ReadOnlyRecommended");
}

bool wxExcelWorkbook::GetRemovePersonalInformation()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("RemovePersonalInformation");
}

void wxExcelWorkbook::SetRemovePersonalInformation(bool removePersonalInformation)
{
    InvokePutProperty("RemovePersonalInformation", removePersonalInformation);
}

long wxExcelWorkbook::GetRevisionNumber()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("RevisionNumber");
}

bool wxExcelWorkbook::GetSaved()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("Saved");
}

void wxExcelWorkbook::SetSaved(bool saved)
{
    InvokePutProperty("Saved", saved);
}

bool wxExcelWorkbook::GetSaveLinkValues()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("SaveLinkValues");
}

void wxExcelWorkbook::SetSaveLinkValues(bool saveLinkValues)
{
    InvokePutProperty("SaveLinkValues", saveLinkValues);
}

wxExcelSheets wxExcelWorkbook::GetSheets()
{
    wxExcelSheets sheets;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("Sheets", sheets);
}

bool wxExcelWorkbook::GetShowConflictHistory()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("ShowConflictHistory");
}

void wxExcelWorkbook::SetShowConflictHistory(bool showConflictHistory)
{
    InvokePutProperty("ShowConflictHistory", showConflictHistory);
}

wxExcelStyles wxExcelWorkbook::GetStyles()
{
    wxExcelStyles styles;
    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("Styles", styles);
}

bool wxExcelWorkbook::GetTemplateRemoveExtData()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("TemplateRemoveExtData");
}

void wxExcelWorkbook::SetTemplateRemoveExtData(bool templateRemoveExtData)
{
    InvokePutProperty("TemplateRemoveExtData", templateRemoveExtData);
}


XlUpdateLinks wxExcelWorkbook::GetUpdateLinks()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("UpdateLinks", XlUpdateLinks, (XlUpdateLinks)0);
}

void wxExcelWorkbook::SetUpdateLinks(XlUpdateLinks updateLinks)
{
    InvokePutProperty("UpdateLinks", (long)updateLinks);;
}

bool wxExcelWorkbook::GetUpdateRemoteReferences()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("UpdateRemoteReferences");
}

void wxExcelWorkbook::SetUpdateRemoteReferences(bool updateRemoteReferences)
{
    InvokePutProperty("UpdateRemoteReferences", updateRemoteReferences);
}

bool wxExcelWorkbook::GetVBASigned()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("VBASigned");
}


wxExcelWindows wxExcelWorkbook::GetWindows()
{
    wxExcelWindows windows;
    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("Windows", windows);         
}

wxExcelWorksheets wxExcelWorkbook::GetWorksheets()
{
    wxExcelWorksheets worksheets;
    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("Worksheets", worksheets);
}

wxString wxExcelWorkbook::GetWritePassword()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("WritePassword");
}

void wxExcelWorkbook::SetWritePassword(const wxString& writePassword)
{
    InvokePutProperty("WritePassword", writePassword);
}

bool wxExcelWorkbook::GetWriteReserved()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("WriteReserved");
}

wxString wxExcelWorkbook::GetWriteReservedBy()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("WriteReservedBy");
}

} // namespace wxAutoExcel
