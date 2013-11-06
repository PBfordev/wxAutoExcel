/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Modified by:
// RCS-ID:      $Id: $
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////


#include "wx/wxAutoExcel_prec.h"

#include "wx/wxAutoExcelObject.h"
#include "wx/wxAutoExcelApplication.h"
#include "wx/wxAutoExcelWindows.h"
#include "wx/wxAutoExcelWorkbooks.h"
#include "wx/wxAutoExcelWorkbook.h"
#include "wx/wxAutoExcelSheets.h"
#include "wx/wxAutoExcelSheet.h"
#include "wx/wxAutoExcelWorksheets.h"
#include "wx/wxAutoExcelRange.h"

#include "wx/wxAutoExcelPrivate.h"

#include <wx/msw/winundef.h>

namespace wxAutoExcel {

wxExcelApplication wxExcelApplication::CreateInstance()
{
    wxExcelApplication instance;

    instance.m_xlObject->CreateInstance(wxS("Excel.Application"));
    return instance;
}


// ***** class wxAutoExcelApplication METHODS *****

void wxExcelApplication::ActivateMicrosoftApp(XlMSApplication index)
{
    WXAUTOEXCEL_CALL_METHOD1_RET("ActivateMicrosoftApp", wxVariant(index), "null");
}

bool wxExcelApplication::AddCustomList(const wxArrayString& listArray)
{
    wxASSERT(!listArray.empty());

    wxVariant vList(listArray, wxS("ListArray"));

    WXAUTOEXCEL_CALL_METHOD1_BOOL("AddCustomList", vList);
}

bool wxExcelApplication::AddCustomList(wxExcelRange listArray, bool byRow)
{
    wxVariant vList;
    if ( ObjectToVariant(&listArray, vList, wxS("ListArray")) )
    {

        WXAUTOEXCEL_CALL_METHOD2("AddCustomList", vList, wxVariant(byRow, wxS("ByRow")), "bool", false);
        return vResult.GetBool();
    }
    return false;
}

void wxExcelApplication::Calculate()
{
    WXAUTOEXCEL_CALL_METHOD0_RET("Calculate", "null");
}

void wxExcelApplication::CalculateFull()
{
    WXAUTOEXCEL_CALL_METHOD0_RET("CalculateFull", "null");
}

void wxExcelApplication::CalculateFullRebuild()
{
    WXAUTOEXCEL_CALL_METHOD0_RET("CalculateFullRebuild", "null");
}

void wxExcelApplication::CalculateUntilAsyncQueriesDone()
{
    WXAUTOEXCEL_CALL_METHOD0_RET("CalculateUntilAsyncQueriesDone", "null");
}

double wxExcelApplication::CentimetersToPoints(double centimeters)
{
    WXAUTOEXCEL_CALL_METHOD1_DOUBLE("CentimetersToPoints", centimeters, 0.);
}

void wxExcelApplication::CheckAbort(wxExcelRange* keepAbort)
{
    wxVariant vKeepAbort;

    if (keepAbort != NULL)
        ObjectToVariant(keepAbort, vKeepAbort);
    WXAUTOEXCEL_CALL_METHOD1_RET("CheckAbort", vKeepAbort, "null");
}

bool wxExcelApplication::CheckSpelling(const wxString& word, const wxString& customDictionary, wxXlTribool ignoreUpperCase)
{
    wxASSERT(!word.empty());

    wxVariant vWord(word, wxS("Word"));

    WXAUTOEXCEL_OPTIONALCPPSTR_TO_OPTIONALVARIANT_NAME(CustomDictionary, customDictionary);
    WXAUTOEXCEL_OPTIONALCPPTBOOL_TO_OPTIONALVARIANT_NAME(IgnoreUpperCase, ignoreUpperCase);

    WXAUTOEXCEL_CALL_METHOD3("CheckSpelling", vWord, vCustomDictionary, vIgnoreUpperCase, "bool", false);
    return vResult.GetBool();
}

wxString wxExcelApplication::ConvertFormula(const wxString& formula, XlReferenceStyle fromReferenceStyle,
    XlReferenceStyle* toReferenceStyle, XlReferenceStyle* toAbsolute,
    wxExcelRange* relativeTo)
{
    wxVariant vFormula(formula, wxS("Formula"));
    wxVariant vFromReferenceStyle((long)fromReferenceStyle, wxS("FromReferenceStyle"));
    WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT_NAME(ToReferenceStyle, (long*)toReferenceStyle);
    WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT_NAME(ToAbsolute, (long*)toAbsolute);
    wxVariant vRelativeTo;

    if (relativeTo != NULL)
        ObjectToVariant(relativeTo, vRelativeTo, wxS("RelativeTo"));

    WXAUTOEXCEL_CALL_METHOD5("ConvertFormula", vFormula, vFromReferenceStyle, vToReferenceStyle, vToAbsolute, vRelativeTo, "string", wxS(""));
    return vResult.GetString();
}

void wxExcelApplication::DDEExecute(long channel, const wxString& str)
{
    WXAUTOEXCEL_CALL_METHOD2_RET("DDEExecute", wxVariant(channel), str, "null");
}

long wxExcelApplication::DDEInitiate(const wxString& app, const wxString& topic)
{
    WXAUTOEXCEL_CALL_METHOD2("DDEInitiate", app, topic, "null", -1);
    return vResult.GetLong();
}

void wxExcelApplication::DDEPoke(long channel, const wxVariant& item, const wxVariant& data)
{
    WXAUTOEXCEL_CALL_METHOD3_RET("DDEPoke", wxVariant(channel), item, data, "null");
}

wxVariant wxExcelApplication::DDERequest(long channel, const wxString& item)
{
    WXAUTOEXCEL_CALL_METHOD2("DDERequest", wxVariant(channel), item, "list", wxVariant());
    return vResult;
}

void wxExcelApplication::DDETerminate(long channel)
{
    WXAUTOEXCEL_CALL_METHOD1_RET("DDETerminate", wxVariant(channel), "null");
}

bool wxExcelApplication::DeleteCustomList(long listNum)
{
    wxASSERT(listNum >= 5);

    WXAUTOEXCEL_CALL_METHOD1_BOOL("DeleteCustomList", wxVariant(listNum));
}


void wxExcelApplication::DoubleClick()
{
    WXAUTOEXCEL_CALL_METHOD0_RET("DoubleClick", "null");
}


bool wxExcelApplication::FindFile()
{
    WXAUTOEXCEL_CALL_METHOD0_BOOL("FindFile");
}

wxArrayString wxExcelApplication::GetCustomListContents(long listNum)
{
    WXAUTOEXCEL_CALL_METHOD1("GetCustomListContents", wxVariant(listNum), "list", wxArrayString());
    wxArrayString list;
    for (size_t i = 0; i < vResult.GetCount(); i++)
        list.push_back(vResult[i].GetString());
    return list;
}

long wxExcelApplication::GetCustomListNum(const wxArrayString& listArray)
{
    WXAUTOEXCEL_CALL_METHOD1_LONG("GetCustomListNum", wxVariant(listArray), -1);

}

wxArrayString wxExcelApplication::GetOpenFilename(const wxString& fileFilter, long* filterIndex,
    const wxString& title, wxXlTribool multiSelect)
{
    wxArrayString result;
    wxVariant vResult;

    WXAUTOEXCEL_OPTIONALCPPTBOOL_TO_OPTIONALVARIANT_NAME(MultiSelect, multiSelect);
    WXAUTOEXCEL_OPTIONALCPPSTR_TO_OPTIONALVARIANT_NAME(FileFilter, fileFilter);
    WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT_NAME(FilterIndex, filterIndex);
    WXAUTOEXCEL_OPTIONALCPPSTR_TO_OPTIONALVARIANT_NAME(Title, title);

    if ( InvokeMethod(wxS("GetOpenFilename"), vResult, vFileFilter, vFilterIndex, vTitle, vMultiSelect) )
    {
        wxString type = vResult.GetType();
        if (type == "string") {
            result.push_back(vResult.GetString());
        } else if (type == "arrstring") {
            result = vResult.GetArrayString();
        } else if (vResult.GetType() == "list") {
            for (size_t i = 0; i < vResult.GetCount(); i++)
                result.push_back(vResult[i].GetString());
        } else {
            // user cancelled the dialog so the method returns false            
        }
    }
    return result;
}

wxString wxExcelApplication::GetPhonetic(const wxString& text)
{
    WXAUTOEXCEL_OPTIONALCPPSTR_TO_OPTIONALVARIANT(Text, text);

    WXAUTOEXCEL_CALL_METHOD1_STRING("GetPhonetic", vText);
}

wxString wxExcelApplication::GetSaveAsFilename(const wxString& initialFilename,
                                               const wxString& fileFilter, long* filterIndex,
                                               const wxString& title)
{

    WXAUTOEXCEL_OPTIONALCPPSTR_TO_OPTIONALVARIANT_NAME(InitialFilename, initialFilename);
    WXAUTOEXCEL_OPTIONALCPPSTR_TO_OPTIONALVARIANT_NAME(FileFilter, fileFilter);
    WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT_NAME(FilterIndex, filterIndex);
    WXAUTOEXCEL_OPTIONALCPPSTR_TO_OPTIONALVARIANT_NAME(Title, title);

    wxVariant vResult;
    wxString result;
    if ( InvokeMethod(wxS("GetSaveAsFilename"), vResult, vInitialFilename, vFileFilter, vFilterIndex, vTitle) )
    {
        // if not than the user cancelled the dialog and we return empty string
        if  ( vResult.GetType() == wxS("string") ) 
            result = vResult.GetString();        
    }
    return result;
}


void wxExcelApplication::Goto(const wxString& reference, bool scroll)
{
    wxVariant vScroll(scroll, wxS("Scroll"));

    WXAUTOEXCEL_OPTIONALCPPSTR_TO_OPTIONALVARIANT_NAME(Reference, reference);

    WXAUTOEXCEL_CALL_METHOD2_RET("Goto", vReference, vScroll, "null");
}

void wxExcelApplication::Goto(wxExcelRange range, bool scroll)
{
    wxVariant vReference;

    if ( ObjectToVariant(&range, vReference, wxS("Reference")) )
    {
        WXAUTOEXCEL_CALL_METHOD2_RET("Goto", vReference, wxVariant(scroll, wxS("Scroll")), "null");
    }
}

void wxExcelApplication::Help(const wxString& helpFile, long* helpContextID)
{
    WXAUTOEXCEL_OPTIONALCPPSTR_TO_OPTIONALVARIANT_NAME(HelpFile, helpFile);
    WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT_NAME(HelpContextID, helpContextID);

    WXAUTOEXCEL_CALL_METHOD2_RET("Help", vHelpFile, vHelpContextID, "null");
}

double wxExcelApplication::InchesToPoints(double inches)
{
    WXAUTOEXCEL_CALL_METHOD1_DOUBLE("InchesToPoints", inches, 0.);
}

wxExcelRange wxExcelApplication::Intersect(const wxExcelRangeVector& ranges)
{
    wxASSERT( ranges.size() > 1 );

    wxVariantVector variants;    
    wxExcelRange range;

    if ( RangesToVariants(ranges, variants) )
    {        
        WXAUTOEXCEL_CALL_METHODARR("Intersect", variants, "void*", range);        
            VariantToObject(vResult, &range);        
    }
    return range;
}


void wxExcelApplication::MailLogoff()
{
    WXAUTOEXCEL_CALL_METHOD0_RET("MailLogoff", "null");
}

void wxExcelApplication::MailLogon(const wxString& name, const wxString& password, wxXlTribool downloadNewMail)
{
    WXAUTOEXCEL_OPTIONALCPPSTR_TO_OPTIONALVARIANT_NAME(Name, name);
    WXAUTOEXCEL_OPTIONALCPPSTR_TO_OPTIONALVARIANT_NAME(Password, password);
    WXAUTOEXCEL_OPTIONALCPPTBOOL_TO_OPTIONALVARIANT_NAME(DownloadNewMail, downloadNewMail);
    WXAUTOEXCEL_CALL_METHOD3_RET("MailLogon", vName, vPassword, vDownloadNewMail, "null");
}


void wxExcelApplication::Quit()
{
    WXAUTOEXCEL_CALL_METHOD0_RET("Quit", "null");
}

void wxExcelApplication::RecordMacro(const wxString& basicCode, const wxString& XlmCode)
{
    if (basicCode.empty() && XlmCode.empty()) {
        WXAUTOEXCEL_CALL_METHOD2_RET("RecordMacro", wxS(""), wxS(""), "null");
    } else {
        WXAUTOEXCEL_CALL_METHOD1_RET("RecordMacro", basicCode, "null");
    }
}

bool wxExcelApplication::RegisterXLL(const wxString& fileName)
{
    WXAUTOEXCEL_CALL_METHOD1_BOOL("RegisterXLL", fileName);
}

        
wxVariant wxExcelApplication::Run(const wxString& macro, const wxVariantVector& arguments)
{
    wxVariant vResult;
    wxVariantVector args;
    
    args.reserve(arguments.size() + 1);
    args.push_back(wxVariant(macro));
    for (size_t i = 0; i < arguments.size(); i++)
        args.push_back(arguments[i]);
    InvokeMethodArray(wxS("Run"), vResult, args);
    return vResult;    
}

void wxExcelApplication::SaveWorkspace(const wxString& fileName)
{
    WXAUTOEXCEL_OPTIONALCPPSTR_TO_OPTIONALVARIANT_NAME(FileName, fileName);

    WXAUTOEXCEL_CALL_METHOD1_RET("SaveWorkspace", vFileName, "null");
}

void wxExcelApplication::SendKeys(wxString& keys, bool wait)
{
    WXAUTOEXCEL_CALL_METHOD2_RET("SendKeys", keys, wait, "null");
}

long wxExcelApplication::SharePointVersion(const wxString& url)
{
    WXAUTOEXCEL_CALL_METHOD1("SharePointVersion", url, "long", 0l);
    return vResult.GetLong();
}


wxExcelRange wxExcelApplication::Union(const wxExcelRangeVector& ranges)
{
    wxASSERT( ranges.size() > 1 );

    wxVariantVector variants;    
    wxExcelRange range;

    if ( RangesToVariants(ranges, variants) )
    {        
        WXAUTOEXCEL_CALL_METHODARR("Union", variants, "void*", range);        
        VariantToObject(vResult, &range);
    }
    return range;
}

// ***** class wxAutoExcelApplication PROPERTIES *****

wxExcelRange wxExcelApplication::GetActiveCell()
{
    wxExcelRange range;
    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("ActiveCell", range);
}

wxString wxExcelApplication::GetActivePrinter()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("ActivePrinter");
}

void wxExcelApplication::SetActivePrinter(const wxString& activePrinter)
{
    InvokePutProperty("ActivePrinter", activePrinter);
}

wxExcelSheet wxExcelApplication::GetActiveSheet()
{
    wxExcelSheet sheet;
    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("ActiveSheet", sheet);
}

wxExcelWindow wxExcelApplication::GetActiveWindow()
{
    wxExcelWindow window;
    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("ActiveWindow", window);
}

wxExcelWorkbook wxExcelApplication::GetActiveWorkbook()
{
    wxExcelWorkbook workbook;
    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("ActiveWorkbook", workbook);
}


bool wxExcelApplication::GetAlertBeforeOverwriting()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("AlertBeforeOverwriting");
}

void wxExcelApplication::SetAlertBeforeOverwriting(bool alertBeforeOverwriting)
{
    InvokePutProperty("AlertBeforeOverwriting", alertBeforeOverwriting);
}

wxString wxExcelApplication::GetAltStartupPath()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("AltStartupPath");
}

void wxExcelApplication::SetAltStartupPath(const wxString& altStartupPath)
{
    InvokePutProperty("AltStartupPath", altStartupPath);
}

bool wxExcelApplication::GetAlwaysUseClearType()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("AlwaysUseClearType");
}

void wxExcelApplication::SetAlwaysUseClearType(bool alwaysUseClearType)
{
    InvokePutProperty("AlwaysUseClearType", alwaysUseClearType);
}


bool wxExcelApplication::GetArbitraryXMLSupportAvailable()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("ArbitraryXMLSupportAvailable");
}

bool wxExcelApplication::GetAskToUpdateLinks()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("AskToUpdateLinks");
}

void wxExcelApplication::SetAskToUpdateLinks(bool askToUpdateLinks)
{
    InvokePutProperty("AskToUpdateLinks", askToUpdateLinks);
}

bool wxExcelApplication::GetAutoFormatAsYouTypeReplaceHyperlinks()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("AutoFormatAsYouTypeReplaceHyperlinks");
}

void wxExcelApplication::SetAutoFormatAsYouTypeReplaceHyperlinks(bool autoFormatAsYouTypeReplaceHyperlinks)
{
    InvokePutProperty("AutoFormatAsYouTypeReplaceHyperlinks", autoFormatAsYouTypeReplaceHyperlinks);
}

MsoAutomationSecurity wxExcelApplication::GetAutomationSecurity()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("AutomationSecurity", MsoAutomationSecurity, msoAutomationSecurityByUI);
}

void wxExcelApplication::SetAutomationSecurity(MsoAutomationSecurity security)
{
    InvokePutProperty("AutomationSecurity", security);
}

bool wxExcelApplication::GetAutoPercentEntry()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("AutoPercentEntry");
}

void wxExcelApplication::SetAutoPercentEntry(bool autoPercentEntry)
{
    InvokePutProperty("AutoPercentEntry", autoPercentEntry);
}


double wxExcelApplication::GetBuild()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("Build");
}

bool wxExcelApplication::GetCalculateBeforeSave()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("CalculateBeforeSave");
}

void wxExcelApplication::SetCalculateBeforeSave(bool calculateBeforeSave)
{
    InvokePutProperty("CalculateBeforeSave", calculateBeforeSave);
}

XlCalculation wxExcelApplication::GetCalculation()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("Calculation", XlCalculation, xlCalculationAutomatic);
}

void wxExcelApplication::SetCalculation(XlCalculation calculation)
{
    InvokePutProperty("Calculation", long(calculation));
}

XlCalculationInterruptKey wxExcelApplication::GetCalculationInterruptKey()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("CalculationInterruptKey", XlCalculationInterruptKey, xlNoKey);
}

void wxExcelApplication::SetCalculationInterruptKey(XlCalculationInterruptKey calculationInterruptKey)
{
    InvokePutProperty("CalculationInterruptKey", long(calculationInterruptKey));
}

XlCalculationState wxExcelApplication::GetCalculationState()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("CalculationState", XlCalculationState, xlDone);
}

long wxExcelApplication::GetCalculationVersion()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("CalculationVersion");
}

wxString wxExcelApplication::GetCaption()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("Caption");
}

void wxExcelApplication::SetCaption(const wxString& caption)
{
    InvokePutProperty("Caption", caption);
}

bool wxExcelApplication::GetCellDragAndDrop()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("CellDragAndDrop");
}

void wxExcelApplication::SetCellDragAndDrop(bool cellDragAndDrop)
{
    InvokePutProperty("CellDragAndDrop", cellDragAndDrop);
}


wxArrayLong wxExcelApplication::GetClipboardFormats()
{
    wxArrayLong result, emptyResult;

    wxVariant vResult;
    if ( InvokeGetProperty(wxS("ClipboardFormats"), vResult) )
    {
        if (vResult.GetType() == "list") {
            size_t count = vResult.GetCount();
            result.reserve(count);
            for (size_t i = 0; i < count; i++) {
                wxVariant& vItem = vResult[i];
                WXAUTOEXCEL_CHECK_VARIANT_TYPE(vItem, "long", "ClipboardFormats", emptyResult);
                result.push_back(vItem.GetLong());
            }
        } else {
            WXAUTOEXCEL_CHECK_VARIANT_TYPE(vResult, "long", "ClipboardFormats", emptyResult);
            result.push_back(vResult.GetLong());
        }
    }
    return result;
}


bool wxExcelApplication::GetConstrainNumeric()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("ConstrainNumeric");
}

void wxExcelApplication::SetConstrainNumeric(bool constrainNumeric)
{
    InvokePutProperty("ConstrainNumeric", constrainNumeric);
}

bool wxExcelApplication::GetControlCharacters()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("ControlCharacters");
}

void wxExcelApplication::SetControlCharacters(bool controlCharacters)
{
    InvokePutProperty("ControlCharacters", controlCharacters);
}

bool wxExcelApplication::GetCopyObjectsWithCells()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("CopyObjectsWithCells");
}

void wxExcelApplication::SetCopyObjectsWithCells(bool copyObjectsWithCells)
{
    InvokePutProperty("CopyObjectsWithCells", copyObjectsWithCells);
}


XlMousePointer wxExcelApplication::GetCursor()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("Cursor", XlMousePointer, xlDefault);
}

void wxExcelApplication::SetCursor(XlMousePointer cursor)
{
    InvokePutProperty("Cursor", long(cursor));
}

long wxExcelApplication::GetCursorMovement()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("CursorMovement");
}

long wxExcelApplication::GetCustomListCount()
{
    wxVariant vResult;
    long result = 0;
    if ( InvokeGetProperty(wxS("CustomListCount"), vResult) )
    {
        if (vResult.GetType() == "long")
            result = vResult.GetLong();
        else if (vResult.GetType() == "double")
             result = (long)vResult.GetDouble();
    }
    return result;
}

long wxExcelApplication::GetCutCopyMode()
{
    wxVariant vResult;
    long result = 0;
    // CutCopyMode can return an XlCutCopyMode constant or false
    if ( InvokeGetProperty(wxS("CutCopyMode"), vResult) )
    {
        if (vResult.GetType() == "long")
            result = vResult.GetLong();
    }
    return result;
}


void wxExcelApplication::SetCutCopyMode()
{
    InvokePutProperty("CutCopyMode", false);
}

long wxExcelApplication::GetDataEntryMode()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("DataEntryMode");
}

void wxExcelApplication::SetDataEntryMode(long dataEntryMode)
{
    InvokePutProperty("DataEntryMode", dataEntryMode);
}

long wxExcelApplication::GetDDEAppReturnCode()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("DDEAppReturnCode");
}

wxString wxExcelApplication::GetDecimalSeparator()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("DecimalSeparator");
}

void wxExcelApplication::SetDecimalSeparator(const wxString& decimalSeparator)
{
    InvokePutProperty("DecimalSeparator", decimalSeparator);
}

wxString wxExcelApplication::GetDefaultFilePath()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("DefaultFilePath");
}

void wxExcelApplication::SetDefaultFilePath(const wxString& defaultFilePath)
{
    InvokePutProperty("DefaultFilePath", defaultFilePath);
}

XlFileFormat wxExcelApplication::GetDefaultSaveFormat()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("DefaultSaveFormat", XlFileFormat, xlWorkbookNormal);
}

void wxExcelApplication::SetDefaultSaveFormat(XlFileFormat defaultSaveFormat)
{
    InvokePutProperty("DefaultSaveFormat", long(defaultSaveFormat));
}

long wxExcelApplication::GetDefaultSheetDirection()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("DefaultSheetDirection");
}

void wxExcelApplication::SetDefaultSheetDirection(long defaultSheetDirection)
{
    InvokePutProperty("DefaultSheetDirection", defaultSheetDirection);
}

bool wxExcelApplication::GetDisplayAlerts()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("DisplayAlerts");
}

void wxExcelApplication::SetDisplayAlerts(bool displayAlerts)
{
    InvokePutProperty("DisplayAlerts", displayAlerts);
}

XlCommentDisplayMode wxExcelApplication::GetDisplayCommentIndicator()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("DisplayCommentIndicator", XlCommentDisplayMode, xlNoIndicator);
}

void wxExcelApplication::SetDisplayCommentIndicator(XlCommentDisplayMode displayCommentIndicator)
{
    InvokePutProperty("DisplayCommentIndicator", (long)displayCommentIndicator);;
}

bool wxExcelApplication::GetDisplayFormulaAutoComplete()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("DisplayFormulaAutoComplete");
}

void wxExcelApplication::SetDisplayFormulaAutoComplete(bool displayFormulaAutoComplete)
{
    InvokePutProperty("DisplayFormulaAutoComplete", displayFormulaAutoComplete);
}

bool wxExcelApplication::GetDisplayFormulaBar()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("DisplayFormulaBar");
}

void wxExcelApplication::SetDisplayFormulaBar(bool displayFormulaBar)
{
    InvokePutProperty("DisplayFormulaBar", displayFormulaBar);
}

bool wxExcelApplication::GetDisplayFullScreen()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("DisplayFullScreen");
}

void wxExcelApplication::SetDisplayFullScreen(bool displayFullScreen)
{
    InvokePutProperty("DisplayFullScreen", displayFullScreen);
}

bool wxExcelApplication::GetDisplayFunctionToolTips()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("DisplayFunctionToolTips");
}

void wxExcelApplication::SetDisplayFunctionToolTips(bool displayFunctionToolTips)
{
    InvokePutProperty("DisplayFunctionToolTips", displayFunctionToolTips);
}

bool wxExcelApplication::GetDisplayNoteIndicator()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("DisplayNoteIndicator");
}

void wxExcelApplication::SetDisplayNoteIndicator(bool displayNoteIndicator)
{
    InvokePutProperty("DisplayNoteIndicator", displayNoteIndicator);
}

bool wxExcelApplication::GetDisplayScrollBars()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("DisplayScrollBars");
}

void wxExcelApplication::SetDisplayScrollBars(bool displayScrollBars)
{
    InvokePutProperty("DisplayScrollBars", displayScrollBars);
}

bool wxExcelApplication::GetDisplayStatusBar()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("DisplayStatusBar");
}

void wxExcelApplication::SetDisplayStatusBar(bool displayStatusBar)
{
    InvokePutProperty("DisplayStatusBar", displayStatusBar);
}

bool wxExcelApplication::GetEditDirectlyInCell()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("EditDirectlyInCell");
}

void wxExcelApplication::SetEditDirectlyInCell(bool editDirectlyInCell)
{
    InvokePutProperty("EditDirectlyInCell", editDirectlyInCell);
}

bool wxExcelApplication::GetEnableAnimations()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("EnableAnimations");
}

void wxExcelApplication::SetEnableAnimations(bool enableAnimations)
{
    InvokePutProperty("EnableAnimations", enableAnimations);
}

bool wxExcelApplication::GetEnableAutoComplete()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("EnableAutoComplete");
}

void wxExcelApplication::SetEnableAutoComplete(bool enableAutoComplete)
{
    InvokePutProperty("EnableAutoComplete", enableAutoComplete);
}

XlEnableCancelKey wxExcelApplication::GetEnableCancelKey()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("EnableCancelKey", XlEnableCancelKey, xlDisabled);
}

void wxExcelApplication::SetEnableCancelKey(XlEnableCancelKey enableCancelKey)
{
    InvokePutProperty("EnableCancelKey", (long)enableCancelKey);;
}

bool wxExcelApplication::GetEnableLargeOperationAlert()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("EnableLargeOperationAlert");
}

void wxExcelApplication::SetEnableLargeOperationAlert(bool enableLargeOperationAlert)
{
    InvokePutProperty("EnableLargeOperationAlert", enableLargeOperationAlert);
}

bool wxExcelApplication::GetEnableLivePreview()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("EnableLivePreview");
}

void wxExcelApplication::SetEnableLivePreview(bool enableLivePreview)
{
    InvokePutProperty("EnableLivePreview", enableLivePreview);
}

bool wxExcelApplication::GetEnableSound()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("EnableSound");
}

void wxExcelApplication::SetEnableSound(bool enableSound)
{
    InvokePutProperty("EnableSound", enableSound);
}

bool wxExcelApplication::GetExtendList()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("ExtendList");
}

void wxExcelApplication::SetExtendList(bool extendList)
{
    InvokePutProperty("ExtendList", extendList);
}

MsoFeatureInstall wxExcelApplication::GetFeatureInstall()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("FeatureInstall", MsoFeatureInstall, msoFeatureInstallNone);
}

void wxExcelApplication::SetFeatureInstall(MsoFeatureInstall featureInstall)
{
    InvokePutProperty("FeatureInstall", (long)featureInstall);
}


bool wxExcelApplication::GetFixedDecimal()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("FixedDecimal");
}

void wxExcelApplication::SetFixedDecimal(bool fixedDecimal)
{
    InvokePutProperty("FixedDecimal", fixedDecimal);
}

long wxExcelApplication::GetFixedDecimalPlaces()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("FixedDecimalPlaces");
}

void wxExcelApplication::SetFixedDecimalPlaces(long fixedDecimalPlaces)
{
    InvokePutProperty("FixedDecimalPlaces", fixedDecimalPlaces);
}

long wxExcelApplication::GetFormulaBarHeight()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("FormulaBarHeight");
}

void wxExcelApplication::SetFormulaBarHeight(long formulaBarHeight)
{
    InvokePutProperty("FormulaBarHeight", formulaBarHeight);
}

bool wxExcelApplication::GetGenerateGetPivotData()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("GenerateGetPivotData");
}

void wxExcelApplication::SetGenerateGetPivotData(bool generateGetPivotData)
{
    InvokePutProperty("GenerateGetPivotData", generateGetPivotData);
}

XlGenerateTableRefs wxExcelApplication::GetGenerateTableRefs()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("GenerateTableRefs", XlGenerateTableRefs, xlA1TableRefs);
}

void wxExcelApplication::SetGenerateTableRefs(XlGenerateTableRefs generateTableRefs)
{
    InvokePutProperty("GenerateTableRefs", (long)generateTableRefs);
}

double wxExcelApplication::GetHeight()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("Height");
}

void wxExcelApplication::SetHeight(double height)
{
    InvokePutProperty("Height", height);
}

long wxExcelApplication::GetHinstance()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("Hinstance");
}
/*
long wxExcelApplication::GetHwnd()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("Hwnd");
}
*/

bool wxExcelApplication::GetIgnoreRemoteRequests()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("IgnoreRemoteRequests");
}

void wxExcelApplication::SetIgnoreRemoteRequests(bool ignoreRemoteRequests)
{
    InvokePutProperty("IgnoreRemoteRequests", ignoreRemoteRequests);
}

bool wxExcelApplication::GetInteractive()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("Interactive");
}

void wxExcelApplication::SetInteractive(bool interactive)
{
    InvokePutProperty("Interactive", interactive);
}

wxVariant wxExcelApplication::GetInternational(XlApplicationInternational* index)
{
    WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT(Index, (long*)index);

    wxVariant vResult;

    InvokeMethod(wxS("International"), vResult, vIndex);
    return vResult;
}

bool wxExcelApplication::GetIteration()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("Iteration");
}

void wxExcelApplication::SetIteration(bool iteration)
{
    InvokePutProperty("Iteration", iteration);
}


long wxExcelApplication::GetLargeOperationCellThousandCount()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("LargeOperationCellThousandCount");
}

void wxExcelApplication::SetLargeOperationCellThousandCount(long largeOperationCellThousandCount)
{
    InvokePutProperty("LargeOperationCellThousandCount", largeOperationCellThousandCount);
}

double wxExcelApplication::GetLeft()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("Left");
}

void wxExcelApplication::SetLeft(double left)
{
    InvokePutProperty("Left", left);;
}

wxString wxExcelApplication::GetLibraryPath()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("LibraryPath");
}

wxString wxExcelApplication::GetMailSession()
{
    wxString result;
    wxVariant vResult;
    if ( InvokeGetProperty(wxS("MailSession"), vResult) )
    {
        if ( vResult.GetType() == "string" )
            result = vResult.GetString();
    }
    return result;
}

XlMailSystem wxExcelApplication::GetMailSystem()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("MailSystem", XlMailSystem, xlNoMailSystem);
}

bool wxExcelApplication::GetMapPaperSize()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("MapPaperSize");
}

void wxExcelApplication::SetMapPaperSize(bool mapPaperSize)
{
    InvokePutProperty("MapPaperSize", mapPaperSize);
}

bool wxExcelApplication::GetMathCoprocessorAvailable()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("MathCoprocessorAvailable");
}

double wxExcelApplication::GetMaxChange()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("MaxChange");
}

void wxExcelApplication::SetMaxChange(double maxChange)
{
    InvokePutProperty("MaxChange", maxChange);
}


long wxExcelApplication::GetMaxIterations()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("MaxIterations");
}

void wxExcelApplication::SetMaxIterations(long iterations)
{
    InvokePutProperty("Iterations", iterations);
}

XlMeasurementUnits wxExcelApplication::GetMeasurementUnit()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("MeasurementUnit", XlMeasurementUnits, xlInches);
}

void wxExcelApplication::SetMeasurementUnit(XlMeasurementUnits measurementUnits)
{
    InvokePutProperty("MeasurementUnit", (long)measurementUnits);
}


bool wxExcelApplication::GetMouseAvailable()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("MouseAvailable");
}

bool wxExcelApplication::GetMoveAfterReturn()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("MoveAfterReturn");
}

void wxExcelApplication::SetMoveAfterReturn(bool moveAfterReturn)
{
    InvokePutProperty("MoveAfterReturn", moveAfterReturn);
}

XlDirection wxExcelApplication::GetMoveAfterReturnDirection()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("MoveAfterReturnDirection", XlDirection, xlDown);
}

void wxExcelApplication::SetMoveAfterReturnDirection(XlDirection moveAfterReturnDirection)
{
    InvokePutProperty("MoveAfterReturnDirection", (long)moveAfterReturnDirection);;
}

wxString wxExcelApplication::GetNetworkTemplatesPath()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("NetworkTemplatesPath");
}


long wxExcelApplication::GetODBCTimeout()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("ODBCTimeout");
}

void wxExcelApplication::SetODBCTimeout(long oDBCTimeout)
{
    InvokePutProperty("ODBCTimeout", oDBCTimeout);
}


wxString wxExcelApplication::GetOperatingSystem()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("OperatingSystem");
}

wxString wxExcelApplication::GetOrganizationName()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("OrganizationName");
}


wxString wxExcelApplication::GetPath()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("Path");
}

wxString wxExcelApplication::GetPathSeparator()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("PathSeparator");
}

bool wxExcelApplication::GetPivotTableSelection()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("PivotTableSelection");
}

void wxExcelApplication::SetPivotTableSelection(bool pivotTableSelection)
{
    InvokePutProperty("PivotTableSelection", pivotTableSelection);
}

wxString wxExcelApplication::GetProductCode()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("ProductCode");
}

bool wxExcelApplication::GetPromptForSummaryInfo()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("PromptForSummaryInfo");
}

void wxExcelApplication::SetPromptForSummaryInfo(bool promptForSummaryInfo)
{
    InvokePutProperty("PromptForSummaryInfo", promptForSummaryInfo);
}


bool wxExcelApplication::GetReady()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("Ready");
}


XlReferenceStyle wxExcelApplication::GetReferenceStyle()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("ReferenceStyle", XlReferenceStyle, xlA1);
}

void wxExcelApplication::SetReferenceStyle(XlReferenceStyle referenceStyle)
{
    InvokePutProperty("ReferenceStyle", (long)referenceStyle);;
}


bool wxExcelApplication::GetRollZoom()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("RollZoom");
}

void wxExcelApplication::SetRollZoom(bool rollZoom)
{
    InvokePutProperty("RollZoom", rollZoom);
}

bool wxExcelApplication::GetScreenUpdating()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("ScreenUpdating");
}

void wxExcelApplication::SetScreenUpdating(bool screenUpdating)
{
    InvokePutProperty("ScreenUpdating", screenUpdating);
}


wxExcelSheets wxExcelApplication::GetSheets()
{
    wxExcelSheets sheets;
    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("Sheets", sheets);
}

long wxExcelApplication::GetSheetsInNewWorkbook()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("SheetsInNewWorkbook");
}

void wxExcelApplication::SetSheetsInNewWorkbook(long sheetsInNewWorkbook)
{
    InvokePutProperty("SheetsInNewWorkbook", sheetsInNewWorkbook);
}

bool wxExcelApplication::GetShowDevTools()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("ShowDevTools");
}

void wxExcelApplication::SetShowDevTools(bool showDevTools)
{
    InvokePutProperty("ShowDevTools", showDevTools);
}

bool wxExcelApplication::GetShowMenuFloaties()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("ShowMenuFloaties");
}

void wxExcelApplication::SetShowMenuFloaties(bool showMenuFloaties)
{
    InvokePutProperty("ShowMenuFloaties", showMenuFloaties);
}

bool wxExcelApplication::GetShowSelectionFloaties()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("ShowSelectionFloaties");
}

void wxExcelApplication::SetShowSelectionFloaties(bool showSelectionFloaties)
{
    InvokePutProperty("ShowSelectionFloaties", showSelectionFloaties);
}

bool wxExcelApplication::GetShowStartupDialog()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("ShowStartupDialog");
}

void wxExcelApplication::SetShowStartupDialog(bool showStartupDialog)
{
    InvokePutProperty("ShowStartupDialog", showStartupDialog);
}

bool wxExcelApplication::GetShowToolTips()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("ShowToolTips");
}

void wxExcelApplication::SetShowToolTips(bool showToolTips)
{
    InvokePutProperty("ShowToolTips", showToolTips);
}

bool wxExcelApplication::GetShowWindowsInTaskbar()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("ShowWindowsInTaskbar");
}

void wxExcelApplication::SetShowWindowsInTaskbar(bool showWindowsInTaskbar)
{
    InvokePutProperty("ShowWindowsInTaskbar", showWindowsInTaskbar);
}

wxString wxExcelApplication::GetStandardFont()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("StandardFont");
}

void wxExcelApplication::SetStandardFont(const wxString& standardFont)
{
    InvokePutProperty("StandardFont", standardFont);
}

long wxExcelApplication::GetStandardFontSize()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("StandardFontSize");
}

void wxExcelApplication::SetStandardFontSize(long standardFontSize)
{
    InvokePutProperty("StandardFontSize", standardFontSize);
}

wxString wxExcelApplication::GetStartupPath()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("StartupPath");
}

wxString wxExcelApplication::GetStatusBar()
{
    wxString result;
    wxVariant vResult;
    if ( InvokeGetProperty(wxS("StatusBar"), vResult) )
    {
        // Excel can also return false if it has control over the status bar
        if ( vResult.GetType() == "string" )
            result = vResult.GetString();
    }
    return result;
}

void wxExcelApplication::SetStatusBar(const wxString& statusBar)
{
    InvokePutProperty("StatusBar", statusBar);
}

void wxExcelApplication::SetStatusBar()
{
    InvokePutProperty("StatusBar", false);
}

wxString wxExcelApplication::GetTemplatesPath()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("TemplatesPath");
}


wxString wxExcelApplication::GetThousandsSeparator()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("ThousandsSeparator");
}

void wxExcelApplication::SetThousandsSeparator(const wxString& thousandsSeparator)
{
    InvokePutProperty("ThousandsSeparator", thousandsSeparator);
}

double wxExcelApplication::GetTop()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("Top");
}

void wxExcelApplication::SetTop(double top)
{
    InvokePutProperty("Top", top);
}

wxString wxExcelApplication::GetTransitionMenuKey()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("TransitionMenuKey");
}

void wxExcelApplication::SetTransitionMenuKey(const wxString& transitionMenuKey)
{
    InvokePutProperty("TransitionMenuKey", transitionMenuKey);
}

long wxExcelApplication::GetTransitionMenuKeyAction()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("TransitionMenuKeyAction");
}

void wxExcelApplication::SetTransitionMenuKeyAction(long transitionMenuKeyAction)
{
    InvokePutProperty("TransitionMenuKeyAction", transitionMenuKeyAction);
}

bool wxExcelApplication::GetTransitionNavigKeys()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("TransitionNavigKeys");
}

void wxExcelApplication::SetTransitionNavigKeys(bool transitionNavigKeys)
{
    InvokePutProperty("TransitionNavigKeys", transitionNavigKeys);
}

double wxExcelApplication::GetUsableHeight()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("UsableHeight");
}

double wxExcelApplication::GetUsableWidth()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("UsableWidth");
}

bool wxExcelApplication::GetUserControl()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("UserControl");
}

void wxExcelApplication::SetUserControl(bool userControl)
{
    InvokePutProperty("UserControl", userControl);
}

wxString wxExcelApplication::GetUserLibraryPath()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("UserLibraryPath");
}

wxString wxExcelApplication::GetUserName()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("UserName");
}

void wxExcelApplication::SetUserName(const wxString& userName)
{
    InvokePutProperty("UserName", userName);
}

bool wxExcelApplication::GetUseSystemSeparators()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("UseSystemSeparators");
}

void wxExcelApplication::SetUseSystemSeparators(bool useSystemSeparators)
{
    InvokePutProperty("UseSystemSeparators", useSystemSeparators);
}


wxString wxExcelApplication::GetVersion()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("Version");
}

bool wxExcelApplication::GetVisible()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("Visible");
}

void wxExcelApplication::SetVisible(bool visible)
{
    InvokePutProperty("Visible", visible);
}

double wxExcelApplication::GetWidth()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("Width");
}

void wxExcelApplication::SetWidth(double width)
{
    InvokePutProperty("Width", width);
}

wxExcelWindows wxExcelApplication::GetWindows()
{
    wxExcelWindows windows;
    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("Windows", windows);
}

bool wxExcelApplication::GetWindowsForPens()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("WindowsForPens");
}

XlWindowState wxExcelApplication::GetWindowState()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("WindowState", XlWindowState, xlNormal);
}

void wxExcelApplication::SetWindowState(XlWindowState windowState)
{
    InvokePutProperty("WindowState", (long)windowState);;
}


wxExcelWorkbooks wxExcelApplication::GetWorkbooks()
{
    wxExcelWorkbooks workbooks;
    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("Workbooks", workbooks);
}


wxExcelWorksheets wxExcelApplication::GetWorksheets()
{
   wxExcelWorksheets sheets;
   WXAUTOEXCEL_PROPERTY_OBJECT_GET0("Worksheets", sheets);
}

bool wxExcelApplication::RangesToVariants(const wxExcelRangeVector& ranges, wxVariantVector& variants)
{
    wxASSERT( variants.empty() );

    wxVariant variant;
    variants.reserve(ranges.size());
    for ( size_t i = 0; i < ranges.size(); i++ )
    {
        variant.MakeNull();
        if ( !ObjectToVariant(&ranges[i], variant) )
        {
            // we need to decrease the ref count back to what it was
            // because wxAutomationObject::Invoke() which would do that
            // won't be called
            for ( size_t j = 0; j < variants.size(); j++ )            
                ReleaseVariantDispatch(variants[j]);                            
            return false;
        }
        variants.push_back(variant);
    }
    return true;
}

} // namespace wxAutoExcel
