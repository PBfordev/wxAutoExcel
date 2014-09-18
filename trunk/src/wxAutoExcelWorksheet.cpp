/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Modified by:
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////


#include "wx/wxAutoExcel_prec.h"

#include "wx/wxAutoExcelWorksheet.h"

#include <wx/dynarray.h>

#include "wx/wxAutoExcelApplication.h"
#include "wx/wxAutoExcelAutoFilter.h"
#include "wx/wxAutoExcelChartObjects.h"
#include "wx/wxAutoExcelComments.h"
#include "wx/wxAutoExcelHyperlinks.h"
#include "wx/wxAutoExcelOLEObjects.h"
#include "wx/wxAutoExcelPageBreaks.h"
#include "wx/wxAutoExcelPageSetup.h"
#include "wx/wxAutoExcelRange.h"
#include "wx/wxAutoExcelShapes.h"
#include "wx/wxAutoExcelSheet.h"
#include "wx/wxAutoExcelTab.h"

#include "wx/wxAutoExcelPrivate.h"

namespace wxAutoExcel {


// ***** class wxAutoExcelWorksheet METHODS *****

bool wxExcelWorksheet::Activate()
{
     WXAUTOEXCEL_CALL_METHOD0_BOOL("Activate");
}

bool wxExcelWorksheet::Calculate()
{
     WXAUTOEXCEL_CALL_METHOD0_BOOL("Calculate");
}

#if WXAUTOEXCEL_USE_CHARTS

wxExcelChartObjects wxExcelWorksheet::ChartObjects()
{
    wxExcelChartObjects objects;

    WXAUTOEXCEL_CALL_METHOD0_OBJECT("ChartObjects", objects);
}

wxExcelChartObjects wxExcelWorksheet::ChartObjects(const wxVector<long>& indices)
{
    wxExcelChartObjects objects;

    wxCHECK(indices.size() > 0, objects);

    wxVariant vIndices;

    vIndices.NullList();
    for (size_t i = 0; i < indices.size(); i++)
        vIndices.Append(indices[i]);

    WXAUTOEXCEL_CALL_METHOD1_OBJECT("ChartObjects", vIndices, objects);
}

wxExcelChartObjects wxExcelWorksheet::ChartObjects(const wxArrayString& names)
{
    wxExcelChartObjects objects;

    wxCHECK(names.size() > 0, objects);

    wxVariant vNames;

    vNames.NullList();
    for (size_t i = 0; i < names.size(); i++)
        vNames.Append(names[i]);

    WXAUTOEXCEL_CALL_METHOD1_OBJECT("ChartObjects", vNames, objects);
}

#endif // #if WXAUTOEXCEL_USE_CHARTS

bool wxExcelWorksheet::CheckSpelling(const wxString& customDictionary, wxXlTribool ignoreUpperCase,
                                        wxXlTribool alwaysSuggest, MsoLanguageID* spellLang)
{
    WXAUTOEXCEL_OPTIONALCPPSTR_TO_OPTIONALVARIANT_NAME(CustomDictionary, customDictionary);
    WXAUTOEXCEL_OPTIONALCPPTBOOL_TO_OPTIONALVARIANT_NAME(IgnoreUpperCase, ignoreUpperCase);
    WXAUTOEXCEL_OPTIONALCPPTBOOL_TO_OPTIONALVARIANT_NAME(AlwaysSuggest, alwaysSuggest);
    WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT_NAME(SpellLang, ((long*)spellLang));

    WXAUTOEXCEL_CALL_METHOD4("CheckSpelling", vCustomDictionary, vIgnoreUpperCase, vAlwaysSuggest, vSpellLang, "bool", false);
    return vResult.GetBool();
}

void wxExcelWorksheet::CircleInvalid()
{
     WXAUTOEXCEL_CALL_METHOD0_RET("CircleInvalid", "null");
}

bool wxExcelWorksheet::ClearArrows()
{
     WXAUTOEXCEL_CALL_METHOD0_BOOL("ClearArrows");
}

bool wxExcelWorksheet::ClearCircles()
{
     WXAUTOEXCEL_CALL_METHOD0_BOOL("ClearCircles");
}

bool wxExcelWorksheet::Copy()
{
    WXAUTOEXCEL_CALL_METHOD0_BOOL("Copy");
}

bool wxExcelWorksheet::CopyAfterOrBefore(wxExcelSheet worksheetAfterOrBefore, bool after)
{
    return DoOrderedCopyOrMove(true, worksheetAfterOrBefore, after);
}


bool wxExcelWorksheet::Delete()
{
     WXAUTOEXCEL_CALL_METHOD0_BOOL("Delete");
}

bool wxExcelWorksheet::ExportAsFixedFormat(XlFixedFormatType type, const wxString& fileName,
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

    return ExportAsFixedFormat(type, args);
}

bool wxExcelWorksheet::ExportAsFixedFormat(XlFixedFormatType type, const wxVariantVector& optionalArgs)
{
    wxVariantVector args(optionalArgs);

    args.push_back(wxVariant((long)type, wxS("Type")));

    WXAUTOEXCEL_CALL_METHODARR("ExportAsFixedFormat", args, "bool", false);
    return vResult.GetBool();
}

bool wxExcelWorksheet::Move()
{
    WXAUTOEXCEL_CALL_METHOD0_BOOL("Move");
}

bool wxExcelWorksheet::MoveAfterOrBefore(wxExcelSheet sheetAfterOrBefore, bool after)
{
    return DoOrderedCopyOrMove(false, sheetAfterOrBefore, after);
}

wxExcelOLEObjects wxExcelWorksheet::OLEObjects()
{
    wxExcelOLEObjects objects;
    WXAUTOEXCEL_CALL_METHOD0_OBJECT("OLEObjects", objects);
}

bool wxExcelWorksheet::Paste(wxExcelRange* destination, wxXlTribool link)
{
    wxVariant vArgument;
    if ( destination )
    {
        if ( !ObjectToVariant(destination, vArgument, wxS("Destination")) )
            return false;
    }
    else if ( !link.IsDefault() )
    {
        vArgument = link.IsTrue();
    }
    WXAUTOEXCEL_CALL_METHOD1_BOOL("Paste", vArgument);
}

bool wxExcelWorksheet::PasteSpecial(const wxString& format, wxXlTribool link,
                                      wxXlTribool displayAsIcon, const wxString& iconFileName,
                                      long* iconIndex, const wxString& iconLabel,
                                      wxXlTribool noHTMLFormatting)
{
    wxVariantVector args;

    WXAUTOEXCEL_OPTIONALCPPSTR_TO_OPTIONALVARIANT_NAME_VECTOR(Format, format, args);
    WXAUTOEXCEL_OPTIONALCPPTBOOL_TO_OPTIONALVARIANT_NAME_VECTOR(Link, link, args);
    WXAUTOEXCEL_OPTIONALCPPTBOOL_TO_OPTIONALVARIANT_NAME_VECTOR(DisplayAsIcon, displayAsIcon, args);
    WXAUTOEXCEL_OPTIONALCPPSTR_TO_OPTIONALVARIANT_NAME_VECTOR(IconFileName, iconFileName, args);
    WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT_NAME_VECTOR(IconIndex, iconIndex, args);
    WXAUTOEXCEL_OPTIONALCPPSTR_TO_OPTIONALVARIANT_NAME_VECTOR(IconLabel, iconLabel, args);
    WXAUTOEXCEL_OPTIONALCPPTBOOL_TO_OPTIONALVARIANT_NAME_VECTOR(NoHTMLFormatting, noHTMLFormatting, args);

    WXAUTOEXCEL_CALL_METHODARR("PasteSpecial", args, "bool", false);
    return vResult.GetBool();
}

bool wxExcelWorksheet::PrintOut(long* from, long* to, long* copies, wxXlTribool preview, const wxString& activePrinter,
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

bool wxExcelWorksheet::PrintOut(const wxVariantVector& args)
{
    WXAUTOEXCEL_CALL_METHODARR("PrintOut", args, "bool", false);
    return vResult.GetBool();
}

bool wxExcelWorksheet::PrintPreview(wxXlTribool enableChanges)
{
    WXAUTOEXCEL_OPTIONALCPPTBOOL_TO_OPTIONALVARIANT_NAME(EnableChanges, enableChanges);

    WXAUTOEXCEL_CALL_METHOD1_BOOL("PrintPreview", vEnableChanges);
}

void wxExcelWorksheet::Protect(const wxString& password, wxXlTribool structure, wxXlTribool windows)
{
    WXAUTOEXCEL_OPTIONALCPPSTR_TO_OPTIONALVARIANT_NAME(Password, password);
    WXAUTOEXCEL_OPTIONALCPPTBOOL_TO_OPTIONALVARIANT_NAME(Structure, structure);
    WXAUTOEXCEL_OPTIONALCPPTBOOL_TO_OPTIONALVARIANT_NAME(Windows, windows);

    WXAUTOEXCEL_CALL_METHOD3_RET("Protect", vPassword, vStructure, vWindows, "null");
}

void wxExcelWorksheet::ResetAllPageBreaks()
{
     WXAUTOEXCEL_CALL_METHOD0_RET("ResetAllPageBreaks", "null");
}

bool wxExcelWorksheet::SaveAs(const wxString& fileName, XlFileFormat* fileFormat,
                const wxString& password, const wxString& writeResPassword,
                wxXlTribool readOnlyRecommended, wxXlTribool createBackup,
                wxXlTribool addToMru, wxXlTribool local)
{
    wxVariantVector args;

    WXAUTOEXCEL_OPTIONALCPPSTR_TO_OPTIONALVARIANT_NAME_VECTOR(FileName, fileName, args);
    WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT_NAME_VECTOR(FileFormat, ((long*)fileFormat), args);
    WXAUTOEXCEL_OPTIONALCPPSTR_TO_OPTIONALVARIANT_NAME_VECTOR(Password, password, args);
    WXAUTOEXCEL_OPTIONALCPPSTR_TO_OPTIONALVARIANT_NAME_VECTOR(WriteResPassword, writeResPassword, args);
    WXAUTOEXCEL_OPTIONALCPPTBOOL_TO_OPTIONALVARIANT_NAME_VECTOR(ReadOnlyRecommended, readOnlyRecommended, args);
    WXAUTOEXCEL_OPTIONALCPPTBOOL_TO_OPTIONALVARIANT_NAME_VECTOR(CreateBackup, createBackup, args);
    WXAUTOEXCEL_OPTIONALCPPTBOOL_TO_OPTIONALVARIANT_NAME_VECTOR(AddToMru, addToMru, args);
    WXAUTOEXCEL_OPTIONALCPPTBOOL_TO_OPTIONALVARIANT_NAME_VECTOR(Local, local, args);

    return SaveAs(args);
}

bool wxExcelWorksheet::SaveAs(const wxVariantVector& optionalArgs)
{
    WXAUTOEXCEL_CALL_METHODARR("SaveAs", optionalArgs, "bool", false);
    return vResult.GetBool();
}


bool wxExcelWorksheet::Select(wxXlTribool replace)
{
    WXAUTOEXCEL_OPTIONALCPPTBOOL_TO_OPTIONALVARIANT_NAME(Replace, replace);
    WXAUTOEXCEL_CALL_METHOD1_BOOL("Select", vReplace);
}

void wxExcelWorksheet::SetBackgroundPicture(const wxString& fileName)
{
     WXAUTOEXCEL_CALL_METHOD1_RET("SetBackgroundPicture", fileName, "null");
}

void wxExcelWorksheet::ShowAllData()
{
    WXAUTOEXCEL_CALL_METHOD0_RET("ShowAllData", "null");
}

void wxExcelWorksheet::ShowDataForm()
{
     WXAUTOEXCEL_CALL_METHOD0_RET("ShowDataForm", "null");
}

void wxExcelWorksheet::Unprotect(const wxString& password)
{
    WXAUTOEXCEL_OPTIONALCPPSTR_TO_OPTIONALVARIANT_NAME(Password, password);

    WXAUTOEXCEL_CALL_METHOD1_RET("Unprotect", vPassword, "null");
}


// ***** class wxAutoExcelWorksheet PROPERTIES *****

wxExcelAutoFilter wxExcelWorksheet::GetAutoFilter()
{
    wxExcelAutoFilter autoFilter;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("AutoFilter", autoFilter);
}

bool wxExcelWorksheet::GetAutoFilterMode()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("AutoFilterMode");
}

void wxExcelWorksheet::SetAutoFilterMode()
{
    InvokePutProperty("AutoFilterMode", false);
}

wxExcelRange wxExcelWorksheet::GetCircularReference()
{
    wxExcelRange range;
    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("CircularReference", range);
}

wxExcelComments wxExcelWorksheet::GetComments()
{
    wxExcelComments comments;
    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("Comments", comments);
}

XlConsolidationFunction wxExcelWorksheet::GetConsolidationFunction()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("ConsolidationFunction", XlConsolidationFunction, xlUnknown);
}

wxArrayShort wxExcelWorksheet::GetConsolidationOptions()
{
    wxArrayShort result;
    wxVariant vResult;

    if ( InvokeGetProperty(wxS("ConsolidationOptions"), vResult) )
    {
        size_t count = vResult.GetCount();
        wxASSERT(count == 3);
        for (size_t i = 0; i < count; i++)
            result.push_back(vResult[i].GetBool() ? 1 : 0);
    }
    return result;
}

wxArrayString wxExcelWorksheet::GetConsolidationSources()
{
    wxArrayString result;
    wxVariant vResult;

    if ( InvokeGetProperty(wxS("ConsolidationSources"), vResult) )
    {
        size_t count = vResult.GetCount();
        for (size_t i = 0; i < count; i++)
            result.push_back(vResult[i].GetString());
    }
    return result;
}


bool wxExcelWorksheet::GetDisplayPageBreaks()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("DisplayPageBreaks");
}

void wxExcelWorksheet::SetDisplayPageBreaks(bool displayPageBreaks)
{
    InvokePutProperty("DisplayPageBreaks", displayPageBreaks);
}

bool wxExcelWorksheet::GetDisplayRightToLeft()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("DisplayRightToLeft");
}

bool wxExcelWorksheet::GetEnableAutoFilter()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("EnableAutoFilter");
}

void wxExcelWorksheet::SetEnableAutoFilter(bool enableAutoFilter)
{
    InvokePutProperty("EnableAutoFilter", enableAutoFilter);
}

bool wxExcelWorksheet::GetEnableCalculation()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("EnableCalculation");
}

void wxExcelWorksheet::SetEnableCalculation(bool enableCalculation)
{
    InvokePutProperty("EnableCalculation", enableCalculation);
}

bool wxExcelWorksheet::GetEnableFormatConditionsCalculation()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("EnableFormatConditionsCalculation");
}

void wxExcelWorksheet::SetEnableFormatConditionsCalculation(bool enableFormatConditionsCalculation)
{
    InvokePutProperty("EnableFormatConditionsCalculation", enableFormatConditionsCalculation);
}

bool wxExcelWorksheet::GetEnableOutlining()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("EnableOutlining");
}

void wxExcelWorksheet::SetEnableOutlining(bool enableOutlining)
{
    InvokePutProperty("EnableOutlining", enableOutlining);
}

bool wxExcelWorksheet::GetEnablePivotTable()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("EnablePivotTable");
}

void wxExcelWorksheet::SetEnablePivotTable(bool enablePivotTable)
{
    InvokePutProperty("EnablePivotTable", enablePivotTable);
}

XlEnableSelection wxExcelWorksheet::GetEnableSelection()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("EnableSelection", XlEnableSelection, xlNoSelection);
}

bool wxExcelWorksheet::GetFilterMode()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("FilterMode");
}

wxExcelPageBreaks wxExcelWorksheet::GetHPageBreaks()
{
    wxExcelPageBreaks pageBreaks;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("HPageBreaks", pageBreaks);
}

wxExcelHyperlinks  wxExcelWorksheet::GetHyperlinks()
{
    wxExcelHyperlinks hyperlinks;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("Hyperlinks", hyperlinks);
}

long wxExcelWorksheet::GetIndex()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("Index");
}


wxString wxExcelWorksheet::GetName()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("Name");
}

void wxExcelWorksheet::SetName(const wxString& name)
{
    InvokePutProperty("Name", name);
}


wxExcelWorksheet wxExcelWorksheet::GetNext()
{
    wxExcelWorksheet worksheet;
    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("Next", worksheet);
}


wxExcelPageSetup wxExcelWorksheet::GetPageSetup()
{
    wxExcelPageSetup pageSetup;
    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("PageSetup", pageSetup);    
}

wxExcelWorksheet wxExcelWorksheet::GetPrevious()
{
    wxExcelWorksheet worksheet;
    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("Previous", worksheet);
}

long wxExcelWorksheet::GetPrintedCommentPages()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("PrintedCommentPages");
}


bool wxExcelWorksheet::GetProtectContents()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("ProtectContents");
}

bool wxExcelWorksheet::GetProtectDrawingObjects()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("ProtectDrawingObjects");
}

bool wxExcelWorksheet::GetProtectionMode()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("ProtectionMode");
}

bool wxExcelWorksheet::GetProtectScenarios()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("ProtectScenarios");
}

wxString wxExcelWorksheet::GetScrollArea()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("ScrollArea");
}

void wxExcelWorksheet::SetScrollArea(const wxString& scrollArea)
{
    InvokePutProperty("ScrollArea", scrollArea);
}

#if WXAUTOEXCEL_USE_SHAPES

wxExcelShapes wxExcelWorksheet::GetShapes()
{
    wxExcelShapes shapes;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("Shapes", shapes);
}

#endif // #if WXAUTOEXCEL_USE_SHAPES

double wxExcelWorksheet::GetStandardHeight()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("StandardHeight");
}


double wxExcelWorksheet::GetStandardWidth()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("StandardWidth");
}

void wxExcelWorksheet::SetStandardWidth(double standardWidth)
{
    InvokePutProperty("StandardWidth", standardWidth);
}

wxExcelTab wxExcelWorksheet::GetTab()
{
    wxExcelTab tab;
    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("Tab", tab);
}

bool wxExcelWorksheet::GetTransitionExpEval()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("TransitionExpEval");
}

void wxExcelWorksheet::SetTransitionExpEval(bool transitionExpEval)
{
    InvokePutProperty("TransitionExpEval", transitionExpEval);
}

bool wxExcelWorksheet::GetTransitionFormEntry()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("TransitionFormEntry");
}

void wxExcelWorksheet::SetTransitionFormEntry(bool transitionFormEntry)
{
    InvokePutProperty("TransitionFormEntry", transitionFormEntry);
}

XlSheetType wxExcelWorksheet::GetType()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("Type", XlSheetType, xlWorksheet);
}

wxExcelRange wxExcelWorksheet::GetUsedRange()
{
    wxExcelRange range;
    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("UsedRange", range);
}

XlSheetVisibility wxExcelWorksheet::GetVisible()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("Visible", XlSheetVisibility, xlSheetVisible);
}

void wxExcelWorksheet::SetVisible(XlSheetVisibility visible)
{
    InvokePutProperty("Visible", (long)visible);;
}

wxExcelPageBreaks  wxExcelWorksheet::GetVPageBreaks()
{
    wxExcelPageBreaks pageBreaks ;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("VPageBreaks", pageBreaks);
}

wxExcelWorksheet::operator wxExcelSheet()
{
    wxExcelSheet sheet;

    CloneDispatch(this, &sheet);
    return sheet;
}

bool wxExcelWorksheet::DoOrderedCopyOrMove(bool copy, wxExcelSheet sheetAfterOrBefore, bool after)
{
    wxVariant vAfterOrBefore;

    if ( ObjectToVariant(&sheetAfterOrBefore, vAfterOrBefore) )
    {
        vAfterOrBefore.SetName(after? wxS("After") : wxS("Before"));
        WXAUTOEXCEL_CALL_METHOD1_BOOL(copy ? "Copy" : "Move", vAfterOrBefore);
    }
    return false;
}

} // namespace wxAutoExcel
