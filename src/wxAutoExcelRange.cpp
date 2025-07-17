/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////


#include "wx/wxAutoExcel_prec.h"

#include "wx/msw/ole/oleutils.h"

#include "wx/wxAutoExcelRange.h"

#include "wx/wxAutoExcelComments.h"
#include "wx/wxAutoExcelCommentsThreaded.h"
#include "wx/wxAutoExcelAreas.h"
#include "wx/wxAutoExcelWorksheet.h"
#include "wx/wxAutoExcelApplication.h"
#include "wx/wxAutoExcelFont.h"
#include "wx/wxAutoExcelCharacters.h"
#include "wx/wxAutoExcelBorders.h"
#include "wx/wxAutoExcelInterior.h"
#include "wx/wxAutoExcelListObject.h"
#include "wx/wxAutoExcelStyles.h"
#include "wx/wxAutoExcelFilters.h"
#include "wx/wxAutoExcelAutoFilter.h"
#include "wx/wxAutoExcelHyperlinks.h"
#include "wx/wxAutoExcelValidation.h"
#include "wx/wxAutoExcelErrors.h"
#include "wx/wxAutoExcelFormatConditions.h"
#include "wx/wxAutoExcelSparklineGroups.h"
#include "wx/wxAutoExcelNames.h"

#include "wx/wxAutoExcel_private.h"

namespace wxAutoExcel {

// ***** class wxAutoExcelRange METHODS *****

bool wxExcelRange::Activate()
{
    WXAUTOEXCEL_CALL_METHOD0_BOOL("Activate");
}

wxExcelComment wxExcelRange::AddComment(const wxString& text)
{
    wxExcelComment comment;
    WXAUTOEXCEL_OPTIONALCPPSTR_TO_OPTIONALVARIANT_NAME(Text, text);

    WXAUTOEXCEL_CALL_METHOD1_OBJECT("AddComment", vText, comment);
}

bool wxExcelRange::AdvancedFilter(enum XlFilterAction action,
                                  wxExcelRange* criteriaRange, wxExcelRange* copyToRange,
                                  wxXlTribool unique)
{
    wxVariant vAction((long)action, wxS("Action"));
    wxVariant vCriteriaRange, vCopyToRange;

    if ( criteriaRange )
    {
        if ( !ObjectToVariant(criteriaRange, vCriteriaRange, wxS("CriteriaRange")) )
            return false;
    }
    if ( copyToRange )
    {
        if ( !ObjectToVariant(copyToRange, vCopyToRange, wxS("CopyToRange")) ) {
            ReleaseVariantDispatch(vCriteriaRange);
            return false;
        }
    }
    WXAUTOEXCEL_OPTIONALCPPTBOOL_TO_OPTIONALVARIANT_NAME(Unique, unique);

    WXAUTOEXCEL_CALL_METHOD4("AdvancedFilter", vAction, vCriteriaRange, vCopyToRange, vUnique, "bool", false);
    return vResult.GetBool();
}

void wxExcelRange::AllocateChanges()
{
    WXAUTOEXCEL_CALL_METHOD0_RET("AllocateChanges", "null");
}


void wxExcelRange::ApplyNames(wxArrayString* names,
                                wxXlTribool ignoreRelativeAbsolute, wxXlTribool useRowColumnNames,
                                wxXlTribool omitColumn, wxXlTribool omitRow,
                                XlApplyNamesOrder* order, wxXlTribool appendLast)
{

    wxVariantVector args;
    wxVariant vNames;
    if ( names && !names->empty() )
    {
        vNames.SetName(wxS("Names"));
        vNames = *names;
    }
    WXAUTOEXCEL_OPTIONALCPPTBOOL_TO_OPTIONALVARIANT_NAME_VECTOR(IgnoreRelativeAbsolute, ignoreRelativeAbsolute, args);
    WXAUTOEXCEL_OPTIONALCPPTBOOL_TO_OPTIONALVARIANT_NAME_VECTOR(UseRowColumnNames, useRowColumnNames, args);
    WXAUTOEXCEL_OPTIONALCPPTBOOL_TO_OPTIONALVARIANT_NAME_VECTOR(OmitColumn, omitColumn, args);
    WXAUTOEXCEL_OPTIONALCPPTBOOL_TO_OPTIONALVARIANT_NAME_VECTOR(OmitRow, omitRow, args);
    WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT_NAME_VECTOR(Order, ((long*)order), args);
    WXAUTOEXCEL_OPTIONALCPPTBOOL_TO_OPTIONALVARIANT_NAME_VECTOR(AppendLast, appendLast, args);

    WXAUTOEXCEL_CALL_METHODARR_RET("ApplyNames", args, "null");
}

void wxExcelRange::ApplyOutlineStyles()
{
    WXAUTOEXCEL_CALL_METHOD0_RET("ApplyOutlineStyles", "null");
}

wxString wxExcelRange::AutoComplete(const wxString& str)
{
    WXAUTOEXCEL_CALL_METHOD1_STRING("AutoComplete", str);
}

bool wxExcelRange::AutoFill(wxExcelRange destination, XlAutoFillType* type)
{
    wxVariant vDestination;
    if ( !ObjectToVariant(&destination, vDestination, wxS("Destination")) )
        return false;

    WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT_NAME(Type, ((long*)type));

    WXAUTOEXCEL_CALL_METHOD2("AutoFill", vDestination, vType, "bool", false);
    return vResult.GetBool();
}

bool wxExcelRange::AutoFilter(long* field, const wxString& criteria1,
                                XlAutoFilterOperator* oper, const wxString& criteria2,
                                wxXlTribool visibleDropDown)
{
    wxVariantVector args;

    WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT_NAME_VECTOR(Field, field, args);
    WXAUTOEXCEL_OPTIONALCPPSTR_TO_OPTIONALVARIANT_NAME_VECTOR(Criteria1, criteria1, args);
    WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT_NAME_VECTOR(Operator, ((long*)oper), args);
    WXAUTOEXCEL_OPTIONALCPPSTR_TO_OPTIONALVARIANT_NAME_VECTOR(Criteria2, criteria2, args);
    WXAUTOEXCEL_OPTIONALCPPTBOOL_TO_OPTIONALVARIANT_NAME_VECTOR(VisibleDropDown, visibleDropDown, args);

    WXAUTOEXCEL_CALL_METHODARR("AutoFilter", args, "bool", false);
    return vResult.GetBool();
}

bool wxExcelRange::AutoFit()
{
    WXAUTOEXCEL_CALL_METHOD0_BOOL("AutoFit");
}

void wxExcelRange::AutoOutline()
{
    WXAUTOEXCEL_CALL_METHOD0_RET("AutoOutline", "null");
}

bool wxExcelRange::BorderAround(XlLineStyle* lineStyle, XlBorderWeight* weight,
                                  long* colorIndex, const wxColour* color)
{
    WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT_NAME(LineStyle, ((long*)lineStyle));
    WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT_NAME(Weight, ((long*)weight));
    WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT_NAME(ColorIndex, colorIndex);

    wxVariant vColor;
    if ( color ) {
        vColor = (long)color->GetRGB();
        vColor.SetName(wxS("Color"));
    }

    WXAUTOEXCEL_CALL_METHOD4("BorderAround", vLineStyle, vWeight, vColorIndex, vColor, "bool", false);
    return vResult.GetBool();
}

void wxExcelRange::Calculate()
{
    WXAUTOEXCEL_CALL_METHOD0_RET("Calculate", "null");
}

void wxExcelRange::CalculateRowMajorOrder()
{
    WXAUTOEXCEL_CALL_METHOD0_RET("CalculateRowMajorOrder", "null");
}

void wxExcelRange::CheckSpelling(const wxString& customDictionary, wxXlTribool ignoreUpperCase,
                                   wxXlTribool alwaysSuggest, MsoLanguageID* spellLang)
{
    WXAUTOEXCEL_OPTIONALCPPSTR_TO_OPTIONALVARIANT_NAME(CustomDictionary, customDictionary);
    WXAUTOEXCEL_OPTIONALCPPTBOOL_TO_OPTIONALVARIANT_NAME(IgnoreUpperCase, ignoreUpperCase);
    WXAUTOEXCEL_OPTIONALCPPTBOOL_TO_OPTIONALVARIANT_NAME(AlwaysSuggest, alwaysSuggest);
    WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT_NAME(SpellLang, ((long*)spellLang));

    WXAUTOEXCEL_CALL_METHOD4_RET("CheckSpelling", vCustomDictionary, vIgnoreUpperCase, vAlwaysSuggest, vSpellLang, "null");
}

bool wxExcelRange::Clear()
{
    WXAUTOEXCEL_CALL_METHOD0_BOOL("Clear");
}

bool wxExcelRange::ClearComments()
{
    WXAUTOEXCEL_CALL_METHOD0_BOOL("ClearComments");
}

bool wxExcelRange::ClearContents()
{
    WXAUTOEXCEL_CALL_METHOD0_BOOL("ClearContents");
}

bool wxExcelRange::ClearFormats()
{
    WXAUTOEXCEL_CALL_METHOD0_BOOL("ClearFormats");
}

bool wxExcelRange::ClearHyperlinks()
{
    WXAUTOEXCEL_CALL_METHOD0_BOOL("ClearHyperlinks");
}

bool wxExcelRange::ClearNotes()
{
    WXAUTOEXCEL_CALL_METHOD0_BOOL("ClearNotes");
}

bool wxExcelRange::ClearOutline()
{
    WXAUTOEXCEL_CALL_METHOD0_BOOL("ClearOutline");
}

wxExcelRange wxExcelRange::ColumnDifferences(wxExcelRange comparison)
{
    wxExcelRange range;
    wxVariant vComparison;

    if ( ObjectToVariant(&comparison, vComparison, wxS("Comparison")) )
    {
        WXAUTOEXCEL_CALL_METHOD1_OBJECT("ColumnDifferences", vComparison, range);
    }
    return range;
}

bool wxExcelRange::Copy(const wxExcelRange* destination)
{
    wxVariant vDestination;

    if ( destination != NULL )
    {
        if ( !ObjectToVariant(destination, vDestination) )
            return false;
    }
    WXAUTOEXCEL_CALL_METHOD1_BOOL("Copy", vDestination);
}


bool wxExcelRange::CopyPicture(XlPictureAppearance* appearance, XlCopyPictureFormat* format)
{
    WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT_NAME(Appearance, ((long*)appearance));
    WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT_NAME(Format, ((long*)format));

    WXAUTOEXCEL_CALL_METHOD2("CopyPicture", vAppearance, vFormat, "bool", false);
    return vResult.GetBool();
}

bool wxExcelRange::CreateNames(wxXlTribool top, wxXlTribool left, wxXlTribool bottom, wxXlTribool right)
{
    WXAUTOEXCEL_OPTIONALCPPTBOOL_TO_OPTIONALVARIANT_NAME(Top, top);
    WXAUTOEXCEL_OPTIONALCPPTBOOL_TO_OPTIONALVARIANT_NAME(Left, left);
    WXAUTOEXCEL_OPTIONALCPPTBOOL_TO_OPTIONALVARIANT_NAME(Bottom, bottom);
    WXAUTOEXCEL_OPTIONALCPPTBOOL_TO_OPTIONALVARIANT_NAME(Right, right);

    WXAUTOEXCEL_CALL_METHOD4("CreateNames", vTop, vLeft, vBottom, vRight, "bool", false);
    return vResult.GetBool();
}

bool wxExcelRange::Cut(const wxExcelRange* destination)
{
    wxVariant vDestination;

    if ( destination != NULL )
    {
        if ( !ObjectToVariant(destination, vDestination) )
            return false;
    }

    WXAUTOEXCEL_CALL_METHOD1_BOOL("Cut", vDestination);
}

bool wxExcelRange::DataSeries(XlRowCol* rowCol, XlDataSeriesType* type, XlDataSeriesDate* date,
                                long* step, const wxVariant& stop, wxXlTribool trend)
{
    wxVariantVector args;

    WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT_NAME_VECTOR(Rowcol, ((long*)rowCol), args);
    WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT_NAME_VECTOR(Type, ((long*)type), args);
    WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT_NAME_VECTOR(Date, ((long*)date), args);
    WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT_NAME_VECTOR(Step, step, args);

    if ( !stop.IsNull() )
    {
        wxVariant vStop(stop);
        vStop.SetName(wxS("Stop"));
        args.push_back(vStop);
    }
    WXAUTOEXCEL_OPTIONALCPPTBOOL_TO_OPTIONALVARIANT_NAME_VECTOR(Trend, trend, args);

    WXAUTOEXCEL_CALL_METHODARR("DataSeries", args, "bool", false);
    return vResult.GetBool();
}

bool wxExcelRange::Delete(XlDeleteShiftDirection* shift)
{
    WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT_NAME(Shift, ((long*)shift));
    WXAUTOEXCEL_CALL_METHOD1_BOOL("Delete", vShift);
}


void wxExcelRange::Dirty()
{
    WXAUTOEXCEL_CALL_METHOD0_RET("Dirty", "null");
}

void wxExcelRange::DiscardChanges()
{
    WXAUTOEXCEL_CALL_METHOD0_RET("DiscardChanges", "null");
}


void wxExcelRange::ExportAsFixedFormat(XlFixedFormatType type, const wxString& fileName,
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

void wxExcelRange::ExportAsFixedFormat(XlFixedFormatType type, const wxVariantVector& optionalArgs)
{
    wxVariantVector args(optionalArgs);

    args.push_back(wxVariant((long)type, wxS("Type")));

    WXAUTOEXCEL_CALL_METHODARR_RET("ExportAsFixedFormat", args, "null");
}


bool wxExcelRange::FillDown()
{
    WXAUTOEXCEL_CALL_METHOD0_BOOL("FillDown");
}

bool wxExcelRange::FillLeft()
{
    WXAUTOEXCEL_CALL_METHOD0_BOOL("FillLeft");
}

bool wxExcelRange::FillRight()
{
    WXAUTOEXCEL_CALL_METHOD0_BOOL("FillRight");
}

bool wxExcelRange::FillUp()
{
    WXAUTOEXCEL_CALL_METHOD0_BOOL("FillUp");
}

void wxExcelRange::FlashFill()
{
    WXAUTOEXCEL_CALL_METHOD0_RET("FlashFill", "null");
}



void wxExcelRange::FunctionWizard()
{
    WXAUTOEXCEL_CALL_METHOD0_RET("FunctionWizard", "null");
}

bool wxExcelRange::Insert(XlInsertShiftDirection* shift)
{
    WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT_NAME(Shift, ((long*)shift));

    WXAUTOEXCEL_CALL_METHOD1_BOOL("Insert", vShift);
}

void wxExcelRange::InsertIndent(long insertAmount)
{
    WXAUTOEXCEL_CALL_METHOD1_RET("InsertIndent", wxVariant(insertAmount), "null");
}

bool wxExcelRange::Justify()
{
    WXAUTOEXCEL_CALL_METHOD0_BOOL("Justify");
}

bool wxExcelRange::ListNames()
{
    WXAUTOEXCEL_CALL_METHOD0_BOOL("ListNames");
}

void wxExcelRange::Merge(wxXlTribool across)
{
    WXAUTOEXCEL_OPTIONALCPPTBOOL_TO_OPTIONALVARIANT_NAME(Across, across);
    WXAUTOEXCEL_CALL_METHOD1_RET("Merge", vAcross, "null");
}

bool wxExcelRange::NavigateArrow(wxXlTribool towardPrecedent, long* arrowNumber, long* linkNumber)
{
    WXAUTOEXCEL_OPTIONALCPPTBOOL_TO_OPTIONALVARIANT_NAME(TowardPrecedent, towardPrecedent);
    WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT_NAME(ArrowNumber, arrowNumber);
    WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT_NAME(LinkNumber, linkNumber);

    WXAUTOEXCEL_CALL_METHOD3("NavigateArrow", vTowardPrecedent, vArrowNumber, vLinkNumber, "bool", false);
    return vResult.GetBool();
}


bool wxExcelRange::PasteSpecial(XlPasteType* paste, XlPasteSpecialOperation* operation,
                                  wxXlTribool skipBlanks, wxXlTribool transpose)
{
    WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT_NAME(Paste, ((long*)paste));
    WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT_NAME(Operation, ((long*)operation));
    WXAUTOEXCEL_OPTIONALCPPTBOOL_TO_OPTIONALVARIANT_NAME(SkipBlanks, skipBlanks);
    WXAUTOEXCEL_OPTIONALCPPTBOOL_TO_OPTIONALVARIANT_NAME(Transpose, transpose);

    WXAUTOEXCEL_CALL_METHOD4("PasteSpecial", vPaste, vOperation, vSkipBlanks, vTranspose, "bool", false);
    return vResult.GetBool();
}


bool wxExcelRange::PrintOut(long* from, long* to, long* copies, wxXlTribool preview, const wxString& activePrinter,
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

bool wxExcelRange::PrintOut(const wxVariantVector& args)
{
    WXAUTOEXCEL_CALL_METHODARR("PrintOut", args, "bool", false);
    return vResult.GetBool();
}

bool wxExcelRange::PrintPreview(wxXlTribool enableChanges)
{
    WXAUTOEXCEL_OPTIONALCPPTBOOL_TO_OPTIONALVARIANT_NAME(EnableChanges, enableChanges);

    WXAUTOEXCEL_CALL_METHOD1_BOOL("PrintPreview", vEnableChanges);
}


void wxExcelRange::RemoveDuplicates(const wxArrayLong& columns, XlYesNoGuess* header)
{
    wxVariant vColumns;

    if ( !columns.empty() )
    {
        vColumns.NullList();
        for (size_t i = 0; i < columns.size(); i++)
            vColumns.Append(columns[i]);
    }
    WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT(Header, ((long*)header));

    WXAUTOEXCEL_CALL_METHOD2_RET("RemoveDuplicates", vColumns, vHeader, "null");
}

void wxExcelRange::RemoveSubtotal()
{
    WXAUTOEXCEL_CALL_METHOD0_RET("RemoveSubtotal", "null");
}


wxExcelRange wxExcelRange::RowDifferences(wxExcelRange comparison)
{
    wxVariant vComparison;
    wxExcelRange range;

    if ( ObjectToVariant(&comparison, vComparison) )
    {
        WXAUTOEXCEL_CALL_METHOD1_OBJECT("RowDifferences", vComparison, range);
    }
    return range;
}

bool wxExcelRange::Select()
{
    WXAUTOEXCEL_CALL_METHOD0_BOOL("Select");
}

bool wxExcelRange::Show()
{
    WXAUTOEXCEL_CALL_METHOD0_BOOL("Show");
}

bool wxExcelRange::ShowDependents(wxXlTribool remove)
{
    WXAUTOEXCEL_OPTIONALCPPTBOOL_TO_OPTIONALVARIANT_NAME(Remove, remove);

    WXAUTOEXCEL_CALL_METHOD1_BOOL("ShowDependents", vRemove);
}

wxExcelRange wxExcelRange::ShowErrors()
{
    wxExcelRange range;
    WXAUTOEXCEL_CALL_METHOD0_OBJECT("ShowErrors", range);
}

bool wxExcelRange::ShowPrecedents(wxXlTribool remove)
{
    WXAUTOEXCEL_OPTIONALCPPTBOOL_TO_OPTIONALVARIANT_NAME(Remove, remove);

    WXAUTOEXCEL_CALL_METHOD1_BOOL("ShowPrecedents", vRemove);

}

wxExcelRange wxExcelRange::SpecialCells(XlCellType type, const wxVariant& value)
{
    wxExcelRange range;

    WXAUTOEXCEL_CALL_METHOD2("SpecialCells", type, value, "void*", range);
    VariantToObject(vResult, &range);
    return range;
}


bool wxExcelRange::Subtotal(XlConsolidationFunction groupBy, XlConsolidationFunction function,
                              wxArrayLong& totalList, wxXlTribool replace, wxXlTribool pageBreaks,
                              XlSummaryRow* summaryBelowData)
{
    wxASSERT(groupBy > 0);
    wxVariant vGroupBy(groupBy, wxS("GroupBy"));

    wxVariant vFunction((long)function, wxS("Function"));

    wxASSERT(!totalList.empty());
    wxVariant vTotalList;
    vTotalList.SetName(wxS("TotalList"));
    vTotalList.NullList();
    for (size_t i = 0; i < totalList.size(); i++)
        vTotalList.Append(totalList[i]);

    WXAUTOEXCEL_OPTIONALCPPTBOOL_TO_OPTIONALVARIANT_NAME(Replace, replace);
    WXAUTOEXCEL_OPTIONALCPPTBOOL_TO_OPTIONALVARIANT_NAME(PageBreaks, pageBreaks);
    WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT_NAME(SummaryBelowData, ((long*)summaryBelowData));

    bool result = false;
    wxVariant vResult;
    if ( InvokeMethod(wxS("Subtotal"), vResult, vGroupBy, vFunction, vTotalList, vReplace, vPageBreaks, vSummaryBelowData) )
    {
        WXAUTOEXCEL_CHECK_VARIANT_TYPE(vResult, "bool", "Subtotal", false);
        result = vResult.GetBool();
    }
    return result;
}

void wxExcelRange::Table(wxExcelRange* rowInput, wxExcelRange* columnInput)
{
    wxVariant vRowInput, vColumnInput;
    if ( rowInput )
    {
        if ( !ObjectToVariant(rowInput, vRowInput, wxS("RowInput")) )
            return;
    }
    if ( columnInput )
    {
        if ( !ObjectToVariant(columnInput, vColumnInput, wxS("ColumnInput")) )
        {
            ReleaseVariantDispatch(vRowInput);
            return;
        }
    }
    WXAUTOEXCEL_CALL_METHOD2_RET("Table", vRowInput, vColumnInput, "null");
}


bool wxExcelRange::Ungroup()
{
    WXAUTOEXCEL_CALL_METHOD0_BOOL("Ungroup");
}

bool wxExcelRange::UnMerge()
{
    WXAUTOEXCEL_CALL_METHOD0_BOOL("UnMerge");
}

// ***** class wxAutoExcelRange PROPERTIES *****

bool wxExcelRange::GetAddIndent()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("AddIndent");
}

void wxExcelRange::SetAddIndent(bool addIndent)
{
    InvokePutProperty(wxS("AddIndent"), addIndent);
}

wxString wxExcelRange::DoGetAddress(const wxString& address, wxXlTribool rowAbsolute, wxXlTribool columnAbsolute,
                    XlReferenceStyle* referenceStyle, wxXlTribool external, wxExcelRange* relativeTo)

{
    WXAUTOEXCEL_OPTIONALCPPTBOOL_TO_OPTIONALVARIANT_NAME(RowAbsolute, rowAbsolute);
    WXAUTOEXCEL_OPTIONALCPPTBOOL_TO_OPTIONALVARIANT_NAME(ColumnAbsolute, columnAbsolute);
    WXAUTOEXCEL_OPTIONALCPPTBOOL_TO_OPTIONALVARIANT_NAME(External, external);
    WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT_NAME(ReferenceStyle, ((long*)referenceStyle));

    wxVariant vRelativeTo;

    if ( relativeTo )
    {
        if ( !ObjectToVariant(relativeTo, vRelativeTo, wxS("RelativeTo")) )
            return wxEmptyString;
    }

    wxVariant vResult;
    if ( InvokeGetProperty(address, vResult,
            vRowAbsolute, vColumnAbsolute, vReferenceStyle, vExternal, vRelativeTo) )
    {
        return vResult.GetString();
    }
    return wxEmptyString;
}


wxString wxExcelRange::GetAddress(wxXlTribool rowAbsolute, wxXlTribool columnAbsolute,
                    XlReferenceStyle* referenceStyle, wxXlTribool external, wxExcelRange* relativeTo)

{
    return DoGetAddress(wxS("Address"), rowAbsolute, columnAbsolute, referenceStyle, external, relativeTo);
}

wxString wxExcelRange::GetAddressLocal(wxXlTribool rowAbsolute, wxXlTribool columnAbsolute,
                                       XlReferenceStyle* referenceStyle, wxXlTribool external, wxExcelRange* relativeTo)

{
    return DoGetAddress(wxS("AddressLocal"), rowAbsolute, columnAbsolute, referenceStyle, external, relativeTo);
}

bool wxExcelRange::GetAllowEdit()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("AllowEdit");
}


wxExcelAreas wxExcelRange::GetAreas()
{
    wxExcelAreas areas;
    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("Areas", areas);
}

wxExcelBorders wxExcelRange::GetBorders()
{
    wxExcelBorders borders;
    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("Borders", borders);
}


wxExcelCharacters wxExcelRange::GetCharacters(long start, long* length)
{
    wxExcelCharacters characters;
    WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT(Length, length);
    WXAUTOEXCEL_PROPERTY_OBJECT_GET2("Characters", start, vLength, characters);
}

long wxExcelRange::GetColumn()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("Column");
}

double wxExcelRange::GetColumnWidth()
{
    double result = -2;
    wxVariant vResult;

    if ( InvokeGetProperty(wxS("ColumnWidth"), vResult) )
    {
        if ( vResult.IsNull() )
            result = -1;
        else
            result = vResult;
    }

    return result;
}

void wxExcelRange::SetColumnWidth(double colWidth)
{
    InvokePutProperty(wxS("ColumnWidth"), colWidth);
}


wxExcelComment wxExcelRange::GetComment()
{
    wxExcelComment comment;
    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("Comment", comment);
}

wxExcelCommentThreaded wxExcelRange::GetCommentThreaded()
{
    wxExcelCommentThreaded comment;
    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("CommentThreaded", comment);
}

long wxExcelRange::GetCount()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("Count");
}

wxLongLong wxExcelRange::GetCountLarge()
{
    wxLongLong result = 0;
    wxVariant vResult;

    if ( InvokeGetProperty(wxS("CountLarge"), vResult) )
    {
        result = vResult.GetLongLong();
    }

    return result;
}


wxExcelRange wxExcelRange::GetCurrentArray()
{
    wxExcelRange range;
    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("CurrentArray", range);
}

wxExcelRange wxExcelRange::GetCurrentRegion()
{
    wxExcelRange range;
    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("CurrentRegion", range);
}

wxExcelRange wxExcelRange::GetDependents()
{
    wxExcelRange range;
    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("Dependents", range);
}

wxExcelRange wxExcelRange::GetDirectDependents()
{
    wxExcelRange range;
    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("DirectDependents", range);
}

wxExcelRange wxExcelRange::GetDirectPrecedents()
{
    wxExcelRange range;
    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("DirectPrecedents", range);
}

wxExcelRange wxExcelRange::GetEnd(XlDirection direction)
{
    wxExcelRange range;
    WXAUTOEXCEL_PROPERTY_OBJECT_GET1("End", wxVariant(long(direction)), range);
}

wxExcelRange wxExcelRange::GetEntireColumn()
{
    wxExcelRange range;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("EntireColumn", range);
}


wxExcelRange wxExcelRange::GetEntireRow()
{
    wxExcelRange range;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("EntireRow", range);
}

wxExcelErrors wxExcelRange::GetErrors()
{
    wxExcelErrors errors;
    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("Errors", errors);
}



wxExcelFont wxExcelRange::GetFont()
{
    wxExcelFont font;
    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("Font", font);
}

#if WXAUTOEXCEL_USE_CONDFORMAT

wxExcelFormatConditions wxExcelRange::GetFormatConditions()
{
    wxExcelFormatConditions conditions;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("FormatConditions", conditions);
}

#endif  // WXAUTOEXCEL_USE_CONDFORMAT

wxString wxExcelRange::GetFormula()
{
   WXAUTOEXCEL_PROPERTY_STRING_GET0("Formula");
}

void wxExcelRange::SetFormula(const wxString& formula)
{
    InvokePutProperty(wxS("Formula"), formula);
}

wxString wxExcelRange::GetFormulaArray()
{
   wxVariant vResult;
   if ( InvokeGetProperty(wxS("Formula"), vResult) )
   {
       // can also return NULL if the range is not in an array
       if ( vResult.GetType() == wxS("string") )
            return vResult.GetString();
   }
   return wxS("");
}

void wxExcelRange::SetFormulaArray(const wxString& formula)
{
    InvokePutProperty(wxS("FormulaArray"), formula);
}

wxXlTribool wxExcelRange::GetFormulaHidden()
{
    wxXlTribool tb;
    wxVariant vResult;

    if ( InvokeGetProperty(wxS("FormulaHidden"), vResult) && vResult.GetType() == wxS("bool") )
    {
        tb = vResult.GetBool();
    }
    return tb;
}

void wxExcelRange::SetFormulaHidden(bool hidden)
{
    InvokePutProperty(wxS("FormulaHidden"), hidden);
}

wxString wxExcelRange::GetFormulaLocal()
{
   WXAUTOEXCEL_PROPERTY_STRING_GET0("FormulaLocal");
}

void wxExcelRange::SetFormulaLocal(const wxString& formula)
{
    InvokePutProperty(wxS("FormulaLocal"), formula);
}

wxString wxExcelRange::GetFormulaR1C1()
{
   WXAUTOEXCEL_PROPERTY_STRING_GET0("FormulaR1C1");
}

void wxExcelRange::SetFormulaR1C1(const wxString& formula)
{
    InvokePutProperty(wxS("FormulaR1C1"), formula);
}

wxString wxExcelRange::GetFormulaR1C1Local()
{
   WXAUTOEXCEL_PROPERTY_STRING_GET0("FormulaR1C1Local");
}

void wxExcelRange::SetFormulaR1C1Local(const wxString& formula)
{
    InvokePutProperty(wxS("FormulaR1C1Local"), formula);
}


wxXlTribool wxExcelRange::GetHasArray()
{
    wxXlTribool tb;
    wxVariant vResult;

    if ( InvokeGetProperty(wxS("HasArray"), vResult) && vResult.GetType() == wxS("bool") )
    {
        tb = vResult.GetBool();
    }
    return tb;
}

wxXlTribool wxExcelRange::GetHasFormula()
{
    wxVariant vResult;
    wxXlTribool result;
    if ( InvokeGetProperty(wxS("HasFormula"), vResult) && vResult.GetType() == wxS("bool") )
    {
        result = vResult.GetBool();
    }
    return result;
}

double wxExcelRange::GetHeight()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("Height");
}

void wxExcelRange::SetHeight(double height)
{
    InvokePutProperty(wxS("Height"), height);
}

bool wxExcelRange::GetHidden()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("Hidden");
}

void wxExcelRange::SetHidden(bool hidden)
{
    InvokePutProperty(wxS("Hidden"), hidden);
}

long wxExcelRange::GetHorizontalAlignment()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("HorizontalAlignment");
}

void wxExcelRange::SetHorizontalAlignment(long alignment)
{
    InvokePutProperty(wxS("HorizontalAlignment"), alignment);
}

wxExcelHyperlinks wxExcelRange::GetHyperlinks()
{
    wxExcelHyperlinks hyperlinks;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("Hyperlinks", hyperlinks);
}

wxString wxExcelRange::GetID()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("ID");
}

void wxExcelRange::SetID(const wxString& ID)
{
    InvokePutProperty(wxS("ID"), ID);
}

long wxExcelRange::GetIndentLevel()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("IndentLevel");
}

void wxExcelRange::SetIndentLevel(long indentLevel)
{
    InvokePutProperty(wxS("IndentLevel"), indentLevel);
}

wxExcelInterior  wxExcelRange::GetInterior()
{
    wxExcelInterior  interior;
    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("Interior", interior);
}

wxExcelRange wxExcelRange::GetItem(long rowIndex, long* columnIndex)
{
    return GetCells(&rowIndex, columnIndex);
}

wxExcelRange wxExcelRange::GetItem(long rowIndex, const wxString& columnIndex)
{
    return DoGetRangeItem(rowIndex, columnIndex);
}

wxExcelRange wxExcelRange::GetItem(const wxString& cell)
{
    return GetRange(cell);
}


double wxExcelRange::GetLeft()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("Left");
}


long wxExcelRange::GetListHeaderRows()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("ListHeaderRows");
}

wxExcelListObject wxExcelRange::GetListObject()
{
    wxExcelListObject object;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("ListObject", object);
}

XlLocationInTable wxExcelRange::GetLocationInTable()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("LocationInTable", XlLocationInTable, (XlLocationInTable)0);
}

wxXlTribool wxExcelRange::GetLocked()
{
    wxXlTribool tb;
    wxVariant vResult;

    if ( InvokeGetProperty(wxS("Locked"), vResult) && vResult.GetType() == wxS("bool") )
    {
        tb = vResult.GetBool();
    }
    return tb;
}

void wxExcelRange::SetLocked(bool locked)
{
    InvokePutProperty(wxS("Locked"), locked);
}

wxString wxExcelRange::GetMDX()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("MDX");
}

wxExcelRange wxExcelRange::GetMergeArea()
{
    wxExcelRange range;
    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("MergeArea", range);
}

bool wxExcelRange::GetMergeCells()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("MergeCells");
}

wxExcelName wxExcelRange::GetName()
{
    // Supress the error in case the range has no Name
    wxAutoExcelObjectErrorModeOverrider emo(0, true);

    wxExcelName name;
    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("Name", name);
}

void wxExcelRange::SetName(const wxString& name)
{
    InvokePutProperty(wxS("Name"), name);
}

wxExcelRange wxExcelRange::GetNext()
{
    wxExcelRange range;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("Next", range);
}

wxString wxExcelRange::GetNumberFormat()
{
    wxVariant vResult;

    // NumberFormat returns either the number format string
    // if all the cells in the range have the same number format;
    // or null if they have not.
    if ( InvokeGetProperty(wxS("NumberFormat"), vResult) )
    {
        if ( vResult.IsType(wxS("string")) )
            return vResult.GetString();
    }

    return wxEmptyString;
}

void wxExcelRange::SetNumberFormat(const wxString& numberFormat)
{
    InvokePutProperty(wxS("NumberFormat"), numberFormat);
}

wxString wxExcelRange::GetNumberFormatLocal()
{
    wxVariant vResult;

    // NumberFormatLocal returns either the number format string
    // if all the cells in the range have the same number format;
    // or null if they have not.
    if ( InvokeGetProperty(wxS("NumberFormatLocal"), vResult) )
    {
        if ( vResult.IsType(wxS("string")) )
            return vResult.GetString();;
    }

    return wxEmptyString;
}

void wxExcelRange::SetNumberFormatLocal(const wxString& numberFormatLocal)
{
    InvokePutProperty(wxS("NumberFormatLocal"), numberFormatLocal);
}

wxExcelRange wxExcelRange::GetOffset(long rowOffset, long columnOffset)
{
    wxExcelRange range;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET2("Offset", wxVariant(rowOffset), columnOffset, range);
}


long wxExcelRange::GetOrientation()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("Orientation");
}

void wxExcelRange::SetOrientation(long orientation)
{
    InvokePutProperty(wxS("Orientation"), orientation);
}

long wxExcelRange::GetOutlineLevel()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("OutlineLevel");
}

void wxExcelRange::SetOutlineLevel(long outlineLevel)
{
    InvokePutProperty(wxS("OutlineLevel"), outlineLevel);
}

XlPageBreak wxExcelRange::GetPageBreak()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("PageBreak", XlPageBreak, (XlPageBreak)0);
}

void wxExcelRange::SetPageBreak(XlPageBreak pageBreak)
{
    InvokePutProperty(wxS("PageBreak"), (long)pageBreak);
}


wxExcelRange wxExcelRange::GetPrecedents()
{
    wxExcelRange range;
    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("Precedents", range);
}

wxString wxExcelRange::GetPrefixCharacter()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("PrefixCharacter");
}

wxExcelRange wxExcelRange::GetPrevious()
{
    wxExcelRange range;
    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("Previous", range);
}


long wxExcelRange::GetReadingOrder()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("ReadingOrder");
}

wxExcelRange wxExcelRange::GetResize(long* rowSize, long* columnSize)
{
    wxExcelRange range;
    WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT_NAME(RowSize, rowSize);
    WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT_NAME(ColumnSize, columnSize);

    WXAUTOEXCEL_PROPERTY_OBJECT_GET2("Resize", vRowSize, vColumnSize, range);
}



long wxExcelRange::GetRow()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("Row");
}

double wxExcelRange::GetRowHeight()
{
    double result = -2;
    wxVariant vResult;

    if ( InvokeGetProperty(wxS("RowHeight"), vResult) )
    {
        if ( vResult.IsNull() )
            result = -1;
        else
            result = vResult;
    }

    return result;
}

void wxExcelRange::SetRowHeight(double rowHeight)
{
    InvokePutProperty(wxS("RowHeight"), rowHeight);
}


bool wxExcelRange::GetShowDetail()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("ShowDetail");
}

wxXlTribool wxExcelRange::GetShrinkToFit()
{
    wxXlTribool tb;
    wxVariant vResult;

    if ( InvokeGetProperty(wxS("ShrinkToFit"), vResult) && vResult.GetType() == wxS("bool") )
    {
        tb = vResult.GetBool();
    }
    return tb;
}

void wxExcelRange::SetShrinkToFit(bool shrinkToFit)
{
    InvokePutProperty(wxS("ShrinkToFit"), shrinkToFit);
}

#if WXAUTOEXCEL_USE_CHARTS

wxExcelSparklineGroups wxExcelRange::GetSparklineGroups()
{
    wxExcelSparklineGroups groups;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("SparklineGroups", groups);
}

#endif  // WXAUTOEXCEL_USE_CHARTS


wxExcelStyle wxExcelRange::GetStyle()
{
    wxExcelStyle style;
    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("Style", style);
}

void wxExcelRange::SetStyle(wxExcelStyle style)
{
    wxVariant vStyle;

    if ( ObjectToVariant(&style, vStyle) )
    {
        InvokePutProperty(wxS("Style"), vStyle);
    }
}

bool wxExcelRange::GetSummary()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("Summary");
}

wxString wxExcelRange::GetText()
{
    wxString result;
    wxVariant vResult;

    if ( InvokeGetProperty(wxS("Text"), vResult) && vResult.GetType() == wxS("string") )
        result = vResult.GetString();

    return result;
}

double wxExcelRange::GetTop()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("Top");
}

wxXlTribool wxExcelRange::GetUseStandardHeight()
{
    wxXlTribool tb;
    wxVariant vResult;

    if ( InvokeGetProperty(wxS("UseStandardHeight"), vResult) && vResult.GetType() == wxS("bool"))
    {
        tb = vResult.GetBool();
    }
    return tb;
}



void wxExcelRange::SetUseStandardHeight(bool useStandardHeight)
{
    InvokePutProperty(wxS("UseStandardHeight"), useStandardHeight);
}

wxXlTribool wxExcelRange::GetUseStandardWidth()
{
    wxXlTribool tb;
    wxVariant vResult;

    if ( InvokeGetProperty(wxS("UseStandardWidth"), vResult) && vResult.GetType() == wxS("bool") )
    {
        tb = vResult.GetBool();
    }
    return tb;
}

void wxExcelRange::SetUseStandardWidth(bool useStandardWidth)
{
    InvokePutProperty(wxS("UseStandardWidth"), useStandardWidth);
}


wxExcelValidation wxExcelRange::GetValidation()
{
    wxExcelValidation validation;
    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("Validation", validation);
}

wxVariant wxExcelRange::GetValue()
{
    wxVariant vResult;

    InvokeGetProperty(wxS("Value"), vResult);
    return vResult;
}

void wxExcelRange::SetValue(const wxVariant& value)
{
    InvokePutProperty(wxS("Value"), value);
}

wxVariant wxExcelRange::GetValue2()
{
    wxVariant vResult;

    InvokeGetProperty(wxS("Value2"), vResult);
    return vResult;
}

void wxExcelRange::SetValue2(const wxVariant& value)
{
    InvokePutProperty(wxS("Value2"), value);
}

long wxExcelRange::GetVerticalAlignment()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("VerticalAlignment");
}

void wxExcelRange::SetVerticalAlignment(long verticalAlignment)
{
    InvokePutProperty(wxS("VerticalAlignment"), verticalAlignment);
}

double wxExcelRange::GetWidth()
{
    WXAUTOEXCEL_PROPERTY_DOUBLE_GET0("Width");
}

wxExcelWorksheet wxExcelRange::GetWorksheet()
{
    wxExcelWorksheet worksheet;
    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("Worksheet", worksheet);
}

wxXlTribool wxExcelRange::GetWrapText()
{
    wxXlTribool tb;
    wxVariant vResult;

    if ( InvokeGetProperty(wxS("WrapText"), vResult) && vResult.GetType() == wxS("bool"))
    {
        tb = vResult.GetBool();
    }
    return tb;
}

void wxExcelRange::SetWrapText(bool wrapText)
{
    InvokePutProperty(wxS("WrapText"), wrapText);
}


long wxExcelRange::GetConvertVariantFlags_()
{
    wxCHECK( m_xlObject && m_xlObject->GetDispatchPtr(), wxOleConvertVariant_Default );

    return m_xlObject->GetConvertVariantFlags();
}

bool wxExcelRange::SetConvertVariantFlags_(long flags)
{
    wxCHECK( m_xlObject && m_xlObject->GetDispatchPtr(), false);

    m_xlObject->SetConvertVariantFlags(flags);
    return true;
}


} // namespace wxAutoExcel
