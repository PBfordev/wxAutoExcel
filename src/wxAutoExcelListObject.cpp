/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2020 PB <pbfordev@gmail.com>
// License:     MIT license
/////////////////////////////////////////////////////////////////////////////

#include "wx/wxAutoExcel_prec.h"

#include "wx/wxAutoExcelListObject.h"

#include "wx/wxAutoExcelAutoFilter.h"
#include "wx/wxAutoExcelListColumn.h"
#include "wx/wxAutoExcelListRow.h"
#include "wx/wxAutoExcelRange.h"
#include "wx/wxAutoExcelSort.h"
#include "wx/wxAutoExcelTableObject.h"
#include "wx/wxAutoExcelTableStyle.h"

#include "wx/wxAutoExcel_private.h"

namespace wxAutoExcel {

// ***** class wxExcelListObject METHODS *****

void wxExcelListObject::Delete()
{
    WXAUTOEXCEL_CALL_METHOD0_RET("Delete", "null");
}

void wxExcelListObject::ExportToVisio()
{
    WXAUTOEXCEL_CALL_METHOD0_RET("ExportToVisio", "null");
}

wxString wxExcelListObject::Publish(const wxArrayString& target, bool linkSource)
{
    wxASSERT( target.size() >= 2 && target.size() <= 3 );

    wxVariant vTarget(target, wxS("Target"));
    wxVariant vLinkSource(linkSource, wxS("LinkSource"));

    WXAUTOEXCEL_CALL_METHOD2("Publish", vTarget, vLinkSource, "string", "");
    return vResult.GetString();
}

void wxExcelListObject::Refresh()
{
    WXAUTOEXCEL_CALL_METHOD0_RET("Refresh", "null");
}

void wxExcelListObject::Resize(wxExcelRange range)
{
    wxVariant vRange;

    if ( ObjectToVariant(&range, vRange, wxS("Range")) )
    {
        WXAUTOEXCEL_CALL_METHOD1_RET("Resize", vRange, "null");
    }
}

void wxExcelListObject::Unlink()
{
    WXAUTOEXCEL_CALL_METHOD0_RET("Unlink", "null");
}

void wxExcelListObject::Unlist()
{
    WXAUTOEXCEL_CALL_METHOD0_RET("Unlist", "null");
}

// ***** class wxExcelListObject PROPERTIES *****

bool wxExcelListObject::GetActive()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("Active");
}

wxString wxExcelListObject::GetAlternativeText()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("AlternativeText");
}

void wxExcelListObject::SetAlternativeText(const wxString& alternativeText)
{
    InvokePutProperty(wxS("AlternativeText"), alternativeText);
}

wxExcelAutoFilter wxExcelListObject::GetAutoFilter()
{
    wxExcelAutoFilter autoFilter;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("AutoFilter", autoFilter);
}

wxString wxExcelListObject::GetComment()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("Comment");
}

void wxExcelListObject::SetComment(const wxString& comment)
{
    InvokePutProperty(wxS("Comment"), comment);
}

wxExcelRange wxExcelListObject::GetDataBodyRange()
{
    wxExcelRange range;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("DataBodyRange", range);
}

wxString wxExcelListObject::GetDisplayName()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("DisplayName");
}

void wxExcelListObject::SetDisplayName(const wxString& displayName)
{
    InvokePutProperty(wxS("DisplayName"), displayName);
}

bool wxExcelListObject::GetDisplayRightToLeft()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("DisplayRightToLeft");
}

wxExcelRange wxExcelListObject::GetHeaderRowRange()
{
    wxExcelRange range;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("HeaderRowRange", range);
}

wxExcelRange wxExcelListObject::GetInsertRowRange()
{
    wxExcelRange range;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("InsertRowRange", range);
}

wxExcelListColumns wxExcelListObject::GetListColumns()
{
    wxExcelListColumns listColumns;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("ListColumns", listColumns);
}

wxExcelListRows wxExcelListObject::GetListRows()
{
    wxExcelListRows listRows;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("ListRows", listRows);
}

wxString wxExcelListObject::GetName()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("Name");
}

void wxExcelListObject::SetName(const wxString& name)
{
    InvokePutProperty(wxS("Name"), name);
}

wxExcelRange wxExcelListObject::GetRange()
{
    wxExcelRange range;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("Range", range);
}

wxString wxExcelListObject::GetSharePointURL()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("SharePointURL");
}

bool wxExcelListObject::GetShowAutoFilter()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("ShowAutoFilter");
}

void wxExcelListObject::SetShowAutoFilter(bool showAutoFilter)
{
    InvokePutProperty(wxS("ShowAutoFilter"), showAutoFilter);
}

bool wxExcelListObject::GetShowAutoFilterDropDown()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("ShowAutoFilterDropDown");
}

void wxExcelListObject::SetShowAutoFilterDropDown(bool showAutoFilterDropDown)
{
    InvokePutProperty(wxS("ShowAutoFilterDropDown"), showAutoFilterDropDown);
}

bool wxExcelListObject::GetShowHeaders()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("ShowHeaders");
}

void wxExcelListObject::SetShowHeaders(bool showHeaders)
{
    InvokePutProperty(wxS("ShowHeaders"), showHeaders);
}

bool wxExcelListObject::GetShowTableStyleColumnStripes()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("ShowTableStyleColumnStripes");
}

void wxExcelListObject::SetShowTableStyleColumnStripes(bool showTableStyleColumnStripes)
{
    InvokePutProperty(wxS("ShowTableStyleColumnStripes"), showTableStyleColumnStripes);
}

bool wxExcelListObject::GetShowTableStyleFirstColumn()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("ShowTableStyleFirstColumn");
}

void wxExcelListObject::SetShowTableStyleFirstColumn(bool showTableStyleFirstColumn)
{
    InvokePutProperty(wxS("ShowTableStyleFirstColumn"), showTableStyleFirstColumn);
}

bool wxExcelListObject::GetShowTableStyleLastColumn()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("ShowTableStyleLastColumn");
}

void wxExcelListObject::SetShowTableStyleLastColumn(bool showTableStyleLastColumn)
{
    InvokePutProperty(wxS("ShowTableStyleLastColumn"), showTableStyleLastColumn);
}

bool wxExcelListObject::GetShowTableStyleRowStripes()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("ShowTableStyleRowStripes");
}

void wxExcelListObject::SetShowTableStyleRowStripes(bool showTableStyleRowStripes)
{
    InvokePutProperty(wxS("ShowTableStyleRowStripes"), showTableStyleRowStripes);
}

bool wxExcelListObject::GetShowTotals()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("ShowTotals");
}

void wxExcelListObject::SetShowTotals(bool showTotals)
{
    InvokePutProperty(wxS("ShowTotals"), showTotals);
}

wxExcelSort wxExcelListObject::GetSort()
{
    wxExcelSort sort;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("Sort", sort);
}

void wxExcelListObject::SetSort(wxExcelSort sort)
{
    wxVariant vSort;

    if ( ObjectToVariant(&sort, vSort) )
    {
        InvokePutProperty(wxS("Sort"), vSort);
    }
}

XlListObjectSourceType wxExcelListObject::GetSourceType()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("SourceType", XlListObjectSourceType, xlSrcRange);
}

wxString wxExcelListObject::GetSummary()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("Summary");
}

void wxExcelListObject::SetSummary(const wxString& summary)
{
    InvokePutProperty(wxS("Summary"), summary);
}

wxExcelTableObject wxExcelListObject::GetTableObject()
{
    wxExcelTableObject tableObject;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("TableObject", tableObject);
}

wxExcelTableStyle wxExcelListObject::GetTableStyle()
{
    wxExcelTableStyle tableStyle;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("TableStyle", tableStyle);
}

void wxExcelListObject::SetTableStyle(wxExcelTableStyle tableStyle)
{
    wxVariant vTableStyle;

    if ( ObjectToVariant(&tableStyle, vTableStyle) )
    {
        InvokePutProperty(wxS("TableStyle"), vTableStyle);
    }
}

wxExcelRange wxExcelListObject::GetTotalsRowRange()
{
    wxExcelRange range;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("TotalsRowRange", range);
}


// ***** class wxExcelListObjects METHODS *****

wxExcelListObject wxExcelListObjects::Add(wxExcelRange* source,
                                          XlYesNoGuess* XlListObjectHasHeaders,
                                          const wxString& tableStyleName)
{
    wxExcelListObject object;
    wxVariantVector args;

    args.push_back(wxVariant((long)xlSrcRange, "SourceType"));

    if ( source )
    {
        wxVariant vSource;

        if ( !ObjectToVariant(source, vSource, "Source") )
            return object;

        args.push_back(vSource);
    }

    WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT_NAME(XlListObjectHasHeaders, ((long*)XlListObjectHasHeaders));
    WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT_NAME_VECTOR(XlListObjectHasHeaders, ((long*)XlListObjectHasHeaders), args);
    WXAUTOEXCEL_OPTIONALCPPSTR_TO_OPTIONALVARIANT_NAME_VECTOR(TableStyleName, tableStyleName, args);

    WXAUTOEXCEL_CALL_METHODARR("Add", args, "void*", object);
    VariantToObject(vResult, &object);
    return object;
}

wxExcelListObject wxExcelListObjects::Add(const wxArrayString& source,
                                          wxExcelRange destination,
                                          wxXlTribool linkSource,
                                          XlYesNoGuess* XlListObjectHasHeaders,
                                          const wxString& tableStyleName)
{
    wxExcelListObject object;
    wxVariantVector args;
    wxVariant vDestination;

    wxCHECK_MSG(source.size() == 3, object, "source must contain 3 strings");
    wxCHECK_MSG(destination.IsOk_(), object, "destination must be a valid range");

    args.push_back(wxVariant((long)xlSrcExternal, "SourceType"));
    args.push_back(wxVariant(source, "Source"));

    if ( !ObjectToVariant(&destination, vDestination, "Destination") )
        return object;

    WXAUTOEXCEL_OPTIONALCPPTBOOL_TO_OPTIONALVARIANT_NAME_VECTOR(LinkSource, linkSource, args);
    WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT_NAME_VECTOR(XlListObjectHasHeaders, ((long*)XlListObjectHasHeaders), args);
    WXAUTOEXCEL_OPTIONALCPPSTR_TO_OPTIONALVARIANT_NAME_VECTOR(TableStyleName, tableStyleName, args);

    WXAUTOEXCEL_CALL_METHODARR("Add", args, "void*", object);
    VariantToObject(vResult, &object);
    return object;
}




// ***** class wxExcelListObjects PROPERTIES *****

long wxExcelListObjects::GetCount()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("Count");
}

wxExcelListObject wxExcelListObjects::GetItem(long index)
{
    wxASSERT( index > 0 );

    wxExcelListObject object;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET1("Item", index, object);
}

wxExcelListObject wxExcelListObjects::GetItem(const wxString& name)
{
     wxASSERT( !name.empty() );

    wxExcelListObject object;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET1("Item", name, object);
}

wxExcelListObject wxExcelListObjects::operator[](long index)
{
    return GetItem(index);
}

wxExcelListObject wxExcelListObjects::operator[](const wxString& name)
{
    return GetItem(name);
}

} // namespace wxAutoExcel
