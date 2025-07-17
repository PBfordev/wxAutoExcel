/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2020 PB <pbfordev@gmail.com>
// License:     MIT license
/////////////////////////////////////////////////////////////////////////////

#include "wx/wxAutoExcel_prec.h"

#include "wx/wxAutoExcelOutline.h"

#include "wx/wxAutoExcel_private.h"

namespace wxAutoExcel {

// ***** class wxExcelOutline METHODS *****

bool wxExcelOutline::ShowLevels(long* rowLevels, long* columnLevels)
{
    wxVariant result;

    WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT_NAME(RowLevels, rowLevels);
    WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT_NAME(ColumnLevels, columnLevels);

    WXAUTOEXCEL_CALL_METHOD2("ShowLevels", vRowLevels, vColumnLevels, "bool", result);

    if ( result.IsType("bool") )
        return result.GetBool();

    return false;

}

// ***** class wxExcelOutline PROPERTIES *****

bool wxExcelOutline::GetAutomaticStyles()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("AutomaticStyles");
}

void wxExcelOutline::SetAutomaticStyles(bool automaticStyles)
{
    InvokePutProperty(wxS("AutomaticStyles"), automaticStyles);
}

XlSummaryColumn wxExcelOutline::GetSummaryColumn()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("SummaryColumn", XlSummaryColumn, xlSummaryOnLeft );
}

void wxExcelOutline::SetSummaryColumn(XlSummaryColumn summaryColumn)
{
    InvokePutProperty(wxS("SummaryColumn"), (long)summaryColumn);
}

XlSummaryRow wxExcelOutline::GetSummaryRow()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("SummaryRow", XlSummaryRow, xlSummaryAbove);
}

void wxExcelOutline::SetSummaryRow(XlSummaryRow summaryRow)
{
    InvokePutProperty(wxS("SummaryRow"), (long)summaryRow);
}

} // namespace wxAutoExcel
