/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2020 PB <pbfordev@gmail.com>
// License:     MIT license
/////////////////////////////////////////////////////////////////////////////

#ifndef _WXAUTOEXCEL_OUTLINE_H
#define _WXAUTOEXCEL_OUTLINE_H

#include "wx/wxAutoExcel_defs.h"

#include "wx/wxAutoExcel_object.h"
#include "wx/wxAutoExcel_enums.h"

namespace wxAutoExcel
{

/**
@brief Represents an outline on a worksheet.
*/
class WXDLLIMPEXP_WXAUTOEXCEL wxExcelOutline : public wxExcelObject
{
public:
    // ***** METHODS *****

    /**
    Displays the specified number of row and/or column levels of an outline.

    [Excel VBA documentation for Outline.ShowLevels](https://docs.microsoft.com/en-us/office/vba/api/excel.outline.showlevels)
    */
    bool ShowLevels(long* rowLevels = NULL, long* columnLevels = NULL);

    // ***** PROPERTIES *****

    /**
    True if the outline uses automatic styles.

    [Excel VBA documentation for Outline.AutomaticStyles](https://docs.microsoft.com/en-us/office/vba/api/excel.outline.automaticstyles)
    */
    bool GetAutomaticStyles();

    /**
    True if the outline uses automatic styles.

    [Excel VBA documentation for Outline.AutomaticStyles](https://docs.microsoft.com/en-us/office/vba/api/excel.outline.automaticstyles)
    */
    void SetAutomaticStyles(bool automaticStyles);

    /**
    Returns the location of the summary columns in the outline. Read/write XlSummaryColumn.

    [Excel VBA documentation for Outline.SummaryColumn](https://docs.microsoft.com/en-us/office/vba/api/excel.outline.summarycolumn)
    */
    XlSummaryColumn GetSummaryColumn();

    /**
    Sets the location of the summary columns in the outline. Read/write XlSummaryColumn.

    [Excel VBA documentation for Outline.SummaryColumn](https://docs.microsoft.com/en-us/office/vba/api/excel.outline.summarycolumn)
    */
    void SetSummaryColumn(XlSummaryColumn summaryColumn);

    /**
    Returns the location of the summary rows in the outline. Read/write XlSummaryRow.

    [Excel VBA documentation for Outline.SummaryRow](https://docs.microsoft.com/en-us/office/vba/api/excel.outline.summaryrow)
    */
    XlSummaryRow GetSummaryRow();

    /**
    Sets the location of the summary rows in the outline. Read/write XlSummaryRow.

    [Excel VBA documentation for Outline.SummaryRow](https://docs.microsoft.com/en-us/office/vba/api/excel.outline.summaryrow)
    */
    void SetSummaryRow(XlSummaryRow summaryRow);

    /**
    Returns "Outline".
    */
    virtual wxString GetAutoExcelObjectName_() const { return wxS("Outline"); }

}; // class wxExcelOutline

} // namespace wxAutoExcel

#endif // #ifndef _WXAUTOEXCEL_OUTLINE_H
