/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2020 PB <pbfordev@gmail.com>
// License:     MIT license
/////////////////////////////////////////////////////////////////////////////

#ifndef _WXAUTOEXCEL_PROTECTION_H
#define _WXAUTOEXCEL_PROTECTION_H

#include "wx/wxAutoExcel_defs.h"

#include "wx/wxAutoExcel_object.h"

namespace wxAutoExcel
{

/**
@brief Represents  the various types of protection options available for a worksheet.
*/
class WXDLLIMPEXP_WXAUTOEXCEL wxExcelProtection : public wxExcelObject
{
public:
    // ***** PROPERTIES *****

    /**
    Returns True if the deletion of columns is allowed on a protected worksheet.

    [Excel VBA documentation for Protection.AllowDeletingColumns](https://docs.microsoft.com/en-us/office/vba/api/excel.protection.allowdeletingcolumns)
    */
    bool GetAllowDeletingColumns();

    /**
    Returns True if the deletion of rows is allowed on a protected worksheet.

    [Excel VBA documentation for Protection.AllowDeletingRows](https://docs.microsoft.com/en-us/office/vba/api/excel.protection.allowdeletingrows)
    */
    bool GetAllowDeletingRows();

    /**
    Returns an AllowEditRanges object.

    [Excel VBA documentation for Protection.AllowEditRanges](https://docs.microsoft.com/en-us/office/vba/api/excel.protection.alloweditranges)
    */
    wxExcelAllowEditRanges GetAllowEditRanges();

    /**
    Returns True if the user is allowed to make use of an AutoFilter that was created before the sheet was protected.

    [Excel VBA documentation for Protection.AllowFiltering](https://docs.microsoft.com/en-us/office/vba/api/excel.protection.allowfiltering)
    */
    bool GetAllowFiltering();

    /**
    Returns True if the formatting of cells is allowed on a protected worksheet.

    [Excel VBA documentation for Protection.AllowFormattingCells](https://docs.microsoft.com/en-us/office/vba/api/excel.protection.allowformattingcells)
    */
    bool GetAllowFormattingCells();

    /**
    Returns True if the formatting of columns is allowed on a protected worksheet.

    [Excel VBA documentation for Protection.AllowFormattingColumns](https://docs.microsoft.com/en-us/office/vba/api/excel.protection.allowformattingcolumns)
    */
    bool GetAllowFormattingColumns();

    /**
    Returns True if the formatting of rows is allowed on a protected worksheet.

    [Excel VBA documentation for Protection.AllowFormattingRows](https://docs.microsoft.com/en-us/office/vba/api/excel.protection.allowformattingrows)
    */
    bool GetAllowFormattingRows();

    /**
    Returns True if the insertion of columns is allowed on a protected worksheet.

    [Excel VBA documentation for Protection.AllowInsertingColumns](https://docs.microsoft.com/en-us/office/vba/api/excel.protection.allowinsertingcolumns)
    */
    bool GetAllowInsertingColumns();

    /**
    Returns True if the insertion of hyperlinks is allowed on a protected worksheet.

    [Excel VBA documentation for Protection.AllowInsertingHyperlinks](https://docs.microsoft.com/en-us/office/vba/api/excel.protection.allowinsertinghyperlinks)
    */
    bool GetAllowInsertingHyperlinks();

    /**
    Returns True if the insertion of rows is allowed on a protected worksheet.

    [Excel VBA documentation for Protection.AllowInsertingRows](https://docs.microsoft.com/en-us/office/vba/api/excel.protection.allowinsertingrows)
    */
    bool GetAllowInsertingRows();

    /**
    Returns True if the sorting option is allowed on a protected worksheet.

    [Excel VBA documentation for Protection.AllowSorting](https://docs.microsoft.com/en-us/office/vba/api/excel.protection.allowsorting)
    */
    bool GetAllowSorting();

    /**
    Returns True if the user is allowed to manipulate pivot tables on a protected worksheet.

    [Excel VBA documentation for Protection.AllowUsingPivotTables](https://docs.microsoft.com/en-us/office/vba/api/excel.protection.allowusingpivottables)
    */
    bool GetAllowUsingPivotTables();

    /**
    Returns "Protection".
    */
    virtual wxString GetAutoExcelObjectName_() const { return wxS("Protection"); }

}; // class wxExcelProtection

} // namespace wxAutoExcel

#endif // #ifndef _WXAUTOEXCEL_PROTECTION_H
