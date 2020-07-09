/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Modified by:
// Copyright:   (c) 2020 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////

#ifndef _WXAUTOEXCEL_TABLEOBJECT_H
#define _WXAUTOEXCEL_TABLEOBJECT_H

#include "wx/wxAutoExcel_defs.h"

#include "wx/wxAutoExcel_object.h"

namespace wxAutoExcel
{

/**
@brief Represents a worksheet table built from data returned from a PowerPivot model.
*/
class WXDLLIMPEXP_WXAUTOEXCEL wxExcelTableObject : public wxExcelObject
{
public:
    // ***** METHODS *****

    /**
    Deletes the object.

    [MSDN documentation for TableObject.Delete](https://docs.microsoft.com/en-us/office/vba/api/excel.tableobject.delete)
    */
    void Delete();

    /**
    Updates the object.

    [MSDN documentation for TableObject.Refresh](https://docs.microsoft.com/en-us/office/vba/api/excel.tableobject.refresh)
    */
    bool Refresh();

    // ***** PROPERTIES *****

    /**
    Specifies if the column widths are automatically adjusted for the best fit each time you refresh the specified query table. The default value is True.

    [MSDN documentation for TableObject.AdjustColumnWidth](https://docs.microsoft.com/en-us/office/vba/api/excel.tableobject.adjustcolumnwidth)
    */
    bool GetAdjustColumnWidth();

    /**
    Specifies if the column widths are automatically adjusted for the best fit each time you refresh the specified query table. The default value is True.

    [MSDN documentation for TableObject.AdjustColumnWidth](https://docs.microsoft.com/en-us/office/vba/api/excel.tableobject.adjustcolumnwidth).
    */
    void SetAdjustColumnWidth(bool adjustColumnWidth);

    /**
    Returns the cell in the upper-left corner of the query table destination range (the range where the resulting query table will be placed). The destination range must be on the worksheet that contains the TableObject object.

    [MSDN documentation for TableObject.Destination](https://docs.microsoft.com/en-us/office/vba/api/excel.tableobject.destination)
    */
    wxExcelRange GetDestination();

    /**
    True if the user can edit the specified query table. False if the user can only refresh the query table.

    [MSDN documentation for TableObject.EnableEditing](https://docs.microsoft.com/en-us/office/vba/api/excel.tableobject.enableediting)
    */
    bool GetEnableEditing();

    /**
    True if the user can edit the specified query table. False if the user can only refresh the query table.

    [MSDN documentation for TableObject.EnableEditing](https://docs.microsoft.com/en-us/office/vba/api/excel.tableobject.enableediting).
    */
    void SetEnableEditing(bool enableEditing);

    /**
    Specifies if the query table can be refreshed by the user.

    [MSDN documentation for TableObject.EnableRefresh](https://docs.microsoft.com/en-us/office/vba/api/excel.tableobject.enablerefresh)
    */
    bool GetEnableRefresh();

    /**
    Specifies if the query table can be refreshed by the user.

    [MSDN documentation for TableObject.EnableRefresh](https://docs.microsoft.com/en-us/office/vba/api/excel.tableobject.enablerefresh).
    */
    void SetEnableRefresh(bool enableRefresh);

    /**
    Specifies if the number of rows returned by the last use of the Refresh method is greater than the number of rows available on the worksheet.

    [MSDN documentation for TableObject.FetchedRowOverflow](https://docs.microsoft.com/en-us/office/vba/api/excel.tableobject.fetchedrowoverflow)
    */
    bool GetFetchedRowOverflow();

    /**
    Returns a ListObject Object (Excel) object for the TableObject Object (Excel) object.

    [MSDN documentation for TableObject.ListObject](https://docs.microsoft.com/en-us/office/vba/api/excel.tableobject.listobject)
    */
    wxExcelListObject GetListObject();

    /**
    Specifies if column sorting, filtering, and layout information is preserved whenever a query table is refreshed. The default value is False.

    [MSDN documentation for TableObject.PreserveColumnInfo](https://docs.microsoft.com/en-us/office/vba/api/excel.tableobject.preservecolumninfo)
    */
    bool GetPreserveColumnInfo();

    /**
    Specifies if column sorting, filtering, and layout information is preserved whenever a query table is refreshed. The default value is False.

    [MSDN documentation for TableObject.PreserveColumnInfo](https://docs.microsoft.com/en-us/office/vba/api/excel.tableobject.preservecolumninfo).
    */
    void SetPreserveColumnInfo(bool preserveColumnInfo);

    /**
    True if any formatting common to the first five rows of data are applied to new rows of data in the query table. Unused cells aren't formatted. The property is False if the last AutoFormat applied to the query table is applied to new rows of data. The default value is True.

    [MSDN documentation for TableObject.PreserveFormatting](https://docs.microsoft.com/en-us/office/vba/api/excel.tableobject.preserveformatting)
    */
    bool GetPreserveFormatting();

    /**
    True if any formatting common to the first five rows of data are applied to new rows of data in the query table. Unused cells aren't formatted. The property is False if the last AutoFormat applied to the query table is applied to new rows of data. The default value is True.

    [MSDN documentation for TableObject.PreserveFormatting](https://docs.microsoft.com/en-us/office/vba/api/excel.tableobject.preserveformatting).
    */
    void SetPreserveFormatting(bool preserveFormatting);

    /**
    Returns the way rows on the specified worksheet are added or deleted to accommodate the number of rows in a record set returned by a query. XlCellInsertionMode Enumeration (Excel) Read/Write

    [MSDN documentation for TableObject.RefreshStyle](https://docs.microsoft.com/en-us/office/vba/api/excel.tableobject.refreshstyle)
    */
    XlCellInsertionMode GetRefreshStyle();

    /**
    Sets the way rows on the specified worksheet are added or deleted to accommodate the number of rows in a record set returned by a query. XlCellInsertionMode Enumeration (Excel) Read/Write

    [MSDN documentation for TableObject.RefreshStyle](https://docs.microsoft.com/en-us/office/vba/api/excel.tableobject.refreshstyle).
    */
    void SetRefreshStyle(XlCellInsertionMode refreshStyle);

    /**
    Returns a Range Object (Excel) object that represents the area of the worksheet occupied by the specified query table.

    [MSDN documentation for TableObject.ResultRange](https://docs.microsoft.com/en-us/office/vba/api/excel.tableobject.resultrange)
    */
    wxExcelRange GetResultRange();

    /**
    Specifies if row numbers are added as the first column of the specified query table.

    [MSDN documentation for TableObject.RowNumbers](https://docs.microsoft.com/en-us/office/vba/api/excel.tableobject.rownumbers)
    */
    bool GetRowNumbers();

    /**
    Specifies if row numbers are added as the first column of the specified query table.

    [MSDN documentation for TableObject.RowNumbers](https://docs.microsoft.com/en-us/office/vba/api/excel.tableobject.rownumbers).
    */
    void SetRowNumbers(bool rowNumbers);

    /**
    Returns "TableObject".
    */
    virtual wxString GetAutoExcelObjectName_() const { return wxS("TableObject"); }

}; // class wxExcelTableObject

} // namespace wxAutoExcel

#endif // #ifndef _WXAUTOEXCEL_TABLEOBJECT_H
