/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2020 PB <pbfordev@gmail.com>
// License:     MIT license
/////////////////////////////////////////////////////////////////////////////

#ifndef _WXAUTOEXCEL_LISTCOLUMN_H
#define _WXAUTOEXCEL_LISTCOLUMN_H

#include "wx/wxAutoExcel_defs.h"

#include "wx/wxAutoExcel_object.h"

namespace wxAutoExcel
{

/**
@brief Represents a column in a table, a member of ListColumns collection.
*/
class WXDLLIMPEXP_WXAUTOEXCEL wxExcelListColumn : public wxExcelObject
{
public:
    // ***** METHODS *****

    /**
    Deletes the column of data in the list.

    [MSDN documentation for ListColumn.Delete](https://docs.microsoft.com/en-us/office/vba/api/excel.listcolumn.delete)
    */
    void Delete();

    // ***** PROPERTIES *****

    /**
    Returns a Range object that is the size of the data portion of a column.

    [MSDN documentation for ListColumn.DataBodyRange](https://docs.microsoft.com/en-us/office/vba/api/excel.listcolumn.databodyrange)
    */
    wxExcelRange GetDataBodyRange();

    /**
    Returns a Long value that represents the index number of the ListColumn object within the ListColumns collection.

    [MSDN documentation for ListColumn.Index](https://docs.microsoft.com/en-us/office/vba/api/excel.listcolumn.index)
    */
    long GetIndex();

    /**
    Returns a String value that represents the name of the list column.

    [MSDN documentation for ListColumn.Name](https://docs.microsoft.com/en-us/office/vba/api/excel.listcolumn.name)
    */
    wxString GetName();

    /**
    Sets a String value that represents the name of the list column.

    [MSDN documentation for ListColumn.Name](https://docs.microsoft.com/en-us/office/vba/api/excel.listcolumn.name)
    */
    void SetName(const wxString& name);

    /**
    Returns a Range object that represents the range to which the specified list object in the above list applies.

    [MSDN documentation for ListColumn.Range](https://docs.microsoft.com/en-us/office/vba/api/excel.listcolumn.range)
    */
    wxExcelRange GetRange();

    /**
    Returns the Total row for a ListColumn object.

    [MSDN documentation for ListColumn.Total](https://docs.microsoft.com/en-us/office/vba/api/excel.listcolumn.total)
    */
    wxExcelRange GetTotal();

    /**
    Determines the type of calculation in the Totals row of the list column based on the value of the XlTotalsCalculation enumeration.

    [MSDN documentation for ListColumn.TotalsCalculation](https://docs.microsoft.com/en-us/office/vba/api/excel.listcolumn.totalscalculation)
    */
    XlTotalsCalculation GetTotalsCalculation();

    /**
    Determines the type of calculation in the Totals row of the list column based on the value of the XlTotalsCalculation enumeration.

    [MSDN documentation for ListColumn.TotalsCalculation](https://docs.microsoft.com/en-us/office/vba/api/excel.listcolumn.totalscalculation).
    */
    void SetTotalsCalculation(XlTotalsCalculation totalsCalculation);

    /**
    Returns "ListColumn".
    */
    virtual wxString GetAutoExcelObjectName_() const { return wxS("ListColumn"); }

}; // class wxExcelListColumn


/**
    @brief Represents a collection of ListColumn objects.
*/
class WXDLLIMPEXP_WXAUTOEXCEL wxExcelListColumns : public wxExcelObject
{
public:
    // ***** METHODS *****

    /**
        Adds a new column to the table.

        @a position is an optional argument that specifies the relative position
        of the new column that starts at 1.

        [MSDN documentation for ListColumns.Add](https://docs.microsoft.com/en-us/office/vba/api/excel.listcolumns.add)
    */
    wxExcelListColumn Add(long* position = NULL);

    // ***** PROPERTIES *****

    /**
        Returns the number of objects in the collection.

        [MSDN documentation for ListColumns.Count](https://docs.microsoft.com/en-us/office/vba/api/excel.listcolumns.count)
    */
    long GetCount();

    //@{
    /**
        Returns the ListColumn with the given index or name.

        [MSDN documentation for ListColumns.Item](https://docs.microsoft.com/en-us/office/vba/api/excel.listcolumns.item)
    */
    wxExcelListColumn GetItem(long index);
    wxExcelListColumn GetItem(const wxString& name);
    wxExcelListColumn operator[](long index);
    wxExcelListColumn operator[](const wxString& name);
    //@}

    /**
    Returns "ListColumns".
    */
    virtual wxString GetAutoExcelObjectName_() const { return wxS("ListColumns"); }

}; // class wxExcelListColumns


} // namespace wxAutoExcel

#endif // #ifndef _WXAUTOEXCEL_LISTCOLUMN_H
