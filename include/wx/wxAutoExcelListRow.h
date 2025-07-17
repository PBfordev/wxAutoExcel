/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2020 PB <pbfordev@gmail.com>
// License:     MIT license
/////////////////////////////////////////////////////////////////////////////

#ifndef _WXAUTOEXCEL_LISTROW_H
#define _WXAUTOEXCEL_LISTROW_H

#include "wx/wxAutoExcel_defs.h"

#include "wx/wxAutoExcel_object.h"

namespace wxAutoExcel
{

/**
@brief Represents a row in table, a member of ListRows collection.
*/
class WXDLLIMPEXP_WXAUTOEXCEL wxExcelListRow : public wxExcelObject
{
public:
    // ***** METHODS *****

    /**
    Deletes the cells of the list row and shifts upward any remaining cells below the deleted row.
    You can delete rows in the list even when the list is linked to a SharePoint site.
    The list on the SharePoint site will not be updated, however, until you synchronize your changes.

    [MSDN documentation for ListRow.Delete](https://docs.microsoft.com/en-us/office/vba/api/excel.listrow.delete)
    */
    void Delete();

    // ***** PROPERTIES *****

    /**
    Returns a long that represents the index number of the ListRow object within the ListRows collection.

    [MSDN documentation for ListRow.Index](https://docs.microsoft.com/en-us/office/vba/api/excel.listrow.index)
    */
    long GetIndex();

    /**
    Returns a Range object that represents the range to which the specified list object in the above list applies.

    [MSDN documentation for ListRow.Range](https://docs.microsoft.com/en-us/office/vba/api/excel.listrow.range)
    */
    wxExcelRange GetRange();

    /**
    Returns "ListRow".
    */
    virtual wxString GetAutoExcelObjectName_() const { return wxS("ListRow"); }

}; // class wxExcelListRow


/**
    @brief Represents a collection of ListRow objects.
*/
class WXDLLIMPEXP_WXAUTOEXCEL wxExcelListRows : public wxExcelObject
{
public:
    // ***** METHODS *****

    /**
        Adds a new row to the table.

        [MSDN documentation for ListRows.Add](https://docs.microsoft.com/en-us/office/vba/api/excel.listrows.add)
    */
    wxExcelListRow Add(long* position = NULL, wxXlTribool alwaysInsert = wxDefaultXlTribool);

    // ***** PROPERTIES *****

    /**
        Returns the number of objects in the collection.
    */
    long GetCount();

    //@{
    /**
        Returns the ListRow with the given index or name.

        [MSDN documentation for ListRows.Item](https://docs.microsoft.com/en-us/office/vba/api/excel.listrows.item)
    */
    wxExcelListRow GetItem(long index);
    wxExcelListRow GetItem(const wxString& name);
    wxExcelListRow operator[](long index);
    wxExcelListRow operator[](const wxString& name);
    //@}

    /**
    Returns "ListRows".
    */
    virtual wxString GetAutoExcelObjectName_() const { return wxS("ListRows"); }

}; // class wxExcelListRows

} // namespace wxAutoExcel

#endif // #ifndef _WXAUTOEXCEL_LISTROW_H
