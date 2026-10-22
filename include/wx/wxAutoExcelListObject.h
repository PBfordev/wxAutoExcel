/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2020 PB <pbfordev@gmail.com>
// License:     MIT license
/////////////////////////////////////////////////////////////////////////////

#ifndef _WXAUTOEXCEL_LISTOBJECT_H
#define _WXAUTOEXCEL_LISTOBJECT_H

#include "wx/wxAutoExcel_defs.h"

#include "wx/wxAutoExcel_object.h"

namespace wxAutoExcel
{

/**
@brief Represents a list object in ListObjects collection.
*/
class WXDLLIMPEXP_WXAUTOEXCEL wxExcelListObject : public wxExcelObject
{
public:
    // ***** METHODS *****

    /**
    Deletes the ListObject object and clears the cell data from the worksheet.

    [MSDN documentation for ListObject.Delete](https://learn.microsoft.com/en-us/office/vba/api/excel.listobject.delete)
    */
    void Delete();

    /**
    Exports a ListObject object to Visio.

    [MSDN documentation for ListObject.ExportToVisio](https://learn.microsoft.com/en-us/office/vba/api/excel.listobject.exporttovisio)
    */
    void ExportToVisio();

    /**
    Publishes the ListObject object to a server that is running Microsoft SharePoint Foundation.

    [MSDN documentation for ListObject.Publish](https://learn.microsoft.com/en-us/office/vba/api/excel.listobject.publish)
    */
    wxString Publish(const wxArrayString& target, bool linkSource);

    /**
    Retrieves the current data and schema for the list from the server that is running Microsoft SharePoint Foundation. This method can be used only with lists that are linked to a SharePoint site. If the SharePoint site is not available, calling this method will return an error.

    [MSDN documentation for ListObject.Refresh](https://learn.microsoft.com/en-us/office/vba/api/excel.listobject.refresh)
    */
    void Refresh();

    /**
    The Resize method allows a ListObject object to be resized over a new range. No cells are inserted or moved.

    [MSDN documentation for ListObject.Resize](https://learn.microsoft.com/en-us/office/vba/api/excel.listobject.resize)
    */
    void Resize(wxExcelRange range);

    /**
    Removes the link to a Microsoft SharePoint Foundation site from a list. Returns Nothing.

    [MSDN documentation for ListObject.Unlink](https://learn.microsoft.com/en-us/office/vba/api/excel.listobject.unlink)
    */
    void Unlink();

    /**
    Removes the list functionality. After you use this method, the range of cells that made up the the list will be a regular range of data.

    [MSDN documentation for ListObject.Unlist](https://learn.microsoft.com/en-us/office/vba/api/excel.listobject.unlist)
    */
    void Unlist();

    // ***** PROPERTIES *****

    /**
    Returns a Boolean value indicating whether a ListObject object in a worksheet is active—that is, whether the active cell is inside the range of the ListObject object.

    [MSDN documentation for ListObject.Active](https://learn.microsoft.com/en-us/office/vba/api/excel.listobject.active)
    */
    bool GetActive();

    /**
    Returns the descriptive (alternative) text string for the specified table.

    [MSDN documentation for ListObject.AlternativeText](https://learn.microsoft.com/en-us/office/vba/api/excel.listobject.alternativetext)
    */
    wxString GetAlternativeText();

    /**
    Sets the descriptive (alternative) text string for the specified table.

    [MSDN documentation for ListObject.AlternativeText](https://learn.microsoft.com/en-us/office/vba/api/excel.listobject.alternativetext).
    */
    void SetAlternativeText(const wxString& alternativeText);

    /**
    Filters a table using the AutoFilter feature.

    [MSDN documentation for ListObject.AutoFilter](https://learn.microsoft.com/en-us/office/vba/api/excel.listobject.autofilter)
    */
    wxExcelAutoFilter GetAutoFilter();

    /**
    Returns the associated comment.

    [MSDN documentation for ListObject.Comment](https://learn.microsoft.com/en-us/office/vba/api/excel.listobject.comment)
    */
    wxString GetComment();

    /**
    Sets the associated comment.

    [MSDN documentation for ListObject.Comment](https://learn.microsoft.com/en-us/office/vba/api/excel.listobject.comment).
    */
    void SetComment(const wxString& comment);

    /**
    Returns a Range that represents the range of values, excluding the header row, in a table.

    [MSDN documentation for ListObject.DataBodyRange](https://learn.microsoft.com/en-us/office/vba/api/excel.listobject.databodyrange)
    */
    wxExcelRange GetDataBodyRange();

    /**
    Returns the display name.

    [MSDN documentation for ListObject.DisplayName](https://learn.microsoft.com/en-us/office/vba/api/excel.listobject.displayname)
    */
    wxString GetDisplayName();

    /**
    Sets the display name.

    [MSDN documentation for ListObject.DisplayName](https://learn.microsoft.com/en-us/office/vba/api/excel.listobject.displayname).
    */
    void SetDisplayName(const wxString& displayName);

    /**
    True if it is displayed from right to left instead of from left to right, false otherwise.

    [MSDN documentation for ListObject.DisplayRightToLeft](https://learn.microsoft.com/en-us/office/vba/api/excel.listobject.displayrighttoleft)
    */
    bool GetDisplayRightToLeft();

    /**
    Returns a Range object that represents the range of the header row for a list.

    [MSDN documentation for ListObject.HeaderRowRange](https://learn.microsoft.com/en-us/office/vba/api/excel.listobject.headerrowrange)
    */
    wxExcelRange GetHeaderRowRange();

    /**
    Returns a Range object representing the Insert row, if any.

    [MSDN documentation for ListObject.InsertRowRange](https://learn.microsoft.com/en-us/office/vba/api/excel.listobject.insertrowrange)
    */
    wxExcelRange GetInsertRowRange();

    /**
    Returns a ListColumns collection that represents all the columns.

    [MSDN documentation for ListObject.ListColumns](https://learn.microsoft.com/en-us/office/vba/api/excel.listobject.listcolumns)
    */
    wxExcelListColumns GetListColumns();

    /**
    Returns a ListRows object that represents all the rows of data.

    [MSDN documentation for ListObject.ListRows](https://learn.microsoft.com/en-us/office/vba/api/excel.listobject.listrows)
    */
    wxExcelListRows GetListRows();

    /**
    Returns a String value that represents the name of the ListObject object.

    [MSDN documentation for ListObject.Name](https://learn.microsoft.com/en-us/office/vba/api/excel.listobject.name)
    */
    wxString GetName();

    /**
    Sets a String value that represents the name of the ListObject object.

    [MSDN documentation for ListObject.Name](https://learn.microsoft.com/en-us/office/vba/api/excel.listobject.name).
    */
    void SetName(const wxString& name);

    /**
    Returns a Range object that represents the range to which the specified list object in the above list applies.

    [MSDN documentation for ListObject.Range](https://learn.microsoft.com/en-us/office/vba/api/excel.listobject.range)
    */
    wxExcelRange GetRange();

    /**
    Returns a String representing the URL of the SharePoint list for a given ListObject object.

    [MSDN documentation for ListObject.SharePointURL](https://learn.microsoft.com/en-us/office/vba/api/excel.listobject.sharepointurl)
    */
    wxString GetSharePointURL();

    /**
    Returns Boolean to indicate whether the AutoFilter will be displayed.

    [MSDN documentation for ListObject.ShowAutoFilter](https://learn.microsoft.com/en-us/office/vba/api/excel.listobject.showautofilter)
    */
    bool GetShowAutoFilter();

    /**
    Returns Boolean to indicate whether the AutoFilter will be displayed.

    [MSDN documentation for ListObject.ShowAutoFilter](https://learn.microsoft.com/en-us/office/vba/api/excel.listobject.showautofilter).
    */
    void SetShowAutoFilter(bool showAutoFilter);

    /**
    True when the AutoFilter drop down for the ListObject object is displayed.

    [MSDN documentation for ListObject.ShowAutoFilterDropDown](https://learn.microsoft.com/en-us/office/vba/api/excel.listobject.showautofilterdropdown)
    */
    bool GetShowAutoFilterDropDown();

    /**
    True when the AutoFilter drop down for the ListObject object is displayed.

    [MSDN documentation for ListObject.ShowAutoFilterDropDown](https://learn.microsoft.com/en-us/office/vba/api/excel.listobject.showautofilterdropdown).
    */
    void SetShowAutoFilterDropDown(bool showAutoFilterDropDown);

    /**
    Returns if the header information should be displayed.

    [MSDN documentation for ListObject.ShowHeaders](https://learn.microsoft.com/en-us/office/vba/api/excel.listobject.showheaders)
    */
    bool GetShowHeaders();

    /**
    Sets if the header information should be displayed.

    [MSDN documentation for ListObject.ShowHeaders](https://learn.microsoft.com/en-us/office/vba/api/excel.listobject.showheaders).
    */
    void SetShowHeaders(bool showHeaders);

    /**
    Returns if the Column Stripes table style is used.

    [MSDN documentation for ListObject.ShowTableStyleColumnStripes](https://learn.microsoft.com/en-us/office/vba/api/excel.listobject.showtablestylecolumnstripes)
    */
    bool GetShowTableStyleColumnStripes();

    /**
    Sets if the Column Stripes table style is used.

    [MSDN documentation for ListObject.ShowTableStyleColumnStripes](https://learn.microsoft.com/en-us/office/vba/api/excel.listobject.showtablestylecolumnstripes).
    */
    void SetShowTableStyleColumnStripes(bool showTableStyleColumnStripes);

    /**
    Returns if the first column is formatted.

    [MSDN documentation for ListObject.ShowTableStyleFirstColumn](https://learn.microsoft.com/en-us/office/vba/api/excel.listobject.showtablestylefirstcolumn)
    */
    bool GetShowTableStyleFirstColumn();

    /**
    Sets if the first column is formatted.

    [MSDN documentation for ListObject.ShowTableStyleFirstColumn](https://learn.microsoft.com/en-us/office/vba/api/excel.listobject.showtablestylefirstcolumn).
    */
    void SetShowTableStyleFirstColumn(bool showTableStyleFirstColumn);

    /**
    Returns if the last column is displayed.

    [MSDN documentation for ListObject.ShowTableStyleLastColumn](https://learn.microsoft.com/en-us/office/vba/api/excel.listobject.showtablestylelastcolumn)
    */
    bool GetShowTableStyleLastColumn();

    /**
    Sets if the last column is displayed.

    [MSDN documentation for ListObject.ShowTableStyleLastColumn](https://learn.microsoft.com/en-us/office/vba/api/excel.listobject.showtablestylelastcolumn).
    */
    void SetShowTableStyleLastColumn(bool showTableStyleLastColumn);

    /**
    Returns if the Row Stripes table style is used.

    [MSDN documentation for ListObject.ShowTableStyleRowStripes](https://learn.microsoft.com/en-us/office/vba/api/excel.listobject.showtablestylerowstripes)
    */
    bool GetShowTableStyleRowStripes();

    /**
    Sets if the Row Stripes table style is used.

    [MSDN documentation for ListObject.ShowTableStyleRowStripes](https://learn.microsoft.com/en-us/office/vba/api/excel.listobject.showtablestylerowstripes).
    */
    void SetShowTableStyleRowStripes(bool showTableStyleRowStripes);

    /**
    Gets or sets a Boolean to indicate whether the Total row is visible.

    [MSDN documentation for ListObject.ShowTotals](https://learn.microsoft.com/en-us/office/vba/api/excel.listobject.showtotals)
    */
    bool GetShowTotals();

    /**
    Gets or sets a Boolean to indicate whether the Total row is visible.

    [MSDN documentation for ListObject.ShowTotals](https://learn.microsoft.com/en-us/office/vba/api/excel.listobject.showtotals).
    */
    void SetShowTotals(bool showTotals);

    /**
    Gets or sets the sort column or columns, and sort order for the ListObject collection.

    [MSDN documentation for ListObject.Sort](https://learn.microsoft.com/en-us/office/vba/api/excel.listobject.sort)
    */
    wxExcelSort GetSort();

    /**
    Returns a XlListObjectSourceType value that represents the current source of the list.

    [MSDN documentation for ListObject.SourceType](https://learn.microsoft.com/en-us/office/vba/api/excel.listobject.sourcetype)
    */
    XlListObjectSourceType GetSourceType();

    /**
    Returns the description associated with the alternative text string for the specified table.

    [MSDN documentation for ListObject.Summary](https://learn.microsoft.com/en-us/office/vba/api/excel.listobject.summary)
    */
    wxString GetSummary();

    /**
    Sets the description associated with the alternative text string for the specified table.

    [MSDN documentation for ListObject.Summary](https://learn.microsoft.com/en-us/office/vba/api/excel.listobject.summary).
    */
    void SetSummary(const wxString& summary);

    /**
    Returns a TableObject object.

    [MSDN documentation for ListObject.TableObject](https://learn.microsoft.com/en-us/office/vba/api/excel.listobject.tableobject)
    */
    wxExcelTableObject GetTableObject();

    /**
    Gets the table style.

    [MSDN documentation for ListObject.TableStyle](https://learn.microsoft.com/en-us/office/vba/api/excel.listobject.tablestyle)
    */
    wxExcelTableStyle GetTableStyle();

    /**
    Sets the table style.

    [MSDN documentation for ListObject.TableStyle](https://learn.microsoft.com/en-us/office/vba/api/excel.listobject.tablestyle).
    */
    void SetTableStyle(wxExcelTableStyle tableStyle);

    /**
    Returns a Range representing the Total row, if any.

    [MSDN documentation for ListObject.TotalsRowRange](https://learn.microsoft.com/en-us/office/vba/api/excel.listobject.totalsrowrange)
    */
    wxExcelRange GetTotalsRowRange();

    /**
    Returns "ListObject".
    */
    virtual wxString GetAutoExcelObjectName_() const { return wxS("ListObject"); }

}; // class wxExcelListObject


/**
    @brief Represents a collection of ListObject objects on a worksheet,
    where each ListObject object represents a table on the worksheet.
*/
class WXDLLIMPEXP_WXAUTOEXCEL wxExcelListObjects : public wxExcelObject
{
public:
    // ***** METHODS *****


    /**
        Creates a new ListObject. Use when the source type is xlSrcRange.

        [MSDN documentation for ListObjects.Add](https://learn.microsoft.com/en-us/office/vba/api/excel.listobjects.add)
    */
    wxExcelListObject Add(wxExcelRange* source = NULL,
                          XlYesNoGuess* XlListObjectHasHeaders = NULL,
                          const wxString& tableStyleName = wxEmptyString);

    /**
        Creates a new ListObject. Use when the source type is xlSrcExternal.

        [MSDN documentation for ListObjects.Add](https://learn.microsoft.com/en-us/office/vba/api/excel.listobjects.add)
    */
    wxExcelListObject Add(const wxArrayString& source,
                          wxExcelRange destination,
                          wxXlTribool linkSource = wxDefaultXlTribool,
                          XlYesNoGuess* XlListObjectHasHeaders = NULL,
                          const wxString& tableStyleName = wxEmptyString);


    // ***** PROPERTIES *****

    /**
        Returns the number of objects in the collection.

        [MSDN documentation for ListObjects.Count](https://learn.microsoft.com/en-us/office/vba/api/excel.listobjects.count)
    */
    long GetCount();

    //@{
    /**
        Returns the ListObject with the given index or name.

        [MSDN documentation for ListObjects.Item](https://learn.microsoft.com/en-us/office/vba/api/excel.listobjects.item)
    */
    wxExcelListObject GetItem(long index);
    wxExcelListObject GetItem(const wxString& name);
    wxExcelListObject operator[](long index);
    wxExcelListObject operator[](const wxString& name);
    //@}

    /**
        Returns "ListObjects".
    */
    virtual wxString GetAutoExcelObjectName_() const { return wxS("ListObjects"); }

}; // class wxExcelListObjects

} // namespace wxAutoExcel

#endif // #ifndef _WXAUTOEXCEL_LISTOBJECT_H
