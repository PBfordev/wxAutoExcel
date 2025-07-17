/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2020 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////

#ifndef _WXAUTOEXCEL_TABLESTYLE_H
#define _WXAUTOEXCEL_TABLESTYLE_H

#include "wx/wxAutoExcel_defs.h"

#include "wx/wxAutoExcel_object.h"

namespace wxAutoExcel
{

/**
@brief Represents a single style that can be applied to a table or slicer.
*/
class WXDLLIMPEXP_WXAUTOEXCEL wxExcelTableStyle : public wxExcelObject
{
public:
    // ***** METHODS *****

    /**
    Deletes the TableStyle object.

    [MSDN documentation for TableStyle.Delete](https://docs.microsoft.com/en-us/office/vba/api/excel.tablestyle.delete)
    */
    void Delete();

    /**
    Duplicates the TableStyle object and returns the copy.

    [MSDN documentation for TableStyle.Duplicate](https://docs.microsoft.com/en-us/office/vba/api/excel.tablestyle.duplicate)
    */
    wxExcelTableStyle Duplicate(const wxString& newTableStyleName = wxEmptyString);

    // ***** PROPERTIES *****

    /**
    True if the style is a built-in style.

    [MSDN documentation for TableStyle.BuiltIn](https://docs.microsoft.com/en-us/office/vba/api/excel.tablestyle.builtin)
    */
    bool GetBuiltIn();

    /**
    Returns the name of the object.

    [MSDN documentation for TableStyle.Name](https://docs.microsoft.com/en-us/office/vba/api/excel.tablestyle.name)
    */
    wxString GetName();

    /**
    Returns the name of the object, in the language of the user.

    [MSDN documentation for TableStyle.NameLocal](https://docs.microsoft.com/en-us/office/vba/api/excel.tablestyle.namelocal)
    */
    wxString GetNameLocal();

    /**
    Returns if a style is shown in the gallery for PivotTable styles or not.

    [MSDN documentation for TableStyle.ShowAsAvailablePivotTableStyle](https://docs.microsoft.com/en-us/office/vba/api/excel.tablestyle.showasavailablepivottablestyle)
    */
    bool GetShowAsAvailablePivotTableStyle();

    /**
    Sets if a style is shown in the gallery for PivotTable styles or not.

    [MSDN documentation for TableStyle.ShowAsAvailablePivotTableStyle](https://docs.microsoft.com/en-us/office/vba/api/excel.tablestyle.showasavailablepivottablestyle).
    */
    void SetShowAsAvailablePivotTableStyle(bool showAsAvailablePivotTableStyle);

    /**
    Returns if the specified table style is shown as available in the slicer styles gallery.

    [MSDN documentation for TableStyle.ShowAsAvailableSlicerStyle](https://docs.microsoft.com/en-us/office/vba/api/excel.tablestyle.showasavailableslicerstyle)
    */
    bool GetShowAsAvailableSlicerStyle();

    /**
    Sets if the specified table style is shown as available in the slicer styles gallery.

    [MSDN documentation for TableStyle.ShowAsAvailableSlicerStyle](https://docs.microsoft.com/en-us/office/vba/api/excel.tablestyle.showasavailableslicerstyle).
    */
    void SetShowAsAvailableSlicerStyle(bool showAsAvailableSlicerStyle);

    /**
    Returns a table style shown as available in the table styles gallery.

    [MSDN documentation for TableStyle.ShowAsAvailableTableStyle](https://docs.microsoft.com/en-us/office/vba/api/excel.tablestyle.showasavailabletablestyle)
    */
    bool GetShowAsAvailableTableStyle();

    /**
    Sets a table style shown as available in the table styles gallery.

    [MSDN documentation for TableStyle.ShowAsAvailableTableStyle](https://docs.microsoft.com/en-us/office/vba/api/excel.tablestyle.showasavailabletablestyle).
    */
    void SetShowAsAvailableTableStyle(bool showAsAvailableTableStyle);

    /**
    Returns whether the specified table style is shown as available in the timeline styles gallery.

    [MSDN documentation for TableStyle.ShowAsAvailableTimelineStyle](https://docs.microsoft.com/en-us/office/vba/api/excel.tablestyle.showasavailabletimelinestyle)
    */
    bool GetShowAsAvailableTimelineStyle();

    /**
    Sets whether the specified table style is shown as available in the timeline styles gallery.

    [MSDN documentation for TableStyle.ShowAsAvailableTimelineStyle](https://docs.microsoft.com/en-us/office/vba/api/excel.tablestyle.showasavailabletimelinestyle).
    */
    void SetShowAsAvailableTimelineStyle(bool showAsAvailableTimelineStyle);

    /**
    Returns the TableStyleElements object.

    [MSDN documentation for TableStyle.TableStyleElements](https://docs.microsoft.com/en-us/office/vba/api/excel.tablestyle.tablestyleelements)
    */
    wxExcelTableStyleElements GetTableStyleElements();

    /**
    Returns "TableStyle".
    */
    virtual wxString GetAutoExcelObjectName_() const { return wxS("TableStyle"); }

}; // class wxExcelTableStyle

/**
    @brief Represents styles that can be applied to a table.
*/
class WXDLLIMPEXP_WXAUTOEXCEL wxExcelTableStyles : public wxExcelObject
{
public:
    // ***** METHODS *****

    /**
        Adds a new style to the table.

        [MSDN documentation for TableStyles.Add](https://docs.microsoft.com/en-us/office/vba/api/excel.tablestyles.add)
    */
    wxExcelTableStyle Add(const wxString& tableStyleName);

    //@{
    /**
        Returns the TableStyle with the given index or name.

        [MSDN documentation for TableStyles.Item](https://docs.microsoft.com/en-us/office/vba/api/excel.tablestyles.item)
    */
    wxExcelTableStyle GetItem(long index);
    wxExcelTableStyle GetItem(const wxString& name);
    wxExcelTableStyle operator[](long index);
    wxExcelTableStyle operator[](const wxString& name);
    //@}


    // ***** PROPERTIES *****

    /**
        Returns the number of objects in the collection.

        [MSDN documentation for TableStyles.Count](https://docs.microsoft.com/en-us/office/vba/api/excel.tablestyles.count)
    */
    long GetCount();

    /**
    Returns "TableStyles".
    */
    virtual wxString GetAutoExcelObjectName_() const { return wxS("TableStyles"); }

}; // class wxExcelTableStyles

} // namespace wxAutoExcel

#endif // #ifndef _WXAUTOEXCEL_TABLESTYLE_H
