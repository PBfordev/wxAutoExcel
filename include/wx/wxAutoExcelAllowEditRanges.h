/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2020 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////

#ifndef _WXAUTOEXCEL_ALLOWEDITRANGES_H
#define _WXAUTOEXCEL_ALLOWEDITRANGES_H

#include "wx/wxAutoExcel_defs.h"

#include "wx/wxAutoExcel_object.h"

namespace wxAutoExcel
{

/**
@brief Represents a collection of all the AllowEditRange objects that represent the cells that can be edited on a protected worksheet.
*/
class WXDLLIMPEXP_WXAUTOEXCEL wxExcelAllowEditRange : public wxExcelObject
{
public:
    // ***** METHODS *****

    /**
    Changes the password for a range that can be edited on a protected worksheet.

    [Excel VBA documentation for AllowEditRange.ChangePassword](https://docs.microsoft.com/en-us/office/vba/api/excel.alloweditrange.changepassword)
    */
    void ChangePassword(const wxString& password);

    /**
    Deletes the object.

    [Excel VBA documentation for AllowEditRange.Delete](https://docs.microsoft.com/en-us/office/vba/api/excel.alloweditrange.delete)
    */
    void Delete();

    /**
    Removes protection from a sheet or workbook. This method has no effect if the sheet or workbook isn't protected.

    [Excel VBA documentation for AllowEditRange.Unprotect](https://docs.microsoft.com/en-us/office/vba/api/excel.alloweditrange.unprotect)
    */
    void Unprotect(const wxString& password);

    // ***** PROPERTIES *****

    /**
    Returns a Range object that represents a subset of the ranges that can be edited edited on a protected worksheet.

    [Excel VBA documentation for AllowEditRange.Range](https://docs.microsoft.com/en-us/office/vba/api/excel.alloweditrange.range)
    */
    wxExcelRange GetRange();

    /**
    Returns the title of the range of cells that can edited on a protected sheet.

    [Excel VBA documentation for AllowEditRange.Title](https://docs.microsoft.com/en-us/office/vba/api/excel.alloweditrange.title)
    */
    wxString GetTitle();

    /**
    Sets the title of the range of cells that can edited on a protected sheet.

    [Excel VBA documentation for AllowEditRange.Title](https://docs.microsoft.com/en-us/office/vba/api/excel.alloweditrange.title)
    */
    void SetTitle(const wxString& title);

    /**
    Returns a UserAccessList object for the protected range on a worksheet.

    [Excel VBA documentation for AllowEditRange.Users](https://docs.microsoft.com/en-us/office/vba/api/excel.alloweditrange.users)
    */
    wxExcelUserAccessList GetUsers();

    /**
    Returns "AllowEditRange".
    */
    virtual wxString GetAutoExcelObjectName_() const { return wxS("AllowEditRange"); }

}; // class wxExcelAllowEditRange


/**
    @brief Represents a collection of all the AllowEditRange objects that represent the cells that can be edited on a protected worksheet.
*/
class WXDLLIMPEXP_WXAUTOEXCEL wxExcelAllowEditRanges : public wxExcelObject
{
public:
    // ***** METHODS *****


    /**
        Adds a range that can be edited on a protected worksheet.

        [MSDN documentation for AllowEditRanges.Add](https://docs.microsoft.com/en-us/office/vba/api/excel.alloweditranges.add)
    */
    wxExcelAllowEditRange Add(const wxString& title, wxExcelRange& range,
                              const wxString& password = wxEmptyString);

    // ***** PROPERTIES *****

    /**
        Returns the number of objects in the collection.

        [MSDN documentation for AllowEditRanges.Count](https://docs.microsoft.com/en-us/office/vba/api/excel.alloweditranges.count)
    */
    long GetCount();

    //@{
    /**
        Returns the AllowEditRange with the given index or name.

        [MSDN documentation for AllowEditRanges.Item](https://docs.microsoft.com/en-us/office/vba/api/excel.alloweditranges.item)
    */
    wxExcelAllowEditRange GetItem(long index);
    wxExcelAllowEditRange GetItem(const wxString& name);
    wxExcelAllowEditRange operator[](long index);
    wxExcelAllowEditRange operator[](const wxString& name);
    //@}

    /**
        Returns "AllowEditRanges".
    */
    virtual wxString GetAutoExcelObjectName_() const { return wxS("AllowEditRanges"); }

}; // class wxExcelAllowEditRanges

} // namespace wxAutoExcel

#endif // #ifndef _WXAUTOEXCEL_ALLOWEDITRANGES_H
