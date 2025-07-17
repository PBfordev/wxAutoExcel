/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2020 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////

#ifndef _WXAUTOEXCEL_USERACCESS_H
#define _WXAUTOEXCEL_USERACCESS_H

#include "wx/wxAutoExcel_defs.h"

#include "wx/wxAutoExcel_object.h"

namespace wxAutoExcel
{

/**
@brief Represents Represents the user access for a protected range.
*/
class WXDLLIMPEXP_WXAUTOEXCEL wxExcelUserAccess : public wxExcelObject
{
public:
    // ***** METHODS *****

    /**
    Deletes the object.

    [Excel VBA documentation for UserAccess.Delete](https://docs.microsoft.com/en-us/office/vba/api/excel.useraccess.delete)
    */
    void Delete();

    // ***** PROPERTIES *****

    /**
    Returns a Boolean value that indicates if the user is allowed access to the specified range on a protected worksheet.

    [Excel VBA documentation for UserAccess.AllowEdit](https://docs.microsoft.com/en-us/office/vba/api/excel.useraccess.allowedit)
    */
    bool GetAllowEdit();

    /**
    Sets a Boolean value that indicates if the user is allowed access to the specified range on a protected worksheet.

    [Excel VBA documentation for UserAccess.AllowEdit](https://docs.microsoft.com/en-us/office/vba/api/excel.useraccess.allowedit)
    */
    void SetAllowEdit(bool allowEdit);

    /**
    Returns a String value that represents the name of the object.

    [Excel VBA documentation for UserAccess.Name](https://docs.microsoft.com/en-us/office/vba/api/excel.useraccess.name)
    */
    wxString GetName();

    /**
    Sets a String value that represents the name of the object.

    [Excel VBA documentation for UserAccess.Name](https://docs.microsoft.com/en-us/office/vba/api/excel.useraccess.name)
    */
    void SetName(const wxString& name);

    /**
    Returns "UserAccess".
    */
    virtual wxString GetAutoExcelObjectName_() const { return wxS("UserAccess"); }

}; // class wxExcelUserAccess


/**
    @brief Represents a collection of all the UserAccess objects that represent the cells that can be edited on a protected worksheet.
*/
class WXDLLIMPEXP_WXAUTOEXCEL wxExcelUserAccessList : public wxExcelObject
{
public:
    // ***** METHODS *****


    /**
        Adds a user access list.

        [MSDN documentation for UserAccessLists.Add](https://docs.microsoft.com/en-us/office/vba/api/excel.useraccesslist.add)
    */
    wxExcelUserAccess Add(const wxString& name, bool allowEdit);

     /**
        AddRemoves all users who have access to a protected range on a worksheet.

        [MSDN documentation for UserAccessLists.DeleteAll](https://docs.microsoft.com/en-us/office/vba/api/excel.useraccesslist.deleteall)
    */
    wxExcelUserAccess DeleteAll();

    // ***** PROPERTIES *****

    /**
        Returns the number of objects in the collection.

        [MSDN documentation for UserAccessLists.Count](https://docs.microsoft.com/en-us/office/vba/api/excel.useraccesslist.count)
    */
    long GetCount();

    //@{
    /**
        Returns the UserAccess with the given index or name.

        [MSDN documentation for UserAccessLists.Item](https://docs.microsoft.com/en-us/office/vba/api/excel.useraccesslist.item)
    */
    wxExcelUserAccess GetItem(long index);
    wxExcelUserAccess GetItem(const wxString& name);
    wxExcelUserAccess operator[](long index);
    wxExcelUserAccess operator[](const wxString& name);
    //@}

    /**
        Returns "UserAccessList".
    */
    virtual wxString GetAutoExcelObjectName_() const { return wxS("UserAccessList"); }

}; // class wxExcelUserAccessList

} // namespace wxAutoExcel

#endif // #ifndef _WXAUTOEXCEL_USERACCESS_H
