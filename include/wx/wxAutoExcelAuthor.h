/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Modified by:
// Copyright:   (c) 2020 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////

#ifndef _WXAUTOEXCEL_AUTHOR_H
#define _WXAUTOEXCEL_AUTHOR_H

#include "wx/wxAutoExcel_defs.h"

#include "wx/wxAutoExcel_object.h"

namespace wxAutoExcel
{

/**
@brief Represents the author of the CommentThreaded object.
*/
class WXDLLIMPEXP_WXAUTOEXCEL wxExcelAuthor : public wxExcelObject
{
public:
    // ***** PROPERTIES *****

    /**
    Returns a String that contains the display name of the specified comment author.

    [Excel VBA documentation for Author.Name](https://docs.microsoft.com/en-us/office/vba/api/excel.author.name)
    */
    wxString GetName();

    /**
    Returns a String that represents the ID of the service providing the contact information.

    [Excel VBA documentation for Author.ProviderID](https://docs.microsoft.com/en-us/office/vba/api/excel.author.providerid)
    */
    wxString GetProviderID();

    /**
    Returns a String that represents the user ID of the contact.

    [Excel VBA documentation for Author.UserID](https://docs.microsoft.com/en-us/office/vba/api/excel.author.userid)
    */
    wxString GetUserID();

    /**
    Returns "Author".
    */
    virtual wxString GetAutoExcelObjectName_() const { return wxS("Author"); }

}; // class wxExcelAuthor

} // namespace wxAutoExcel 

#endif // #ifndef _WXAUTOEXCEL_AUTHOR_H
