/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2020 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////

#ifndef _WXAUTOEXCEL_DATABARBORDER_H
#define _WXAUTOEXCEL_DATABARBORDER_H

#include "wx/wxAutoExcel_defs.h"

#if WXAUTOEXCEL_USE_CONDFORMAT

#include "wx/wxAutoExcel_object.h"
#include "wx/wxAutoExcel_enums.h"

namespace wxAutoExcel
{

/**
@brief Represents the border of the data bars specified by a conditional formatting rule.
*/
class WXDLLIMPEXP_WXAUTOEXCEL wxExcelDataBarBorder : public wxExcelObject
{
public:
    // ***** PROPERTIES *****

    /**
    Returns an object that specifies the color of the border of data bars specified by a conditional formatting rule. Read-only

    [Excel VBA documentation for DataBarBorder.Color](https://docs.microsoft.com/en-us/office/vba/api/excel.databarborder.color)
    */
    wxColour GetColor();

    /**
    Returns the type of border for data bars specified by a conditional formatting rule. Read/write

    [Excel VBA documentation for DataBarBorder.Type](https://docs.microsoft.com/en-us/office/vba/api/excel.databarborder.type)
    */
    XlDataBarBorderType GetType();

    /**
    Sets the type of border for data bars specified by a conditional formatting rule. Read/write

    [Excel VBA documentation for DataBarBorder.Type](https://docs.microsoft.com/en-us/office/vba/api/excel.databarborder.type)
    */
    void SetType(XlDataBarBorderType type);

    /**
    Returns "DataBarBorder".
    */
    virtual wxString GetAutoExcelObjectName_() const { return wxS("DataBarBorder"); }

}; // class wxExcelDataBarBorder

} // namespace wxAutoExcel

#endif // WXAUTOEXCEL_USE_CONDFORMAT

#endif // #ifndef _WXAUTOEXCEL_DATABARBORDER_H
