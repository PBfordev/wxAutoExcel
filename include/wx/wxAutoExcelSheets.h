/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// License:     MIT license
/////////////////////////////////////////////////////////////////////////////


#ifndef _WXAUTOEXCEL_SHEETS_H
#define _WXAUTOEXCEL_SHEETS_H

#include "wx/wxAutoExcel_defs.h"

#include "wx/wxAutoExcel_object.h"

namespace wxAutoExcel {

/**
    @brief Represents Microsoft Excel Sheets collection.
*/
class WXDLLIMPEXP_WXAUTOEXCEL wxExcelSheets : public wxExcelObject
{
public:
    //@{
    /**
        Creates a new sheet.
    */
    wxExcelSheet Add(long count = 1, XlSheetType type = xlWorksheet);
    wxExcelSheet AddAfterOrBefore(wxExcelSheet sheetAfterOrBefore, bool after, long count = 1, XlSheetType type = xlWorksheet);
    //@}

    /**
        Returns the number of sheets in the collection.
    */
    long GetCount();

    //@{
    /**
        Returns the sheet with the given index or name.
    */
    wxExcelSheet GetItem(long index);
    wxExcelSheet GetItem(const wxString& name);
    wxExcelSheet operator[](long index);
    wxExcelSheet operator[](const wxString& name);
    //@}

    /**
    Returns "Sheets".
    */
    virtual wxString GetAutoExcelObjectName_() const { return wxS("Sheets"); }
private:
    wxExcelSheet DoAdd(wxExcelSheet* sheetAfterOrBefore, bool after, long count, XlSheetType type);
};

 } // namespace wxAutoExcel

#endif //_WXAUTOEXCEL_SHEETS_H
