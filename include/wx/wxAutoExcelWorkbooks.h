/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////


#ifndef _WXAUTOEXCEL_WORKBOOKS_H
#define _WXAUTOEXCEL_WORKBOOKS_H

#include "wx/wxAutoExcel_defs.h"

#include "wx/wxAutoExcel_object.h"
#include "wx/wxAutoExcel_enums.h"

namespace wxAutoExcel {

    /**
    @brief Represents Microsoft Excel Workbooks collection.
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelWorkbooks : public wxExcelObject
    {
    public:
        // ***** METHODS *****

        //@{
        /**
        Creates a new workbook. The new workbook becomes the active workbook.

        [MSDN documentation for Workbooks.Add](http://msdn.microsoft.com/en-us/library/bb179164.aspx).
        */
        wxExcelWorkbook Add(XlWBATemplate sheetType = xlWBATWorksheet);
        wxExcelWorkbook Add(const wxString& templateFile);
        //@}

        /**
        True if Microsoft Excel can check out a specified workbook from a server. Read/write Boolean.

        [MSDN documentation for Workbooks.CanCheckOut](http://msdn.microsoft.com/en-us/library/bb223219.aspx).
        */
        bool CanCheckOut(const wxString& fileName);

        /**
        Checks out a specified workbook from a server to a local computer for editing.

        [MSDN documentation for Workbooks.CheckOut](http://msdn.microsoft.com/en-us/library/bb223249.aspx).
        */
        void CheckOut(const wxString& fileName);

        /**
        Closes the workbook.

        [MSDN documentation for Workbooks.Close](http://msdn.microsoft.com/en-us/library/bb179166.aspx).
        */
        bool Close();

        //@{
        /**
        Opens a workbook.

        [MSDN documentation for Workbooks.Open](http://msdn.microsoft.com/en-us/library/bb179167.aspx).
        */
        wxExcelWorkbook Open(const wxString& fileName, long* updateLinks = NULL, wxXlTribool readOnly = wxDefaultXlTribool,
                              long* format = NULL,
                              const wxString& password = wxEmptyString,
                              const wxString& writeResPassword = wxEmptyString,
                              wxXlTribool ignoreReadOnlyRecommended = wxDefaultXlTribool, XlPlatform* origin = NULL,
                              wxXlTribool editable = wxDefaultXlTribool, wxXlTribool notify = wxDefaultXlTribool, long* converter = NULL,
                              wxXlTribool addToMru = wxDefaultXlTribool, wxXlTribool local = wxDefaultXlTribool, XlCorruptLoad* corruptLoad = NULL);

        wxExcelWorkbook Open(const wxString& fileName, const wxVariantVector& optionalArgs);
        //@}


        /**
        Returns a workbook representing a database.

        [MSDN documentation for Workbooks.OpenDatabase](http://msdn.microsoft.com/en-us/library/bb223508.aspx).
        */
        wxExcelWorkbook OpenDatabase(const wxString& fileName, const wxString& commandText = wxEmptyString,
                                  XlCmdType commandType = xlCmdDefault, bool backgroundQuery = false,
                                  XlImportDataAs importDataAs = xlQueryTable);

        // ***** PROPERTIES *****

        /**
        Returns the number of workbooks in the collection.

        [MSDN documentation for Workbooks.Count](http://msdn.microsoft.com/en-us/library/bb148825.aspx).
        */
        long GetCount();
        //@{
        /**
        Returns a workbook.

        [MSDN documentation for Workbooks.Item](http://msdn.microsoft.com/en-us/library/bb148827.aspx).
        */
        wxExcelWorkbook GetItem(long index);
        wxExcelWorkbook GetItem(const wxString& name);
        wxExcelWorkbook operator[](long index);
        wxExcelWorkbook operator[](const wxString& name);

        /**
        Returns "Workbooks".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("Workbooks"); }
    private:
        wxExcelWorkbook DoAdd(const wxVariant& templateFile);
    };

} // namespace wxAutoExcel

#endif //_WXAUTOEXCEL_WORKBOOKS_H

