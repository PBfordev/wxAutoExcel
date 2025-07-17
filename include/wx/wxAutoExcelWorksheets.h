/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// License:     MIT license
/////////////////////////////////////////////////////////////////////////////


#ifndef _WXAUTOEXCEL_WORKSHEETS_H
#define _WXAUTOEXCEL_WORKSHEETS_H

#include "wx/wxAutoExcel_defs.h"

#include "wx/wxAutoExcel_object.h"
#include "wx/wxAutoExcel_enums.h"

namespace wxAutoExcel {

    /**
    @brief Represents Microsoft Excel Worksheets collection.
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelWorksheets : public wxExcelObject
    {
    public:
        // ***** METHODS *****

        //@{
        /**
        Creates a new worksheet. The new worksheet becomes the active sheet.

        [MSDN documentation for Worksheets.Add](http://msdn.microsoft.com/en-us/library/bb179213.aspx).
        */
        wxExcelWorksheet Add(long count = 1, const wxString& templateFile = wxEmptyString);
        wxExcelWorksheet AddAfterOrBefore(wxExcelSheet sheetAfterOrBefore, bool after, long count = 1, const wxString& templateFile = wxEmptyString);
        //@}

        /**
        Copies a range to the same area on all other worksheets in a collection.

        [MSDN documentation for Worksheets.FillAcrossSheets](http://msdn.microsoft.com/en-us/library/bb212427.aspx).
        */
        void FillAcrossSheets(wxExcelRange range, XlFillWith* type);

        /**
        Selects the object.

        [MSDN documentation for Worksheets.Select](http://msdn.microsoft.com/en-us/library/bb214169.aspx).
        */
        void Select(wxXlTribool replace =  wxDefaultXlTribool);

        // ***** PROPERTIES *****

        /**
        Returns a Long value that represents the number of objects in the collection.

        [MSDN documentation for Worksheets.Count](http://msdn.microsoft.com/en-us/library/bb238434.aspx).
        */
        long GetCount();

        /**
        Returns a PageBreaks collection that represents the horizontal page breaks on the sheet.

        [MSDN documentation for Worksheets.HPageBreaks](http://msdn.microsoft.com/en-us/library/bb238448.aspx).
        */
        wxExcelPageBreaks GetHPageBreaks();

        //@{
        /**
        Returns a single object from a collection.

        [MSDN documentation for Worksheets.Item](http://msdn.microsoft.com/en-us/library/bb238455.aspx).
        */
        wxExcelWorksheet GetItem(long index);
        wxExcelWorksheet GetItem(const wxString& name);
        wxExcelWorksheet operator[](long index);
        wxExcelWorksheet operator[](const wxString& name);
        //@}

        /**
        Returns a Variant value that determines whether the object is visible.

        [MSDN documentation for Worksheets.Visible](http://msdn.microsoft.com/en-us/library/bb215265.aspx).
        */
        bool GetVisible();

        /**
        Returns a PageBreaks collection that represents the vertical page breaks on the sheet.

        [MSDN documentation for Worksheets.VPageBreaks](http://msdn.microsoft.com/en-us/library/bb224511.aspx).
        */
        wxExcelPageBreaks  GetVPageBreaks();

        /**
        Returns "Worksheets".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("Worksheets"); }
    private:
        wxExcelWorksheet DoAdd(wxExcelSheet* sheetAfterOrBefore, bool after, long count, const wxString& templateFile);
    };

} // namespace wxAutoExcel

#endif //_WXAUTOEXCEL_WORKSHEETS_H
