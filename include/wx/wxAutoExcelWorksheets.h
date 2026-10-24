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

        [MSDN documentation for Worksheets.Add](https://learn.microsoft.com/en-us/office/vba/api/excel.worksheets.add).
        */
        wxExcelWorksheet Add(long count = 1, const wxString& templateFile = wxEmptyString);
        wxExcelWorksheet AddAfterOrBefore(wxExcelSheet sheetAfterOrBefore, bool after, long count = 1, const wxString& templateFile = wxEmptyString);
        //@}

        /**
        Copies a range to the same area on all other worksheets in a collection.

        [MSDN documentation for Worksheets.FillAcrossSheets](https://learn.microsoft.com/en-us/office/vba/api/excel.worksheets.fillacrosssheets).
        */
        void FillAcrossSheets(wxExcelRange range, XlFillWith* type);

        /**
        Selects the object.

        [MSDN documentation for Worksheets.Select](https://learn.microsoft.com/en-us/office/vba/api/excel.worksheets.select).
        */
        void Select(wxXlTribool replace =  wxDefaultXlTribool);

        // ***** PROPERTIES *****

        /**
        Returns a Long value that represents the number of objects in the collection.

        [MSDN documentation for Worksheets.Count](https://learn.microsoft.com/en-us/office/vba/api/excel.worksheets.count).
        */
        long GetCount();

        /**
        Returns a PageBreaks collection that represents the horizontal page breaks on the sheet.

        [MSDN documentation for Worksheets.HPageBreaks](https://learn.microsoft.com/en-us/office/vba/api/excel.worksheets.hpagebreaks).
        */
        wxExcelPageBreaks GetHPageBreaks();

        //@{
        /**
        Returns a single object from a collection.

        [MSDN documentation for Worksheets.Item](https://learn.microsoft.com/en-us/office/vba/api/excel.worksheets.item).
        */
        wxExcelWorksheet GetItem(long index);
        wxExcelWorksheet GetItem(const wxString& name);
        wxExcelWorksheet operator[](long index);
        wxExcelWorksheet operator[](const wxString& name);
        //@}

        /**
        Returns a Variant value that determines whether the object is visible.

        [MSDN documentation for Worksheets.Visible](https://learn.microsoft.com/en-us/office/vba/api/excel.worksheets.visible).
        */
        bool GetVisible();

        /**
        Returns a PageBreaks collection that represents the vertical page breaks on the sheet.

        [MSDN documentation for Worksheets.VPageBreaks](https://learn.microsoft.com/en-us/office/vba/api/excel.worksheets.vpagebreaks).
        */
        wxExcelPageBreaks  GetVPageBreaks();

        /**
        The iterator over the items, see wxExcelCollectionIterator.
        */
        typedef wxExcelCollectionIterator<wxExcelWorksheets, wxExcelWorksheet> iterator;

        //@{
        /**
        Returns the iterator referring to the first item, or past the last item,
        which allows iterating over the items with a range-based for loop,
        see wxExcelCollectionIterator.
        */
        iterator begin() { return iterator(this, 1); }
        iterator end() { return iterator(this, GetCount() + 1); }
        //@}

        /**
        Returns "Worksheets".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("Worksheets"); }
    private:
        wxExcelWorksheet DoAdd(wxExcelSheet* sheetAfterOrBefore, bool after, long count, const wxString& templateFile);
    };

} // namespace wxAutoExcel

#endif //_WXAUTOEXCEL_WORKSHEETS_H
