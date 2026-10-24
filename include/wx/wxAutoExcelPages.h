/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// License:     MIT license
/////////////////////////////////////////////////////////////////////////////


#ifndef _WXAUTOEXCEL_PAGES_H
#define _WXAUTOEXCEL_PAGES_H

#include "wx/wxAutoExcel_defs.h"

#include "wx/wxAutoExcel_object.h"

namespace wxAutoExcel {

    /**
    @brief Represents Microsoft Excel Page object.
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelPage : public wxExcelObject
    {
    public:
        // ***** PROPERTIES *****

        /**
        Returns the footer. Since Excel 2007.

        [MSDN documentation for Page.CenterFooter](https://learn.microsoft.com/en-us/office/vba/api/excel.page.centerfooter).
        */
        wxExcelHeaderFooter GetCenterFooter();

        /**
        Returns the header. Since Excel 2007.

        [MSDN documentation for Page.CenterHeader](https://learn.microsoft.com/en-us/office/vba/api/excel.page.centerheader).
        */
        wxExcelHeaderFooter GetCenterHeader();

        /**
        Returns the footer. Since Excel 2007.

        [MSDN documentation for Page.LeftFooter](https://learn.microsoft.com/en-us/office/vba/api/excel.page.leftfooter).
        */
        wxExcelHeaderFooter GetLeftFooter();

        /**
        Returns the header. Since Excel 2007.

        [MSDN documentation for Page.LeftHeader](https://learn.microsoft.com/en-us/office/vba/api/excel.page.leftheader).
        */
        wxExcelHeaderFooter GetLeftHeader();

        /**
        Returns the footer. Since Excel 2007.

        [MSDN documentation for Page.RightFooter](https://learn.microsoft.com/en-us/office/vba/api/excel.page.rightfooter).
        */
        wxExcelHeaderFooter GetRightFooter();

        /**
        Returns the header. Since Excel 2007.

        [MSDN documentation for Page.RightHeader](https://learn.microsoft.com/en-us/office/vba/api/excel.page.rightheader).
        */
        wxExcelHeaderFooter GetRightHeader();

        /**
        Returns "Page".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("Page"); }
    };

    /**
    @brief Represents Microsoft Excel Pages collection.
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelPages : public wxExcelObject
    {
    public:
        // ***** PROPERTIES *****

        /**
        Number of pages in the collection. Since Excel 2007.

        [MSDN documentation for Pages.Count](https://learn.microsoft.com/en-us/office/vba/api/excel.pages.count).
        */
        long GetCount();

        //@{
        /**
        Returns the page. Since Excel 2007.

        [MSDN documentation for Pages.Item](https://learn.microsoft.com/en-us/office/vba/api/excel.pages.item).
        */
        wxExcelPage GetItem(long index);
        wxExcelPage operator[](long index);
        //@}

        /**
        The iterator over the items, see wxExcelCollectionIterator.
        */
        typedef wxExcelCollectionIterator<wxExcelPages, wxExcelPage> iterator;

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
        Returns "Pages".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("Pages"); }
    };


} // namespace wxAutoExcel

#endif //_WXAUTOEXCEL_PAGES_H
