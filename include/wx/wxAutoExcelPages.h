/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
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
        Returns the string footer. Since Excel 2007.

        [MSDN documentation for Page.CenterFooter](http://msdn.microsoft.com/en-us/library/bb147924).
        */
        wxString GetCenterFooter();

        /**
        Returns the string header. Since Excel 2007.

        [MSDN documentation for Page.CenterHeader](http://msdn.microsoft.com/en-us/library/bb147928).
        */
        wxString GetCenterHeader();

        /**
        Returns the string footer. Since Excel 2007.

        [MSDN documentation for Page.LeftFooter](http://msdn.microsoft.com/en-us/library/bb147931).
        */
        wxString GetLeftFooter();

        /**
        Returns the string header. Since Excel 2007.

        [MSDN documentation for Page.LeftHeader](http://msdn.microsoft.com/en-us/library/bb147933).
        */
        wxString GetLeftHeader();

        /**
        Returns the string footer. Since Excel 2007.

        [MSDN documentation for Page.RightFooter](http://msdn.microsoft.com/en-us/library/bb147937).
        */
        wxString GetRightFooter();

        /**
        Returns the string header. Since Excel 2007.

        [MSDN documentation for Page.RightHeader](http://msdn.microsoft.com/en-us/library/bb147939).
        */
        wxString GetRightHeader();

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

        [MSDN documentation for Pages.Count](http://msdn.microsoft.com/en-us/library/bb147942).
        */
        long GetCount();

        //@{
        /**
        Returns the page. Since Excel 2007.

        [MSDN documentation for Pages.Item](http://msdn.microsoft.com/en-us/library/bb147947).
        */
        wxExcelPage GetItem(long index);
        wxExcelPage operator[](long index);
        //@}

        /**
        Returns "Pages".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("Pages"); }
    };


} // namespace wxAutoExcel

#endif //_WXAUTOEXCEL_PAGES_H
