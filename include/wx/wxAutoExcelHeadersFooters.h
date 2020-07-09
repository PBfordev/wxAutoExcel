/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Modified by:
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////


#ifndef _WXAUTOEXCEL_HEADERSFOOTERS_H
#define _WXAUTOEXCEL_HEADERSFOOTERS_H

#include "wx/wxAutoExcel_defs.h"
#include "wx/wxAutoExcel_object.h"

namespace wxAutoExcel {


    /**
    @brief Represents Microsoft Excel HeaderFooter.
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelHeaderFooter: public wxExcelObject
    {
    public:
        // ***** PROPERTIES *****

        /**
        Returns text included in the specified header or footer. Since Excel 2007.

        [MSDN documentation for HeaderFooter.Text](http://msdn.microsoft.com/en-us/library/bb224469).
        */
        wxString GetText();

        /**
        Sets text included in the specified header or footer.

        [MSDN documentation for HeaderFooter.Text](http://msdn.microsoft.com/en-us/library/bb224469).
        */
        void SetText(const wxString& text);

        /**
        Returns "HeaderFooter".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("HeaderFooter"); }
    };

} // namespace wxAutoExcel

#endif //_WXAUTOEXCEL_HEADERSFOOTERS_H
