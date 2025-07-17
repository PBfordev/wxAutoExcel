/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// License:     MIT license
/////////////////////////////////////////////////////////////////////////////


#ifndef _WXAUTOEXCEL_ICON_H
#define _WXAUTOEXCEL_ICON_H

#include "wx/wxAutoExcel_defs.h"
#include "wx/wxAutoExcel_object.h"

namespace wxAutoExcel {


    /**
    @brief Represents Microsoft Excel Icon object.
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelIcon: public wxExcelObject
    {
    public:
        /**
        Returns a Long value specifying the index number of the Icon object within the IconSet object. Since Excel 2007.

        [MSDN documentation for Icon.Index](http://msdn.microsoft.com/en-us/library/bb224565).
        */
        long GetIndex();
        /**
        Returns "Icon".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("Icon"); }
    };


} // namespace wxAutoExcel

#endif //_WXAUTOEXCEL_ICON_H
