/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Modified by:
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////


#ifndef _WXAUTOEXCEL_HILOLINES_H
#define _WXAUTOEXCEL_HILOLINES_H

#include "wx/wxAutoExcel_defs.h"

#if WXAUTOEXCEL_USE_CHARTS

#include "wx/wxAutoExcel_object.h"

namespace wxAutoExcel {

    /**
    Represents Microsoft Excel HiLoLines object.
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelHiLoLines : public wxExcelObject
    {
    public:
        // ***** METHODS *****

        /**
        Deletes the object.

        [MSDN documentation for HiLoLines.Delete](http://msdn.microsoft.com/en-us/library/bb211794).
        */
        bool Delete();

        /**
        Selects the object.

        [MSDN documentation for HiLoLines.Select](http://msdn.microsoft.com/en-us/library/bb237953).
        */
        bool Select();

        // ***** PROPERTIES *****
        

        /**
        Returns a Border object that represents the border of the object.

        [MSDN documentation for HiLoLines.Border](http://msdn.microsoft.com/en-us/library/bb148507).
        */
        wxExcelBorder GetBorder();

        /**
        Read-only Since Excel 2007.

        [MSDN documentation for HiLoLines.Format](http://msdn.microsoft.com/en-us/library/bb242531).
        */
        wxExcelChartFormat GetFormat();

        /**
        Returns a String value that represents the name of the object.

        [MSDN documentation for HiLoLines.Name](http://msdn.microsoft.com/en-us/library/bb148509).
        */
        wxString GetName();

        /**
        Returns "HiLoLines".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("HiLoLines"); }
    };


} // namespace wxAutoExcel

#endif // #if WXAUTOEXCEL_USE_CHARTS

#endif // _WXAUTOEXCEL_HILOLINES_H
