/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Modified by:
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////


#ifndef _WXAUTOEXCEL_DROPLINES_H
#define _WXAUTOEXCEL_DROPLINES_H

#include "wx/wxAutoExcel_defs.h"

#if WXAUTOEXCEL_USE_CHARTS

#include "wx/wxAutoExcel_object.h"

namespace wxAutoExcel {

    /**
    Represents Microsoft Excel DropLines object.
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelDropLines : public wxExcelObject
    {
    public:
        // ***** METHODS *****

        /**
        Deletes the object.

        [MSDN documentation for DropLines.Delete](http://msdn.microsoft.com/en-us/library/bb211750).
        */
        bool Delete();

        /**
        Selects the object.

        [MSDN documentation for DropLines.Select](http://msdn.microsoft.com/en-us/library/bb237880).
        */
        bool Select();

        // ***** PROPERTIES *****

        /**
        Returns a Border object that represents the border of the object.

        [MSDN documentation for DropLines.Border](http://msdn.microsoft.com/en-us/library/bb236958).
        */
        wxExcelBorder GetBorder();

        /**
        Returns the ChartFormat object. Since Excel 2007.

        [MSDN documentation for DropLines.Format](http://msdn.microsoft.com/en-us/library/bb242527).
        */
        wxExcelChartFormat GetFormat();

        /**
        Returns a String value that represents the name of the object.

        [MSDN documentation for DropLines.Name](http://msdn.microsoft.com/en-us/library/bb236960).
        */
        wxString GetName();

        /**
        Returns "DropLines".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("DropLines"); }
    };


} // namespace wxAutoExcel

#endif // #if WXAUTOEXCEL_USE_CHARTS

#endif // _WXAUTOEXCEL_DROPLINES_H
