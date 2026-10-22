/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// License:     MIT license
/////////////////////////////////////////////////////////////////////////////


#ifndef _WXAUTOEXCEL_HILOLINES_H
#define _WXAUTOEXCEL_HILOLINES_H

#include "wx/wxAutoExcel_defs.h"

#if WXAUTOEXCEL_USE_CHARTS

#include "wx/wxAutoExcel_object.h"

namespace wxAutoExcel {

    /**
    @brief Represents Microsoft Excel HiLoLines object.
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelHiLoLines : public wxExcelObject
    {
    public:
        // ***** METHODS *****

        /**
        Deletes the object.

        [MSDN documentation for HiLoLines.Delete](https://learn.microsoft.com/en-us/office/vba/api/excel.hilolines.delete).
        */
        bool Delete();

        /**
        Selects the object.

        [MSDN documentation for HiLoLines.Select](https://learn.microsoft.com/en-us/office/vba/api/excel.hilolines.select).
        */
        bool Select();

        // ***** PROPERTIES *****


        /**
        Returns a Border object that represents the border of the object.

        [MSDN documentation for HiLoLines.Border](https://learn.microsoft.com/en-us/office/vba/api/excel.hilolines.border).
        */
        wxExcelBorder GetBorder();

        /**
        Read-only Since Excel 2007.

        [MSDN documentation for HiLoLines.Format](https://learn.microsoft.com/en-us/office/vba/api/excel.hilolines.format).
        */
        wxExcelChartFormat GetFormat();

        /**
        Returns a String value that represents the name of the object.

        [MSDN documentation for HiLoLines.Name](https://learn.microsoft.com/en-us/office/vba/api/excel.hilolines.name).
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
