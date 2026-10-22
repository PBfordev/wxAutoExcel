/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// License:     MIT license
/////////////////////////////////////////////////////////////////////////////


#ifndef _WXAUTOEXCEL_DROPLINES_H
#define _WXAUTOEXCEL_DROPLINES_H

#include "wx/wxAutoExcel_defs.h"

#if WXAUTOEXCEL_USE_CHARTS

#include "wx/wxAutoExcel_object.h"

namespace wxAutoExcel {

    /**
    @brief Represents Microsoft Excel DropLines object.
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelDropLines : public wxExcelObject
    {
    public:
        // ***** METHODS *****

        /**
        Deletes the object.

        [MSDN documentation for DropLines.Delete](https://learn.microsoft.com/en-us/office/vba/api/excel.droplines.delete).
        */
        bool Delete();

        /**
        Selects the object.

        [MSDN documentation for DropLines.Select](https://learn.microsoft.com/en-us/office/vba/api/excel.droplines.select).
        */
        bool Select();

        // ***** PROPERTIES *****

        /**
        Returns a Border object that represents the border of the object.

        [MSDN documentation for DropLines.Border](https://learn.microsoft.com/en-us/office/vba/api/excel.droplines.border).
        */
        wxExcelBorder GetBorder();

        /**
        Returns the ChartFormat object. Since Excel 2007.

        [MSDN documentation for DropLines.Format](https://learn.microsoft.com/en-us/office/vba/api/excel.droplines.format).
        */
        wxExcelChartFormat GetFormat();

        /**
        Returns a String value that represents the name of the object.

        [MSDN documentation for DropLines.Name](https://learn.microsoft.com/en-us/office/vba/api/excel.droplines.name).
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
