/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// License:     MIT license
/////////////////////////////////////////////////////////////////////////////


#ifndef _WXAUTOEXCEL_GRIDLINES_H
#define _WXAUTOEXCEL_GRIDLINES_H

#include "wx/wxAutoExcel_defs.h"

#if WXAUTOEXCEL_USE_CHARTS

#include "wx/wxAutoExcel_object.h"

namespace wxAutoExcel {


    /**
    @brief Represents Microsoft Excel Gridlines object.
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelGridlines : public wxExcelObject
    {
    public:
        // ***** METHODS *****

        /**
        Deletes the object.

        [MSDN documentation for Gridlines.Delete](https://learn.microsoft.com/en-us/office/vba/api/excel.gridlines.delete).
        */
        bool Delete();

        /**
        Selects the object.

        [MSDN documentation for Gridlines.Select](https://learn.microsoft.com/en-us/office/vba/api/excel.gridlines.select).
        */
        bool Select();

        // ***** PROPERTIES *****


        /**
        Returns a Border object that represents the border of the object.

        [MSDN documentation for Gridlines.Border](https://learn.microsoft.com/en-us/office/vba/api/excel.gridlines.border).
        */
        wxExcelBorder GetBorder();

        /**
        Returns the ChartFormat object.

        [MSDN documentation for Gridlines.Format](https://learn.microsoft.com/en-us/office/vba/api/excel.gridlines.format).
        */
        wxExcelChartFormat GetFormat();

        /**
        Returns a String value that represents the name of the object.

        [MSDN documentation for Gridlines.Name](https://learn.microsoft.com/en-us/office/vba/api/excel.gridlines.name).
        */
        wxString GetName();

        /**
        Returns "Gridlines".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("Gridlines"); }
    };


} // namespace wxAutoExcel


#endif // #if WXAUTOEXCEL_USE_CHARTS

#endif // _WXAUTOEXCEL_GRIDLINES_H
