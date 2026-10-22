/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// License:     MIT license
/////////////////////////////////////////////////////////////////////////////


#ifndef _WXAUTOEXCEL_ERRORBARS_H
#define _WXAUTOEXCEL_ERRORBARS_H

#include "wx/wxAutoExcel_defs.h"

#if WXAUTOEXCEL_USE_CHARTS


#include "wx/wxAutoExcel_object.h"

namespace wxAutoExcel {

    /**
    @brief Represents Microsoft Excel ErrorBars object.
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelErrorBars : public wxExcelObject
    {
    public:
        // ***** METHODS *****

        /**
        Clears the formatting of the object.

        [MSDN documentation for ErrorBars.ClearFormats](https://learn.microsoft.com/en-us/office/vba/api/excel.errorbars.clearformats).
        */
        bool ClearFormats();

        /**
        Deletes the object.

        [MSDN documentation for ErrorBars.Delete](https://learn.microsoft.com/en-us/office/vba/api/excel.errorbars.delete).
        */
        bool Delete();

        /**
        Selects the object.

        [MSDN documentation for ErrorBars.Select](https://learn.microsoft.com/en-us/office/vba/api/excel.errorbars.select).
        */
        bool Select();

        // ***** PROPERTIES *****

        /**
        Returns a Border object that represents the border of the object.

        [MSDN documentation for ErrorBars.Border](https://learn.microsoft.com/en-us/office/vba/api/excel.errorbars.border).
        */
        wxExcelBorder GetBorder();

        /**
        Returns the end style for the error bars. Can be one of the following XlEndStyleCap constants: xlCap or xlNoCap.

        [MSDN documentation for ErrorBars.EndStyle](https://learn.microsoft.com/en-us/office/vba/api/excel.errorbars.endstyle).
        */
        XlEndStyleCap GetEndStyle();

        /**
        Sets the end style for the error bars. Can be one of the following XlEndStyleCap constants: xlCap or xlNoCap.

        [MSDN documentation for ErrorBars.EndStyle](https://learn.microsoft.com/en-us/office/vba/api/excel.errorbars.endstyle).
        */
        void SetEndStyle(XlEndStyleCap endStyle);

        /**
        Returns the ChartFormat object. Since Excel 2007.

        [MSDN documentation for ErrorBars.Format](https://learn.microsoft.com/en-us/office/vba/api/excel.errorbars.format).
        */
        wxExcelChartFormat GetFormat();

        /**
        Returns a String value that represents the name of the object.

        [MSDN documentation for ErrorBars.Name](https://learn.microsoft.com/en-us/office/vba/api/excel.errorbars.name).
        */
        wxString GetName();

        /**
        Returns "ErrorBars".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("ErrorBars"); }
    };


} // namespace wxAutoExcel


#endif // #if WXAUTOEXCEL_USE_CHARTS

#endif // _WXAUTOEXCEL_ERRORBARS_H
