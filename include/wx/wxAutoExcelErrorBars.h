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

        [MSDN documentation for ErrorBars.ClearFormats](http://msdn.microsoft.com/en-us/library/bb211753).
        */
        bool ClearFormats();

        /**
        Deletes the object.

        [MSDN documentation for ErrorBars.Delete](http://msdn.microsoft.com/en-us/library/bb211756).
        */
        bool Delete();

        /**
        Selects the object.

        [MSDN documentation for ErrorBars.Select](http://msdn.microsoft.com/en-us/library/bb237889).
        */
        bool Select();

        // ***** PROPERTIES *****

        /**
        Returns a Border object that represents the border of the object.

        [MSDN documentation for ErrorBars.Border](http://msdn.microsoft.com/en-us/library/bb236962).
        */
        wxExcelBorder GetBorder();

        /**
        Returns the end style for the error bars. Can be one of the following XlEndStyleCap constants: xlCap or xlNoCap.

        [MSDN documentation for ErrorBars.EndStyle](http://msdn.microsoft.com/en-us/library/bb208459).
        */
        XlEndStyleCap GetEndStyle();

        /**
        Sets the end style for the error bars. Can be one of the following XlEndStyleCap constants: xlCap or xlNoCap.

        [MSDN documentation for ErrorBars.EndStyle](http://msdn.microsoft.com/en-us/library/bb208459).
        */
        void SetEndStyle(XlEndStyleCap endStyle);

        /**
        Returns the ChartFormat object. Since Excel 2007.

        [MSDN documentation for ErrorBars.Format](http://msdn.microsoft.com/en-us/library/bb242528).
        */
        wxExcelChartFormat GetFormat();

        /**
        Returns a String value that represents the name of the object.

        [MSDN documentation for ErrorBars.Name](http://msdn.microsoft.com/en-us/library/bb236965).
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
