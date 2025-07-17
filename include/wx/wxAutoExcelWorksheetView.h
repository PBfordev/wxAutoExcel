/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// License:     MIT license
/////////////////////////////////////////////////////////////////////////////


#ifndef _WXAUTOEXCEL_WORKSHEETVIEW_H
#define _WXAUTOEXCEL_WORKSHEETVIEW_H

#include "wx/wxAutoExcel_defs.h"

#include "wx/wxAutoExcel_object.h"

namespace wxAutoExcel {

    /**
    @brief Represents Microsoft Excel WorksheetView.
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelWorksheetView : public wxExcelObject
    {
    public:
        // ***** PROPERTIES *****

        /**
        Controls displaying  formulas. Since Excel 2007.

        [MSDN documentation for WorksheetView.DisplayFormulas](http://msdn.microsoft.com/en-us/library/bb211056.aspx).
        */
        bool GetDisplayFormulas();

        /**
        Controls displaying  formulas. Since Excel 2007.

        [MSDN documentation for WorksheetView.DisplayFormulas](http://msdn.microsoft.com/en-us/library/bb211056.aspx).
        */
        void SetDisplayFormulas(bool displayFormulas);

        /**
        Controls displaying  gridlines. Since Excel 2007.

        [MSDN documentation for WorksheetView.DisplayGridlines](http://msdn.microsoft.com/en-us/library/bb211060.aspx).
        */
        bool GetDisplayGridlines();

        /**
        Controls displaying  gridlines. Since Excel 2007.

        [MSDN documentation for WorksheetView.DisplayGridlines](http://msdn.microsoft.com/en-us/library/bb211060.aspx).
        */
        void SetDisplayGridlines(bool displayGridlines);

        /**
        Controls displaying  row and column headings. Since Excel 2007.

        [MSDN documentation for WorksheetView.DisplayHeadings](http://msdn.microsoft.com/en-us/library/bb211063.aspx).
        */
        bool GetDisplayHeadings();

        /**
        Controls displaying  row and column headings. Since Excel 2007.

        [MSDN documentation for WorksheetView.DisplayHeadings](http://msdn.microsoft.com/en-us/library/bb211063.aspx).
        */
        void SetDisplayHeadings(bool displayHeadings);

        /**
        Controls displaying  outline symbols. Since Excel 2007.

        [MSDN documentation for WorksheetView.DisplayOutline](http://msdn.microsoft.com/en-us/library/bb211066.aspx).
        */
        bool GetDisplayOutline();

        /**
        Controls displaying  outline symbols. Since Excel 2007.

        [MSDN documentation for WorksheetView.DisplayOutline](http://msdn.microsoft.com/en-us/library/bb211066.aspx).
        */
        void SetDisplayOutline(bool displayOutline);

        /**
        Controls displaying  zero values. Since Excel 2007.

        [MSDN documentation for WorksheetView.DisplayZeros](http://msdn.microsoft.com/en-us/library/bb211069.aspx).
        */
        bool GetDisplayZeros();

        /**
        Controls displaying  zero values. Since Excel 2007.

        [MSDN documentation for WorksheetView.DisplayZeros](http://msdn.microsoft.com/en-us/library/bb211069.aspx).
        */
        void SetDisplayZeros(bool displayZeros);


        /**
        Returns name of the sheet. Since Excel 2007.

        [MSDN documentation for WorksheetView.Sheet](http://msdn.microsoft.com/en-us/library/bb211073.aspx).
        */
        wxString GetSheet();

        /**
        Returns "WorksheetView".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("WorksheetView"); }

    };


} // namespace wxAutoExcel

#endif // _WXAUTOEXCEL_WORKSHEETVIEW_H
