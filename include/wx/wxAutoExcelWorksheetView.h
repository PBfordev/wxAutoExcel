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

        [MSDN documentation for WorksheetView.DisplayFormulas](https://learn.microsoft.com/en-us/office/vba/api/excel.worksheetview.displayformulas).
        */
        bool GetDisplayFormulas();

        /**
        Controls displaying  formulas. Since Excel 2007.

        [MSDN documentation for WorksheetView.DisplayFormulas](https://learn.microsoft.com/en-us/office/vba/api/excel.worksheetview.displayformulas).
        */
        void SetDisplayFormulas(bool displayFormulas);

        /**
        Controls displaying  gridlines. Since Excel 2007.

        [MSDN documentation for WorksheetView.DisplayGridlines](https://learn.microsoft.com/en-us/office/vba/api/excel.worksheetview.displaygridlines).
        */
        bool GetDisplayGridlines();

        /**
        Controls displaying  gridlines. Since Excel 2007.

        [MSDN documentation for WorksheetView.DisplayGridlines](https://learn.microsoft.com/en-us/office/vba/api/excel.worksheetview.displaygridlines).
        */
        void SetDisplayGridlines(bool displayGridlines);

        /**
        Controls displaying  row and column headings. Since Excel 2007.

        [MSDN documentation for WorksheetView.DisplayHeadings](https://learn.microsoft.com/en-us/office/vba/api/excel.worksheetview.displayheadings).
        */
        bool GetDisplayHeadings();

        /**
        Controls displaying  row and column headings. Since Excel 2007.

        [MSDN documentation for WorksheetView.DisplayHeadings](https://learn.microsoft.com/en-us/office/vba/api/excel.worksheetview.displayheadings).
        */
        void SetDisplayHeadings(bool displayHeadings);

        /**
        Controls displaying  outline symbols. Since Excel 2007.

        [MSDN documentation for WorksheetView.DisplayOutline](https://learn.microsoft.com/en-us/office/vba/api/excel.worksheetview.displayoutline).
        */
        bool GetDisplayOutline();

        /**
        Controls displaying  outline symbols. Since Excel 2007.

        [MSDN documentation for WorksheetView.DisplayOutline](https://learn.microsoft.com/en-us/office/vba/api/excel.worksheetview.displayoutline).
        */
        void SetDisplayOutline(bool displayOutline);

        /**
        Controls displaying  zero values. Since Excel 2007.

        [MSDN documentation for WorksheetView.DisplayZeros](https://learn.microsoft.com/en-us/office/vba/api/excel.worksheetview.displayzeros).
        */
        bool GetDisplayZeros();

        /**
        Controls displaying  zero values. Since Excel 2007.

        [MSDN documentation for WorksheetView.DisplayZeros](https://learn.microsoft.com/en-us/office/vba/api/excel.worksheetview.displayzeros).
        */
        void SetDisplayZeros(bool displayZeros);


        /**
        Returns the worksheet of the view. Since Excel 2007.

        [MSDN documentation for WorksheetView.Sheet](https://learn.microsoft.com/en-us/office/vba/api/excel.worksheetview.sheet).
        */
        wxExcelWorksheet GetSheet();

        /**
        Returns "WorksheetView".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("WorksheetView"); }

    };


} // namespace wxAutoExcel

#endif // _WXAUTOEXCEL_WORKSHEETVIEW_H
