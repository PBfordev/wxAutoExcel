/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////


#ifndef _WXAUTOEXCEL_SHEET_H
#define _WXAUTOEXCEL_SHEET_H

#include "wx/wxAutoExcel_defs.h"

#include "wx/wxAutoExcel_object.h"
#include "wx/wxAutoExcel_enums.h"

namespace wxAutoExcel {
    /**
    @brief Represents a member of Microsoft Excel Sheets collection.
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelSheet : public wxExcelObject
    {
    public:
        /**
            Returns the name of the object.
        */
        wxString GetName();
        /**
            Returns the sheet type.
            May return wrong value for Charts, see e.g. here http://www.pcreview.co.uk/forums/xlchart-returning-wrong-value-t3579141.html
        */
        XlSheetType GetType();
        /**
            Returns true if the sheet is a worksheet.
        */
        bool IsWorksheet();
        /**
            Returns true if the sheet is a chart.
        */
        bool IsChart();

        /**
            Copies the sheet to the new workbook.
        */
        bool Copy();
        /**
            Copies the sheet within the same workbook, after or before the specified sheet.
        */
        bool CopyAfterOrBefore(wxExcelSheet sheetAfterOrBefore, bool after);

        /**
            Moves the sheet to the new workbook.
        */
        bool Move();
        /**
            Moves the sheet within the same workbook, after or before the specified sheet.
        */
        bool MoveAfterOrBefore(wxExcelSheet sheetAfterOrBefore, bool after);

        /**
            Deletes the sheet.
        */
        bool Delete();

        /**
            If possible, returns the sheet as a worksheet (the sheet object is still preserved)
        */
        wxExcelWorksheet ToWorksheet();

#if WXAUTOEXCEL_USE_CHARTS
        /**
            If possible, returns the sheet as a chart (the sheet object is still preserved)
        */
        wxExcelChart ToChart();
#endif // #if WXAUTOEXCEL_USE_CHARTS

        /**
        Returns "Sheet".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("Sheet"); }
    private:
        bool DoOrderedCopyOrMove(bool copy, wxExcelSheet sheetAfterOrBefore, bool after);
    };


} // namespace wxAutoExcel

#endif //_WXAUTOEXCEL_SHEET_H
