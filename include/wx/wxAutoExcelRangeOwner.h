/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// License:     MIT license
/////////////////////////////////////////////////////////////////////////////


#ifndef _WXAUTOEXCEL_RANGEOWNER_H
#define _WXAUTOEXCEL_RANGEOWNER_H

#include "wx/wxAutoExcel_defs.h"

#include "wx/wxAutoExcel_object.h"

namespace wxAutoExcel {

    /**
    @brief Helper object that contains methods common to objects that can return Ranges.
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelRangeOwner: public wxExcelObject
    {
    public:
        //@{
        /**
        wxExcelApplication: Returns a Range representing all the cells on the active worksheet. If the active document isn't a worksheet, this property fails.
        [MSDN documentation for Application.Cells](https://learn.microsoft.com/en-us/office/vba/api/excel.application.cells).

        wxExcelWorksheet: Returns a Range representing all the cells on the worksheet (not just the cells that are currently in use.
        [MSDN documentation for Worksheet.Cells](https://learn.microsoft.com/en-us/office/vba/api/excel.worksheet.cells).

        wxExcelRange: Returns a Range representing the cells in the specified range.
        [MSDN documentation for Range.Cells](https://learn.microsoft.com/en-us/office/vba/api/excel.range.cells).
        */
        wxExcelRange GetCells(long* row = NULL, long* column = NULL);
        wxExcelRange GetCells(long row, const wxString& column);
        //@}

        //@{
        /**
        wxExcelWorksheet: Returns a Range representing a cell or a range of cells.
        [MSDN documentation for Worksheet.Range](https://learn.microsoft.com/en-us/office/vba/api/excel.worksheet.range).

        wxExcelRange: Returns a Range representing a cell or a range of cells.
        [MSDN documentation for Range.Range](https://learn.microsoft.com/en-us/office/vba/api/excel.range.range).

        An address of a range with several areas, e.g. "A1:B2,D4:E5" in VBA, must use the list
        separator of the Windows regional settings instead of the comma, regardless of the
        automation LCID, e.g. "A1:B2;D4:E5" with a semicolon as the list separator. Build such
        addresses with wxExcelApplication::GetListSeparator_(), or combine the ranges of the areas
        with wxExcelApplication::Union() instead.
        */
        wxExcelRange GetRange(const wxString& cell1, const wxString& cell2 = wxEmptyString);
        wxExcelRange GetRange(const wxExcelRange cell1, const wxExcelRange cell2);
        wxExcelRange GetRange(const wxExcelRange cell1, const wxString& cell2);
        //@}

        /**
        wxExcelWorksheet: Returns a Range representing all the rows on the specified worksheet. Read-only Range object.
        [MSDN documentation for Worksheet.Rows](https://learn.microsoft.com/en-us/office/vba/api/excel.worksheet.rows).

        wxExcelRange: Returns a Range representing all the rows in the specified range. Read-only Range object.
        [MSDN documentation for Range.Rows](https://learn.microsoft.com/en-us/office/vba/api/excel.range.rows).
        */
        wxExcelRange GetRows();

        /**
        wxExcelRange: Returns a Range containing just one row with rowIndex (starting with 1).
        [MSDN documentation for Range.Rows](https://learn.microsoft.com/en-us/office/vba/api/excel.range.rows).
        */
        wxExcelRange GetRows(long rowIndex);

        /**
        wxExcelWorksheet: Returns a Range that can contain several contiguous rows.
        Pass the address in the format "firstRowIndex:lastRowIndex", e.g. "1:5" to get first five rows of the range.

        [MSDN documentation for Range.Rows](https://learn.microsoft.com/en-us/office/vba/api/excel.range.rows).
        */
        wxExcelRange GetRows(const wxString& rowRange);

        /**
        wxExcelWorksheet: Returns a Range representing all the columns on the active worksheet. If the active document isn't a worksheet, the Columns property fails.
        [MSDN documentation for Worksheet.Columns](https://learn.microsoft.com/en-us/office/vba/api/excel.worksheet.columns).

        wxExcelRange: Returns a Range representing the columns in the specified range.
        [MSDN documentation for Range.Columns](https://learn.microsoft.com/en-us/office/vba/api/excel.range.columns).
        */
        wxExcelRange GetColumns();

        /**
        wxExcelRange: Returns a Range containing just one column with columnIndex (starting with 1).
        [MSDN documentation for Range.Columns](https://learn.microsoft.com/en-us/office/vba/api/excel.range.columns).
        */
        wxExcelRange GetColumns(long columnIndex);

        /**
        wxExcelWorksheet: Returns a Range that can contain several contiguous columns.
        Pass the address in the format "firstColumnLetter:lastColumnLetter", e.g. "A:E" to get first five columns of the range.

        [MSDN documentation for Range.Columns](https://learn.microsoft.com/en-us/office/vba/api/excel.range.columns).
        */
        wxExcelRange GetColumns(const wxString& columnRange);

        /**
        Returns "RangeOwner (internal object)".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("RangeOwner (internal object)"); }
    protected:
        /**
        @cond PRIVATE
        */
        wxExcelRange DoGetRangeItem(long rowIndex, const wxVariant& columnIndex);
        wxExcelRange DoGetRange(const wxVariant& cell1, const wxVariant& cell2);
        /**
        @endcond
        */
   };

} // namespace wxAutoExcel



#endif // _WXAUTOEXCEL_RANGEOWNER_H

