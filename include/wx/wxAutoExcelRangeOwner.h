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
        [MSDN documentation for Application.Cells](http://msdn.microsoft.com/en-us/library/bb212512.aspx).

        wxExcelWorksheet: Returns a Range representing all the cells on the worksheet (not just the cells that are currently in use.
        [MSDN documentation for Worksheet.Cells](http://msdn.microsoft.com/en-us/library/bb148836.aspx).

        wxExcelRange: Returns a Range representing the cells in the specified range.
        [MSDN documentation for Range.Cells](http://msdn.microsoft.com/en-us/library/bb213513.aspx).
        */
        wxExcelRange GetCells(long* row = NULL, long* column = NULL);
        wxExcelRange GetCells(long row, const wxString& column);
        //@}

        //@{
        /**
        wxExcelWorksheet: Returns a Range representing a cell or a range of cells.
        [MSDN documentation for Worksheet.Range](http://msdn.microsoft.com/en-us/library/bb238427.aspx).

        wxExcelRange: Returns a Range representing a cell or a range of cells.
        [MSDN documentation for Range.Range](http://msdn.microsoft.com/en-us/library/bb237494.aspx).
        */
        wxExcelRange GetRange(const wxString& cell1, const wxString& cell2 = wxEmptyString);
        wxExcelRange GetRange(const wxExcelRange cell1, const wxExcelRange cell2);
        wxExcelRange GetRange(const wxExcelRange cell1, const wxString& cell2);
        //@}

        /**
        wxExcelWorksheet: Returns a Range representing all the rows on the specified worksheet. Read-only Range object.
        [MSDN documentation for Worksheet.Rows](http://msdn.microsoft.com/en-us/library/bb215251.aspx).

        wxExcelRange: Returns a Range representing all the rows in the specified range. Read-only Range object.
        [MSDN documentation for Range.Rows](http://msdn.microsoft.com/en-us/library/bb238590.aspx).
        */
        wxExcelRange GetRows();

        /**
        wxExcelRange: Returns a Range containing just one row with rowIndex (starting with 1).
        [MSDN documentation for Range.Rows](http://msdn.microsoft.com/en-us/library/bb238590.aspx).
        */
        wxExcelRange GetRows(long rowIndex);

        /**
        wxExcelWorksheet: Returns a Range that can contain several contiguous rows.
        Pass the address in the format "firstRowIndex:lastRowIndex", e.g. "1:5" to get first five rows of the range.

        [MSDN documentation for Range.Rows](http://msdn.microsoft.com/en-us/library/bb238590.aspx).
        */
        wxExcelRange GetRows(const wxString& rowRange);

        /**
        wxExcelWorksheet: Returns a Range representing all the columns on the active worksheet. If the active document isn't a worksheet, the Columns property fails.
        [MSDN documentation for Worksheet.Columns](http://msdn.microsoft.com/en-us/library/bb148843.aspx).

        wxExcelRange: Returns a Range representing the columns in the specified range.
        [MSDN documentation for Range.Columns](http://msdn.microsoft.com/en-us/library/bb213515.aspx).
        */
        wxExcelRange GetColumns();

        /**
        wxExcelRange: Returns a Range containing just one column with columnIndex (starting with 1).
        [MSDN documentation for Range.Rows](http://msdn.microsoft.com/en-us/library/bb238590.aspx).
        */
        wxExcelRange GetColumns(long columnIndex);

        /**
        wxExcelWorksheet: Returns a Range that can contain several contiguous columns.
        Pass the address in the format "firstColumnLetter:lastColumnLetter", e.g. "A:E" to get first five columns of the range.

        [MSDN documentation for Range.Rows](http://msdn.microsoft.com/en-us/library/bb238590.aspx).
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

