/////////////////////////////////////////////////////////////////////////////
// Author:      Claude Opus 5.5 (Anthropic), under the direction of PB
// Copyright:   (c) 2026 PB <pbfordev@gmail.com>
// License:     MIT license
/////////////////////////////////////////////////////////////////////////////


#ifndef _WXAUTOEXCEL_RANGEDATA_H
#define _WXAUTOEXCEL_RANGEDATA_H

#include <wx/variant.h>
#include <wx/vector.h>

#include "wx/wxAutoExcel_defs.h"

namespace wxAutoExcel {

    /**
    @brief A rectangular block of the values or formulas of range cells.

    The rows and columns are numbered from 0, so that the cell in the
    top left corner of the range is (0, 0). Each cell is a wxVariant, which
    is null for an empty cell, a double for a number, a string, a bool,
    a wxDateTime for a date (when not obtained with Value2), or a
    wxVariantDataErrorCode for an error value.

    Use wxExcelRange::GetValues_() and wxExcelRange::SetValues_(), or
    wxExcelRange::GetFormulas_() and wxExcelRange::SetFormulas_() to transfer
    the data at once, which is much faster than cell by cell:
    @code
    // range is a valid wxExcelRange of 3 rows and 2 columns
    wxExcelRangeData data(3, 2);

    for ( size_t row = 0; row < data.GetRowCount(); ++row )
    {
        data(row, 0) = wxString::Format("Item %d", static_cast<int>(row) + 1);
        data(row, 1) = 10.0 * row;
    }
    if ( !range.SetValues_(data) )
        wxLogError(_("Failed to set the values."));
    @endcode
    */
    class WXDLLIMPEXP_WXAUTOEXCEL wxExcelRangeData
    {
    public:
        /**
        Creates data of the given size, with all cells empty.
        */
        wxExcelRangeData(size_t rows = 0, size_t columns = 0);

        /**
        Returns the number of rows.
        */
        size_t GetRowCount() const { return m_rows; }

        /**
        Returns the number of columns.
        */
        size_t GetColumnCount() const { return m_columns; }

        /**
        Returns true if the data have no cells.
        */
        bool IsEmpty() const { return m_rows == 0 || m_columns == 0; }

        /**
        Changes the size. The cells which remain keep their values, the new
        cells are empty.
        */
        void Resize(size_t rows, size_t columns);

        /**
        Removes all the cells.
        */
        void Clear();

        //@{
        /**
        Returns the cell in the given row and column, numbered from 0.
        */
        wxVariant& operator()(size_t row, size_t column);
        const wxVariant& operator()(size_t row, size_t column) const;
        //@}

        /**
        Returns the cells of the given row, numbered from 0.
        */
        wxVector<wxVariant> GetRow(size_t row) const;

        /**
        Returns the cells of the given column, numbered from 0.
        */
        wxVector<wxVariant> GetColumn(size_t column) const;

        /**
        Converts the data to a wxVariant containing a two-dimensional
        SAFEARRAY of VARIANTs, which can be passed e.g. to wxExcelRange::SetValue().
        Returns a null wxVariant when the data are empty.

        The SAFEARRAY is destroyed when the wxVariant is passed to an
        automation call, otherwise the caller must destroy it with
        SafeArrayDestroy().
        */
        wxVariant ToVariant() const;

        /**
        Sets @a data from the value of a range returned by Excel, which is
        a two-dimensional SAFEARRAY of VARIANTs for a range with more than
        one cell, or a single value for a single cell.

        A SAFEARRAY is not destroyed. Returns false, leaving @a data empty,
        when @a value is a SAFEARRAY of another type or with another number
        of dimensions, or a list, which has no dimensions.
        */
        static bool FromVariant(const wxVariant& value, wxExcelRangeData& data);

    private:
        size_t m_rows;
        size_t m_columns;
        // the cells, row after row
        wxVector<wxVariant> m_cells;
    };

} // namespace wxAutoExcel

#endif //_WXAUTOEXCEL_RANGEDATA_H
