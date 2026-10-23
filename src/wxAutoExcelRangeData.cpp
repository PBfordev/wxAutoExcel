/////////////////////////////////////////////////////////////////////////////
// Author:      Claude Opus 5.5 (Anthropic), under the direction of PB
// Copyright:   (c) 2026 PB <pbfordev@gmail.com>
// License:     MIT license
/////////////////////////////////////////////////////////////////////////////


#include "wx/wxAutoExcel_prec.h"

#include "wx/wxAutoExcelRangeData.h"

#include <wx/msw/ole/oleutils.h>
#include <wx/msw/ole/safearray.h>

namespace wxAutoExcel {

wxExcelRangeData::wxExcelRangeData(size_t rows, size_t columns)
    : m_rows(0), m_columns(0)
{
    Resize(rows, columns);
}

void wxExcelRangeData::Resize(size_t rows, size_t columns)
{
    if ( rows == 0 || columns == 0 )
    {
        Clear();
        return;
    }

    wxVector<wxVariant> cells(rows * columns);

    for ( size_t row = 0; row < rows && row < m_rows; ++row )
    {
        for ( size_t column = 0; column < columns && column < m_columns; ++column )
            cells[row * columns + column] = m_cells[row * m_columns + column];
    }

    m_cells.swap(cells);
    m_rows = rows;
    m_columns = columns;
}

void wxExcelRangeData::Clear()
{
    m_cells.clear();
    m_rows = 0;
    m_columns = 0;
}

wxVariant& wxExcelRangeData::operator()(size_t row, size_t column)
{
    wxASSERT_MSG( row < m_rows && column < m_columns, wxS("Invalid cell") );
    return m_cells[row * m_columns + column];
}

const wxVariant& wxExcelRangeData::operator()(size_t row, size_t column) const
{
    wxASSERT_MSG( row < m_rows && column < m_columns, wxS("Invalid cell") );
    return m_cells[row * m_columns + column];
}

wxVector<wxVariant> wxExcelRangeData::GetRow(size_t row) const
{
    wxVector<wxVariant> cells;

    wxCHECK_MSG( row < m_rows, cells, wxS("Invalid row") );

    cells.reserve(m_columns);
    for ( size_t column = 0; column < m_columns; ++column )
        cells.push_back(m_cells[row * m_columns + column]);
    return cells;
}

wxVector<wxVariant> wxExcelRangeData::GetColumn(size_t column) const
{
    wxVector<wxVariant> cells;

    wxCHECK_MSG( column < m_columns, cells, wxS("Invalid column") );

    cells.reserve(m_rows);
    for ( size_t row = 0; row < m_rows; ++row )
        cells.push_back(m_cells[row * m_columns + column]);
    return cells;
}

wxVariant wxExcelRangeData::ToVariant() const
{
    if ( IsEmpty() )
        return wxVariant();

    wxSafeArray<VT_VARIANT> safeArray;
    // the first dimension is the row and the second the column, as in Excel
    SAFEARRAYBOUND bounds[2];

    bounds[0].lLbound = 1;
    bounds[0].cElements = static_cast<ULONG>(m_rows);
    bounds[1].lLbound = 1;
    bounds[1].cElements = static_cast<ULONG>(m_columns);

    if ( !safeArray.Create(bounds, 2) )
        return wxVariant();

    long indices[2];

    for ( size_t row = 0; row < m_rows; ++row )
    {
        indices[0] = static_cast<long>(row) + 1;
        for ( size_t column = 0; column < m_columns; ++column )
        {
            const wxVariant& cell = m_cells[row * m_columns + column];

            // the elements of a new array are already empty, a null
            // wxVariant would be converted to VT_NULL instead
            if ( cell.IsNull() )
                continue;

            indices[1] = static_cast<long>(column) + 1;
            if ( !safeArray.SetElement(indices, cell) )
                return wxVariant();
        }
    }

    return wxVariant(new wxVariantDataSafeArray(safeArray.Detach()));
}

bool wxExcelRangeData::FromVariant(const wxVariant& value, wxExcelRangeData& data)
{
    data.Clear();

    if ( value.GetType() == wxS("list") )
        return false;

    if ( value.GetType() != wxS("safearray") )
    {
        // a single cell
        data.Resize(1, 1);
        data(0, 0) = value;
        return true;
    }

    SAFEARRAY* array = wxStaticCastVariantData(value.GetData(), wxVariantDataSafeArray)->GetValue();
    VARTYPE varType;

    if ( !array || FAILED(SafeArrayGetVartype(array, &varType)) || varType != VT_VARIANT
         || SafeArrayGetDim(array) != 2 )
        return false;

    wxSafeArray<VT_VARIANT> safeArray;

    if ( !safeArray.Attach(array) )
        return false;

    long rowLBound = 0, columnLBound = 0;
    bool ok = safeArray.GetLBound(1, rowLBound) && safeArray.GetLBound(2, columnLBound);

    if ( ok )
    {
        const size_t rows = safeArray.GetCount(1);
        const size_t columns = safeArray.GetCount(2);
        long indices[2];

        data.Resize(rows, columns);
        for ( size_t row = 0; ok && row < rows; ++row )
        {
            indices[0] = rowLBound + static_cast<long>(row);
            for ( size_t column = 0; ok && column < columns; ++column )
            {
                indices[1] = columnLBound + static_cast<long>(column);
                ok = safeArray.GetElement(indices, data(row, column));
            }
        }
    }

    // the array belongs to the caller
    wxUnusedVar(safeArray.Detach());

    if ( !ok )
        data.Clear();
    return ok;
}

} // namespace wxAutoExcel
