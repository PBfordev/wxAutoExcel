/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Modified by:
// RCS-ID:      $Id: $
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////


#include "wx/wxAutoExcel_prec.h"

#include "wx/wxAutoExcelRange.h"

#include "wx/wxAutoExcelPrivate.h"

namespace wxAutoExcel {


wxExcelRange wxExcelRangeOwner::GetCells(long* row, long* column)
{
    wxExcelRange range;
    WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT(Row, row);
    WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT(Column, column);

    WXAUTOEXCEL_PROPERTY_OBJECT_GET2("Cells", vRow, vColumn, range);
}

wxExcelRange wxExcelRangeOwner::GetCells(long row, const wxString& column)
{
    wxExcelRange range;
    WXAUTOEXCEL_PROPERTY_OBJECT_GET2("Cells", row, column, range);
}


wxExcelRange wxExcelRangeOwner::GetRange(const wxString& cell1, const wxString& cell2)
{
    WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT(Cell2, (cell2.empty() ? NULL : &cell2));

    return DoGetRange(cell1, vCell2);
}


wxExcelRange wxExcelRangeOwner::GetRange(const wxExcelRange cell1, const wxExcelRange cell2)
{
    wxExcelRange range;
    wxVariant vCell1, vCell2;

    if ( ObjectToVariant(&cell1, vCell1) )
    {
        if ( ObjectToVariant(&cell2, vCell2) )
        {
            range = DoGetRange(vCell1, vCell2);
        }
        else
        {
            ReleaseVariantDispatch(vCell1);
        }
    }

    return range;
}



wxExcelRange wxExcelRangeOwner::GetRange(const wxExcelRange cell1, const wxString& cell2)
{
    wxExcelRange range;
    wxVariant vCell1;

    if ( ObjectToVariant(&cell1, vCell1)  )
    {
        range = DoGetRange(vCell1, cell2);
    }

    return range;
}

wxExcelRange wxExcelRangeOwner::GetRows()
{
    wxExcelRange range;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("Rows", range);
}

wxExcelRange wxExcelRangeOwner::GetColumns()
{
    wxExcelRange range;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("Columns", range);
}


wxExcelRange wxExcelRangeOwner::DoGetRangeItem(long rowIndex, const wxVariant& columnIndex)
{
    wxExcelRange range;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET2("Item", rowIndex, columnIndex, range);
}

wxExcelRange wxExcelRangeOwner::DoGetRange(const wxVariant& cell1, const wxVariant& cell2)
{
    wxExcelRange range;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET2("Range", cell1, cell2, range);
}


} // namespace wxAutoExcel
