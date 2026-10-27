/////////////////////////////////////////////////////////////////////////////
// Author:      Claude Opus 5.5 (Anthropic), under the direction of PB
// Copyright:   (c) 2026 PB <pbfordev@gmail.com>
// License:     MIT license
/////////////////////////////////////////////////////////////////////////////


#ifndef _WXAUTOEXCEL_CELLCONTROL_H
#define _WXAUTOEXCEL_CELLCONTROL_H

#include "wx/wxAutoExcel_defs.h"
#include "wx/wxAutoExcel_object.h"
#include "wx/wxAutoExcel_enums.h"

namespace wxAutoExcel {

    /**
    @brief Represents Microsoft Excel CellControl object, the cell controls,
    such as checkboxes, of a range. Since Microsoft 365 (2024).

    A cell with a checkbox has the Boolean value of the checkbox, so the
    checkbox is checked and unchecked by setting the value of the cell to
    true or false.

    @see wxExcelRange::GetCellControl(), wxExcelRange::ResetContents(),
    wxExcelRange::RemoveControls()
    */
    class WXDLLIMPEXP_WXAUTOEXCEL wxExcelCellControl : public wxExcelObject
    {
    public:
        // ***** METHODS *****

        /**
        Adds a checkbox to all cells of the range, replacing their cell
        controls. The values of the cells are not changed, empty cells
        get the value false.

        [MSDN documentation for CellControl](https://learn.microsoft.com/en-us/office/vba/api/excel.range.cellcontrol).
        */
        bool SetCheckbox();

        // ***** PROPERTIES *****

        /**
        Returns the type of the cell controls of the range, e.g. xlTypeCheckbox,
        or xlTypeNone for cells without a cell control.

        When the cells of the range have cell controls of different types,
        including cells without a cell control, Excel returns no value, so
        this returns xlTypeNone and IsLastResultNull_() returns true.

        [MSDN documentation for CellControl](https://learn.microsoft.com/en-us/office/vba/api/excel.range.cellcontrol).
        */
        XlCellControlType GetType();

        /**
        Returns "CellControl".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("CellControl"); }
    };

} // namespace wxAutoExcel

#endif //_WXAUTOEXCEL_CELLCONTROL_H
