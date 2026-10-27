/////////////////////////////////////////////////////////////////////////////
// Author:      Claude Opus 5.5 (Anthropic), under the direction of PB
// Copyright:   (c) 2026 PB <pbfordev@gmail.com>
// License:     MIT license
/////////////////////////////////////////////////////////////////////////////


#include "wx/wxAutoExcel_prec.h"

#include "wx/wxAutoExcelCellControl.h"

#include "wx/wxAutoExcel_private.h"

namespace wxAutoExcel {

// ***** class wxExcelCellControl METHODS *****

bool wxExcelCellControl::SetCheckbox()
{
    WXAUTOEXCEL_CALL_METHOD0_BOOL("SetCheckbox");
}

// ***** class wxExcelCellControl PROPERTIES *****

XlCellControlType wxExcelCellControl::GetType()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("Type", XlCellControlType, xlTypeNone);
}

} // namespace wxAutoExcel
