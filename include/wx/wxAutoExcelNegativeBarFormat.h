/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// License:     MIT license
/////////////////////////////////////////////////////////////////////////////


#ifndef _WXAUTOEXCEL_NEGATIVEBARFORMAT_H
#define _WXAUTOEXCEL_NEGATIVEBARFORMAT_H

#include "wx/wxAutoExcel_defs.h"

#if WXAUTOEXCEL_USE_CONDFORMAT

#include "wx/wxAutoExcel_object.h"
#include "wx/wxAutoExcel_enums.h"

namespace wxAutoExcel {
    /**
    @brief Represents a Microsoft Excel NegativeBarFormat object.
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelNegativeBarFormat : public wxExcelObject
   {
   public:
    // ***** PROPERTIES *****

    /**
    Returns a FormatColor object that you can use to specify the border color for negative data bars.

    [MSDN documentation for NegativeBarFormat.BorderColor](https://learn.microsoft.com/en-us/office/vba/api/excel.negativebarformat.bordercolor).aspx).
    */
    wxExcelFormatColor GetBorderColor();

    /**
    Specifies whether to use the same border color as positive data bars.

    [MSDN documentation for NegativeBarFormat.BorderColorType](https://learn.microsoft.com/en-us/office/vba/api/excel.negativebarformat.bordercolortype).aspx).
    */
    XlDataBarNegativeColorType GetBorderColorType();

    /**
    Specifies whether to use the same border color as positive data bars.

    [MSDN documentation for NegativeBarFormat.BorderColorType](https://learn.microsoft.com/en-us/office/vba/api/excel.negativebarformat.bordercolortype).aspx).
    */
    void SetBorderColorType(XlDataBarNegativeColorType borderColorType);

    /**
    Returns a FormatColor object that you can use to specify the fill color for negative data bars.

    [MSDN documentation for NegativeBarFormat.Color](https://learn.microsoft.com/en-us/office/vba/api/excel.negativebarformat.color).aspx).
    */
    wxExcelFormatColor GetColor();

    /**
    Specifies whether to use the same fill color as positive data bars.

    [MSDN documentation for NegativeBarFormat.ColorType](https://learn.microsoft.com/en-us/office/vba/api/excel.negativebarformat.colortype).aspx).
    */
    XlDataBarNegativeColorType GetColorType();

    /**
    Specifies whether to use the same fill color as positive data bars.

    [MSDN documentation for NegativeBarFormat.ColorType](https://learn.microsoft.com/en-us/office/vba/api/excel.negativebarformat.colortype).aspx).
    */
    void SetColorType(XlDataBarNegativeColorType colorType);

    /**
    Returns "NegativeBarFormat".
    */
    virtual wxString GetAutoExcelObjectName_() const { return wxS("NegativeBarFormat"); }

    };

} // namespace wxAutoExcel

#endif // WXAUTOEXCEL_USE_CONDFORMAT

#endif // _WXAUTOEXCEL_NEGATIVEBARFORMAT_H