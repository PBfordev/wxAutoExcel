/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
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

    [MSDN documentation for NegativeBarFormat.BorderColor](http://msdn.microsoft.com/en-us/library/office/ff840020(v=office.14).aspx).
    */
    wxExcelFormatColor GetBorderColor();

    /**
    Specifies whether to use the same border color as positive data bars.

    [MSDN documentation for NegativeBarFormat.BorderColorType](http://msdn.microsoft.com/en-us/library/office/ff192970(v=office.14).aspx).
    */
    XlDataBarNegativeColorType GetBorderColorType();

    /**
    Specifies whether to use the same border color as positive data bars.

    [MSDN documentation for NegativeBarFormat.BorderColorType](http://msdn.microsoft.com/en-us/library/office/ff192970(v=office.14).aspx).
    */
    void SetBorderColorType(XlDataBarNegativeColorType borderColorType);

    /**
    Returns a FormatColor object that you can use to specify the fill color for negative data bars.

    [MSDN documentation for NegativeBarFormat.Color](http://msdn.microsoft.com/en-us/library/office/ff820859(v=office.14).aspx).
    */
    wxExcelFormatColor GetColor();

    /**
    Specifies whether to use the same fill color as positive data bars.

    [MSDN documentation for NegativeBarFormat.ColorType](http://msdn.microsoft.com/en-us/library/office/ff192968(v=office.14).aspx).
    */
    XlDataBarNegativeColorType GetColorType();

    /**
    Specifies whether to use the same fill color as positive data bars.

    [MSDN documentation for NegativeBarFormat.ColorType](http://msdn.microsoft.com/en-us/library/office/ff192968(v=office.14).aspx).
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