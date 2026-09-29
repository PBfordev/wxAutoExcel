/////////////////////////////////////////////////////////////////////////////
// Purpose:     Shared NumberFormat getter compatible with Windows API macros
// Author:      OpenAI Codex, under the direction of PB
// Generated:   AI-generated using GPT-6 with High reasoning effort
// Copyright:   (c) 2026 PB <pbfordev@gmail.com>
// License:     MIT license
/////////////////////////////////////////////////////////////////////////////

#ifndef _WXAUTOEXCEL_NUMBERFORMATACCESS_H
#define _WXAUTOEXCEL_NUMBERFORMATACCESS_H

#include <wx/string.h>

// Windows can define GetNumberFormat as an A/W macro. Only this header needs
// to suspend it; the derived classes declare GetNumberFormatW explicitly.
#pragma push_macro("GetNumberFormat")
#undef GetNumberFormat

namespace wxAutoExcel {

/**
    @brief Provides the NumberFormat getter for Excel objects supporting it.

    This helper has no data members or virtual functions. Derived supplies
    GetNumberFormatW(), preserving its existing DLL entry point. Calls work
    whether or not the Windows GetNumberFormat macro is defined.
*/
template<typename Derived>
class wxExcelNumberFormatAccess
{
public:
    /**
        Returns the object's number format code.

        For Range and DisplayFormat objects, returns an empty string if the
        cells do not all have the same number format.
    */
    wxString GetNumberFormat()
    {
        return static_cast<Derived*>(this)->GetNumberFormatW();
    }

protected:
    wxExcelNumberFormatAccess() {}
    ~wxExcelNumberFormatAccess() {}
};

} // namespace wxAutoExcel

#pragma pop_macro("GetNumberFormat")

#endif // _WXAUTOEXCEL_NUMBERFORMATACCESS_H
