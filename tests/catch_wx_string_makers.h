/////////////////////////////////////////////////////////////////////////////
// Purpose:     Catch2 string conversions for wxWidgets value types
// Author:      OpenAI Codex, under the direction of PB
// Generated:   AI-generated using GPT-5.6 Sol with High reasoning effort
// Copyright:   (c) 2026 PB <pbfordev@gmail.com>
// License:     MIT license
/////////////////////////////////////////////////////////////////////////////

#ifndef WXAUTOEXCEL_TESTS_CATCH_WX_STRING_MAKERS_H
#define WXAUTOEXCEL_TESTS_CATCH_WX_STRING_MAKERS_H

#include <catch2/catch_tostring.hpp>

#include <wx/colour.h>
#include <wx/datetime.h>

namespace Catch
{

template<>
struct StringMaker<wxColour>
{
    static std::string convert(const wxColour& value)
    {
        if ( !value.IsOk() )
            return "<invalid colour>";

        return value.GetAsString(wxC2S_CSS_SYNTAX).ToStdString();
    }
};

template<>
struct StringMaker<wxDateTime>
{
    static std::string convert(const wxDateTime& value)
    {
        if ( !value.IsValid() )
            return "<invalid date/time>";

        return value.FormatISOCombined().ToStdString();
    }
};

} // namespace Catch

#endif // WXAUTOEXCEL_TESTS_CATCH_WX_STRING_MAKERS_H
