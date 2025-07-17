/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2020 PB <pbfordev@gmail.com>
// License:     MIT license
/////////////////////////////////////////////////////////////////////////////

#include "wx/wxAutoExcel_prec.h"

#include "wx/wxAutoExcelListDataFormat.h"

#include "wx/wxAutoExcel_private.h"

namespace wxAutoExcel {

// ***** class wxExcelListDataFormat PROPERTIES *****

bool wxExcelListDataFormat::GetAllowFillIn()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("AllowFillIn");
}

wxArrayString wxExcelListDataFormat::GetChoices()
{
    wxArrayString strings;
    wxVariant vResult;

    if ( InvokeGetProperty(wxS("Choices"), vResult) )
    {
        const wxString type = vResult.GetType();

        if ( type == wxS("arrstring") )
        {
            return vResult.GetArrayString();
        }
        else if ( type == wxS("list") )
        {
            strings.reserve(vResult.GetCount());
            for ( size_t i = 0; i < vResult.GetCount(); i++)
                strings.push_back(vResult[i].GetString());
        }
        else
        {
            OnError(Err_InvalidReturnType, "Choices");
        }
    }

    return strings;
}

long wxExcelListDataFormat::GetDecimalPlaces()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("DecimalPlaces");
}

wxVariant wxExcelListDataFormat::GetDefaultValue()
{
    wxVariant vResult;

    InvokeGetProperty(wxS("DefaultValue"), vResult);
    return vResult;
}

bool wxExcelListDataFormat::GetIsPercent()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("IsPercent");
}

long wxExcelListDataFormat::Getlcid()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("lcid");
}

long wxExcelListDataFormat::GetMaxCharacters()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("MaxCharacters");
}

wxVariant wxExcelListDataFormat::GetMaxNumber()
{
    wxVariant vResult;

    InvokeGetProperty(wxS("MaxNumber"), vResult);
    return vResult;
}

wxVariant wxExcelListDataFormat::GetMinNumber()
{
    wxVariant vResult;

    InvokeGetProperty(wxS("MinNumber"), vResult);
    return vResult;
}

bool wxExcelListDataFormat::GetReadOnly()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("ReadOnly");
}

bool wxExcelListDataFormat::GetRequired()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("Required");
}

XlListDataType wxExcelListDataFormat::GetType()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("Type", XlListDataType, xlListDataTypeNone);
}

} // namespace wxAutoExcel
