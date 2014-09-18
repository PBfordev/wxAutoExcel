/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Modified by:
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////


#include "wx/wxAutoExcel_prec.h"

#include "wx/wxAutoExcelConditionValue.h"

#if WXAUTOEXCEL_USE_CONDFORMAT

#include "wx/wxAutoExcelPrivate.h"

namespace wxAutoExcel {

// ***** class wxExcelConditionValue METHODS *****

void wxExcelConditionValue::Modify(XlConditionValueTypes newType, const wxVariant& newValue)
{
    wxVariantVector args;
    wxVariant v;

    v = (long)newType;
    v.SetName(wxS("NewType"));
    args.push_back(v);
    
    if ( !newValue.IsNull() )
    {
        v = newValue;
        v.SetName(wxS("NewValue"));
        args.push_back(v);
    }
    
    if ( !newValue.IsNull() )
    {
        v = newValue;
        v.SetName(wxS("NewValue"));
        args.push_back(v);
    }
    WXAUTOEXCEL_CALL_METHODARR_RET("Modify", args, "null");
}

// ***** class wxExcelConditionValue PROPERTIES *****


XlConditionValueTypes wxExcelConditionValue::GetType()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("Type", XlConditionValueTypes, xlConditionValueNone);
}

wxVariant wxExcelConditionValue::GetValue()
{
    wxVariant vResult;

    InvokeGetProperty(wxS("Value"), vResult);
    return vResult;
}

void wxExcelConditionValue::SetValue(const wxVariant& value)
{
    InvokePutProperty(wxS("Value"), value);
}

} // namespace wxAutoExcel


#endif // WXAUTOEXCEL_USE_CONDFORMAT