/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Modified by:
// Copyright:   (c) 2020 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////

#include "wx/wxAutoExcel_prec.h"

#include "wx/wxAutoExcelTableStyleElement.h"

#include "wx/wxAutoExcelBorders.h"
#include "wx/wxAutoExcelFont.h"
#include "wx/wxAutoExcelInterior.h"

#include "wx/wxAutoExcel_private.h"

namespace wxAutoExcel {

// ***** class wxExcelTableStyleElement METHODS *****

void wxExcelTableStyleElement::Clear()
{
    WXAUTOEXCEL_CALL_METHOD0_RET("Clear", "null");
}

// ***** class wxExcelTableStyleElement PROPERTIES *****

wxExcelBorders wxExcelTableStyleElement::GetBorders()
{
    wxExcelBorders borders;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("Borders", borders);
}

wxExcelFont wxExcelTableStyleElement::GetFont()
{
    wxExcelFont font;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("Font", font);
}

bool wxExcelTableStyleElement::GetHasFormat()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("HasFormat");
}

wxExcelInterior wxExcelTableStyleElement::GetInterior()
{
    wxExcelInterior interior;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("Interior", interior);
}

long wxExcelTableStyleElement::GetStripeSize()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("StripeSize");
}

void wxExcelTableStyleElement::SetStripeSize(long stripeSize)
{
    InvokePutProperty(wxS("StripeSize"), stripeSize);
}


// ***** class wxExcelTableStyleElements METHODS *****


wxExcelTableStyleElement wxExcelTableStyleElements::GetItem(XlTableStyleElementType index)
{
    wxExcelTableStyleElement element;

    WXAUTOEXCEL_CALL_METHOD1_OBJECT("Item", index, element);
}

wxExcelTableStyleElement wxExcelTableStyleElements::operator[](XlTableStyleElementType index)
{
    return GetItem(index);
}

// ***** class wxExcelTableStyleElements PROPERTIES *****

long wxExcelTableStyleElements::GetCount()
{
    WXAUTOEXCEL_PROPERTY_LONG_GET0("Count");
}

} // namespace wxAutoExcel
