/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Modified by:
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////


#include "wx/wxAutoExcel_prec.h"

#include "wx/wxAutoExcelSort.h"

#include "wx/wxAutoExcelSortFields.h"
#include "wx/wxAutoExcelRange.h"

#include "wx/wxAutoExcelPrivate.h"

namespace wxAutoExcel {

// ***** class wxExcelSort METHODS *****

void wxExcelSort::Apply()
{
    WXAUTOEXCEL_CALL_METHOD0_RET("Apply", "null");
}

void wxExcelSort::SetRange(wxExcelRange rng)
{
    wxVariant vRng;

    if ( ObjectToVariant(&rng, vRng) )
    {
        WXAUTOEXCEL_CALL_METHOD1_RET("SetRange", vRng, "null");
    }
}

// ***** class wxExcelSort PROPERTIES *****


XlYesNoGuess wxExcelSort::GetHeader()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("Header", XlYesNoGuess, xlGuess);
}

void wxExcelSort::SetHeader(XlYesNoGuess header)
{
    InvokePutProperty(wxS("Header"), (long)header);
}

bool wxExcelSort::GetMatchCase()
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET0("MatchCase");
}

void wxExcelSort::SetMatchCase(bool matchCase)
{
    InvokePutProperty(wxS("MatchCase"), matchCase);
}

XlSortOrientation wxExcelSort::GetOrientation()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("Orientation", XlSortOrientation, xlSortColumns);
}

void wxExcelSort::SetOrientation(XlSortOrientation orientation)
{
    InvokePutProperty(wxS("Orientation"), (long)orientation);
}


wxExcelRange wxExcelSort::GetRng()
{
    wxExcelRange range;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("Rng", range);
}

wxExcelSortFields wxExcelSort::GetSortFields()
{
    wxExcelSortFields sortFields;

    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("SortFields", sortFields);

}

XlSortMethod wxExcelSort::GetSortMethod()
{
    WXAUTOEXCEL_PROPERTY_ENUM_GET0("SortMethod", XlSortMethod, xlPinYin);
}



} // namespace wxAutoExcel
