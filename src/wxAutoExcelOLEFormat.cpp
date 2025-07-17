/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////


#include "wx/wxAutoExcel_prec.h"


#include "wx/wxAutoExcelOLEFormat.h"

#if WXAUTOEXCEL_USE_SHAPES

#include "wx/wxAutoExcel_private.h"

namespace wxAutoExcel {

// ***** class wxExcelOLEFormat METHODS *****

void wxExcelOLEFormat::Activate()
{
    WXAUTOEXCEL_CALL_METHOD0_RET("Activate", "null");
}

void wxExcelOLEFormat::Verb(XlOLEVerb* verb)
{
    WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT_NAME(Verb, ((long*)verb));

    WXAUTOEXCEL_CALL_METHOD1_RET("Verb", vVerb, "null");
}

// ***** class wxExcelOLEFormat PROPERTIES *****

wxString wxExcelOLEFormat::GetprogID()
{
    WXAUTOEXCEL_PROPERTY_STRING_GET0("progID");
}

} // namespace wxAutoExcel


#endif // #if WXAUTOEXCEL_USE_SHAPES
