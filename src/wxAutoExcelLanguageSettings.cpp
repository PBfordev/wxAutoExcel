/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Modified by:
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////


#include "wx/wxAutoExcel_prec.h"

#include "wx/wxAutoExcelLanguageSettings.h"

#include "wx/wxAutoExcel_enums.h"

#include "wx/wxAutoExcel_private.h"

namespace wxAutoExcel {

// ***** class wxExcelLanguageSettings PROPERTIES *****

WXLCID wxExcelLanguageSettings::GetLanguageID(MsoAppLanguageID id)
{
    WXAUTOEXCEL_PROPERTY_LONG_GET1("LanguageID", (long)id);
}

bool wxExcelLanguageSettings::GetLanguagePreferredForEditing(MsoLanguageID lid)
{
    WXAUTOEXCEL_PROPERTY_BOOL_GET1("LanguagePreferredForEditing", (long)lid);
}



} // namespace wxAutoExcel
