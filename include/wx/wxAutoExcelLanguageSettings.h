/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// License:     MIT license
/////////////////////////////////////////////////////////////////////////////


#ifndef _WXAUTOEXCEL_LANGUAGESETTINGS_H
#define _WXAUTOEXCEL_LANGUAGESETTINGS_H

#include "wx/wxAutoExcel_defs.h"

#include "wx/wxAutoExcel_object.h"
#include "wx/wxAutoExcel_enums.h"


namespace wxAutoExcel {


    /**
    @brief Represents Microsoft Excel LanguageSettings object.
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelLanguageSettings : public wxExcelObject
    {
    public:

    // ***** PROPERTIES *****

    /**
    Gets a MsoAppLanguageID constant representing the locale identifier (LCID) for the install language, the user interface language, or the Help language.

    [MSDN documentation for LanguageSettings.LanguageID](https://learn.microsoft.com/en-us/office/vba/api/office.languagesettings.languageid).aspx).
    */
    WXLCID GetLanguageID(MsoAppLanguageID id);

    /**
    Gets True if the value for the MsoLanguageID constant has been identified in the Windows registry as a preferred language for editing.

    [MSDN documentation for LanguageSettings.LanguagePreferredForEditing](https://learn.microsoft.com/en-us/office/vba/api/office.languagesettings.languagepreferredforediting).aspx).
    */
    bool GetLanguagePreferredForEditing(MsoLanguageID lid);

        /**
        Returns "LanguageSettings".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("LanguageSettings"); }
    };


} // namespace wxAutoExcel

#endif //_WXAUTOEXCEL_LANGUAGESETTINGS_H
