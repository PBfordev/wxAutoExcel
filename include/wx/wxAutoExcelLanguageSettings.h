/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
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

    [MSDN documentation for LanguageSettings.LanguageID](http://msdn.microsoft.com/en-us/library/office/ff863438(v=office.14).aspx).
    */
    WXLCID GetLanguageID(MsoAppLanguageID id);

    /**
    Gets True if the value for the MsoLanguageID constant has been identified in the Windows registry as a preferred language for editing.

    [MSDN documentation for LanguageSettings.LanguagePreferredForEditing](http://msdn.microsoft.com/en-us/library/office/ff861143(v=office.14).aspx).
    */
    bool GetLanguagePreferredForEditing(MsoLanguageID lid);

        /**
        Returns "LanguageSettings".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("LanguageSettings"); }
    };


} // namespace wxAutoExcel

#endif //_WXAUTOEXCEL_LANGUAGESETTINGS_H
