/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Modified by:
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////


#ifndef _WXAUTOEXCEL_CHARACTERS_H
#define _WXAUTOEXCEL_CHARACTERS_H

#include "wx/wxAutoExcel_defs.h"
#include "wx/wxAutoExcelObject.h"

namespace wxAutoExcel {

    /**
    @brief Represents Microsoft Excel Characters.
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelCharacters: public wxExcelObject
    {
    public:

        // ***** METHODS *****

        /**
        Deletes the object.
        [MSDN documentation for Characters.Delete](http://msdn.microsoft.com/en-us/library/bb211613).
        */
        void Delete();

        /**
        Inserts a string preceding the selected characters.
        [MSDN documentation for Characters.Insert](http://msdn.microsoft.com/en-us/library/bb211621).
        */
        wxString Insert(const wxString& str);

        // ***** PROPERTIES *****

        /**
        Returns the text of this range of characters.

        [MSDN documentation for Characters.Caption](http://msdn.microsoft.com/en-us/library/bb179402).
        */
        wxString GetCaption();
        /**
        Sets the text of this range of characters.

        [MSDN documentation for Characters.Caption](http://msdn.microsoft.com/en-us/library/bb179402).
        */
        void SetCaption(wxString& caption);

        /**
        Returns the number of characters.

        [MSDN documentation for Characters.Count](http://msdn.microsoft.com/en-us/library/bb179406).
        */
        long GetCount();

        /**
        Returns the font of this range of characters.

        [MSDN documentation for Characters.Font](http://msdn.microsoft.com/en-us/library/bb179409).
        */
        wxExcelFont GetFont();
        

        /**
        Returns its text.

        [MSDN documentation for Characters.Text](http://msdn.microsoft.com/en-us/library/bb148819).
        */
        wxString GetText();

        /**
        Sets the text for the specified object.

        [MSDN documentation for Characters.Text](http://msdn.microsoft.com/en-us/library/bb148819).
        */
        void SetText(const wxString& text);

        /**
        Returns "Characters".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("Characters"); }
    };

} // namespace wxAutoExcel

#endif //_WXAUTOEXCEL_CHARACTERS_H
