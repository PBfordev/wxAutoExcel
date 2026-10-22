/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// License:     MIT license
/////////////////////////////////////////////////////////////////////////////


#ifndef _WXAUTOEXCEL_CHARACTERS_H
#define _WXAUTOEXCEL_CHARACTERS_H

#include "wx/wxAutoExcel_defs.h"
#include "wx/wxAutoExcel_object.h"

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
        [MSDN documentation for Characters.Delete](https://learn.microsoft.com/en-us/office/vba/api/excel.characters.delete).
        */
        void Delete();

        /**
        Inserts a string preceding the selected characters.
        [MSDN documentation for Characters.Insert](https://learn.microsoft.com/en-us/office/vba/api/excel.characters.insert).
        */
        wxString Insert(const wxString& str);

        // ***** PROPERTIES *****

        /**
        Returns the text of this range of characters.

        [MSDN documentation for Characters.Caption](https://learn.microsoft.com/en-us/office/vba/api/excel.characters.caption).
        */
        wxString GetCaption();
        /**
        Sets the text of this range of characters.

        [MSDN documentation for Characters.Caption](https://learn.microsoft.com/en-us/office/vba/api/excel.characters.caption).
        */
        void SetCaption(const wxString& caption);

        /**
        Returns the number of characters.

        [MSDN documentation for Characters.Count](https://learn.microsoft.com/en-us/office/vba/api/excel.characters.count).
        */
        long GetCount();

        /**
        Returns the font of this range of characters.

        [MSDN documentation for Characters.Font](https://learn.microsoft.com/en-us/office/vba/api/excel.characters.font).
        */
        wxExcelFont GetFont();


        /**
        Returns its text.

        [MSDN documentation for Characters.Text](https://learn.microsoft.com/en-us/office/vba/api/excel.characters.text).
        */
        wxString GetText();

        /**
        Sets the text for the specified object.

        [MSDN documentation for Characters.Text](https://learn.microsoft.com/en-us/office/vba/api/excel.characters.text).
        */
        void SetText(const wxString& text);

        /**
        Returns "Characters".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("Characters"); }
    };

} // namespace wxAutoExcel

#endif //_WXAUTOEXCEL_CHARACTERS_H
