/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////


#ifndef _WXAUTOEXCEL_NAMES_H
#define _WXAUTOEXCEL_NAMES_H

#include "wx/wxAutoExcel_defs.h"

#include "wx/wxAutoExcel_object.h"
#include "wx/wxAutoExcel_enums.h"

namespace wxAutoExcel {

    /**
    @brief Represents Microsoft Excel Name object.
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelName : public wxExcelObject
    {
    public:
        // ***** METHODS *****

        /**
        Deletes the object.

        [MSDN documentation for Name.Delete](http://msdn.microsoft.com/en-us/library/bb211872).
        */
        void Delete();

        // ***** PROPERTIES *****

        /**
        Returns the category for the specified name in the language of the macro. The name must refer to a custom function or command.

        [MSDN documentation for Name.Category](http://msdn.microsoft.com/en-us/library/bb220906).
        */
        wxString GetCategory();

        /**
        Sets the category for the specified name in the language of the macro. The name must refer to a custom function or command.

        [MSDN documentation for Name.Category](http://msdn.microsoft.com/en-us/library/bb220906).
        */
        void SetCategory(const wxString& category);

        /**
        Returns the category for the specified name, in the language of the user, if the name refers to a custom function or command.

        [MSDN documentation for Name.CategoryLocal](http://msdn.microsoft.com/en-us/library/bb220907).
        */
        wxString GetCategoryLocal();

        /**
        Sets the category for the specified name, in the language of the user, if the name refers to a custom function or command.

        [MSDN documentation for Name.CategoryLocal](http://msdn.microsoft.com/en-us/library/bb220907).
        */
        void SetCategoryLocal(const wxString& categoryLocal);

        /**
        Returns the comment associated with the name. Since Excel 2007.

        [MSDN documentation for Name.Comment](http://msdn.microsoft.com/en-us/library/bb237157).
        */
        wxString GetComment();

        /**
        Sets the comment associated with the name. Since Excel 2007.

        [MSDN documentation for Name.Comment](http://msdn.microsoft.com/en-us/library/bb237157).
        */
        void SetComment(const wxString& comment);

        /**
        Returns a Long value that represents the index number of the object within the collection of similar objects.

        [MSDN documentation for Name.Index](http://msdn.microsoft.com/en-us/library/bb237160).
        */
        long GetIndex();

        /**
        Returns what the name refers to. Read/write XlXLMMacroType.

        [MSDN documentation for Name.MacroType](http://msdn.microsoft.com/en-us/library/bb208702).
        */
        XlXLMMacroType GetMacroType();

        /**
        Sets what the name refers to. Read/write XlXLMMacroType.

        [MSDN documentation for Name.MacroType](http://msdn.microsoft.com/en-us/library/bb208702).
        */
        void SetMacroType(XlXLMMacroType macroType);

        /**
        Returns a String value representing the name of the object.

        [MSDN documentation for Name.Name](http://msdn.microsoft.com/en-us/library/bb237163).
        */
        wxString GetName();

        /**
        Sets a String value representing the name of the object.

        [MSDN documentation for Name.Name](http://msdn.microsoft.com/en-us/library/bb237163).
        */
        void SetName(const wxString& name);

        /**
        Returns the name of the object, in the language of the user.

        [MSDN documentation for Name.NameLocal](http://msdn.microsoft.com/en-us/library/bb237167).
        */
        wxString GetNameLocal();

        /**
        Sets the name of the object, in the language of the user.

        [MSDN documentation for Name.NameLocal](http://msdn.microsoft.com/en-us/library/bb237167).
        */
        void SetNameLocal(const wxString& nameLocal);

        /**
        Returns the formula that the name is defined to refer to, in the language of the macro and in A1-style notation, beginning with an equal sign.

        [MSDN documentation for Name.RefersTo](http://msdn.microsoft.com/en-us/library/bb209077).
        */
        wxString GetRefersTo();

        /**
        Sets the formula that the name is defined to refer to, in the language of the macro and in A1-style notation, beginning with an equal sign.

        [MSDN documentation for Name.RefersTo](http://msdn.microsoft.com/en-us/library/bb209077).
        */
        void SetRefersTo(const wxString& refersTo);

        /**
        Returns the formula that the name refers to. The formula is in the language of the user, and it's in A1-style notation, beginning with an equal sign.

        [MSDN documentation for Name.RefersToLocal](http://msdn.microsoft.com/en-us/library/bb209080).
        */
        wxString GetRefersToLocal();

        /**
        Sets the formula that the name refers to. The formula is in the language of the user, and it's in A1-style notation, beginning with an equal sign.

        [MSDN documentation for Name.RefersToLocal](http://msdn.microsoft.com/en-us/library/bb209080).
        */
        void SetRefersToLocal(const wxString& refersToLocal);

        /**
        Returns the formula that the name refers to. The formula is in the language of the macro, and it's in R1C1-style notation, beginning with an equal sign.

        [MSDN documentation for Name.RefersToR1C1](http://msdn.microsoft.com/en-us/library/bb209083).
        */
        wxString GetRefersToR1C1();

        /**
        Sets the formula that the name refers to. The formula is in the language of the macro, and it's in R1C1-style notation, beginning with an equal sign.

        [MSDN documentation for Name.RefersToR1C1](http://msdn.microsoft.com/en-us/library/bb209083).
        */
        void SetRefersToR1C1(const wxString& refersToR1C1);

        /**
        Returns the formula that the name refers to. This formula is in the language of the user, and it's in R1C1-style notation, beginning with an equal sign.

        [MSDN documentation for Name.RefersToR1C1Local](http://msdn.microsoft.com/en-us/library/bb209085).
        */
        wxString GetRefersToR1C1Local();

        /**
        Sets the formula that the name refers to. This formula is in the language of the user, and it's in R1C1-style notation, beginning with an equal sign.

        [MSDN documentation for Name.RefersToR1C1Local](http://msdn.microsoft.com/en-us/library/bb209085).
        */
        void SetRefersToR1C1Local(const wxString& refersToR1C1Local);

        /**
        Returns the Range object referred to by a Name object.

        [MSDN documentation for Name.RefersToRange](http://msdn.microsoft.com/en-us/library/bb209088).
        */
        wxExcelRange GetRefersToRange();

        /**
        Returns the shortcut key for a name defined as a custom Microsoft Excel 4.0 macro command.

        [MSDN documentation for Name.ShortcutKey](http://msdn.microsoft.com/en-us/library/bb221671).
        */
        wxString GetShortcutKey();

        /**
        Sets the shortcut key for a name defined as a custom Microsoft Excel 4.0 macro command.

        [MSDN documentation for Name.ShortcutKey](http://msdn.microsoft.com/en-us/library/bb221671).
        */
        void SetShortcutKey(const wxString& shortcutKey);

        /**
        True if the specified Name object is a valid workbook parameter. Since Excel 2007.

        [MSDN documentation for Name.ValidWorkbookParameter](http://msdn.microsoft.com/en-us/library/bb240197).
        */
        bool GetValidWorkbookParameter();

        /**
        Returns a String value that represents the formula that the name is defined to refer to.

        [MSDN documentation for Name.Value](http://msdn.microsoft.com/en-us/library/bb214875).
        */
        wxString GetValue();

        /**
        Sets a String value that represents the formula that the name is defined to refer to.

        [MSDN documentation for Name.Value](http://msdn.microsoft.com/en-us/library/bb214875).
        */
        void SetValue(const wxString& value);

        /**
        Returns a Boolean value that determines whether the object is visible.

        [MSDN documentation for Name.Visible](http://msdn.microsoft.com/en-us/library/bb214877).
        */
        bool GetVisible();

        /**
        Sets a Boolean value that determines whether the object is visible.

        [MSDN documentation for Name.Visible](http://msdn.microsoft.com/en-us/library/bb214877).
        */
        void SetVisible(bool visible);

        /**
        Returns the specified Name object as a workbook parameter. Since Excel 2007.

        [MSDN documentation for Name.WorkbookParameter](http://msdn.microsoft.com/en-us/library/bb257124).
        */
        bool GetWorkbookParameter();

        /**
        Sets the specified Name object as a workbook parameter. Since Excel 2007.

        [MSDN documentation for Name.WorkbookParameter](http://msdn.microsoft.com/en-us/library/bb257124).
        */
        void SetWorkbookParameter(bool workbookParameter);

        /**
        Returns "Name".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("Name"); }
    };

    /**
    @brief Represents Microsoft Excel Names collection.
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelNames : public wxExcelObject
    {
    public:
        // ***** METHODS *****

        /**
        Defines a new name.

        [MSDN documentation for Names.Add](http://msdn.microsoft.com/en-us/library/bb211876).
        */
        wxExcelName Add(const wxString& name = wxEmptyString, const wxString& refersTo = wxEmptyString,
                        wxXlTribool visible = wxDefaultXlTribool, long* macroType = NULL,
                        const wxString& shortCutKey = wxEmptyString,
                        const wxString& nameLocal = wxEmptyString, const wxString& refersToLocal = wxEmptyString,
                        const wxString& categoryLocal = wxEmptyString,
                        const wxString& refersToR1C1 = wxEmptyString, const wxString& refersToR1C1Local = wxEmptyString);

        //@{
        /**
        Returns a single Name object from a Names collection.

        [MSDN documentation for Names.Item](http://msdn.microsoft.com/en-us/library/bb211879).
        */
        wxExcelName Item(const wxString& index = wxEmptyString,  const wxString& indexLocal = wxEmptyString,
                         const wxString& refersTo = wxEmptyString);
        wxExcelName Item(long index);
        wxExcelName operator[](long index);
        //@}

        // ***** PROPERTIES *****

        /**
        Returns a Long value that represents the number of objects in the collection.

        [MSDN documentation for Names.Count](http://msdn.microsoft.com/en-us/library/bb237170).
        */
        long GetCount();

        /**
        Returns "Names".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("Names"); }
    };


} // namespace wxAutoExcel

#endif //_WXAUTOEXCEL_NAMES_H
