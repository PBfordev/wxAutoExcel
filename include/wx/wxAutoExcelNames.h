/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// License:     MIT license
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

        [MSDN documentation for Name.Delete](https://learn.microsoft.com/en-us/office/vba/api/excel.name.delete).
        */
        void Delete();

        // ***** PROPERTIES *****

        /**
        Returns the category for the specified name in the language of the macro. The name must refer to a custom function or command.

        [MSDN documentation for Name.Category](https://learn.microsoft.com/en-us/office/vba/api/excel.name.category).
        */
        wxString GetCategory();

        /**
        Sets the category for the specified name in the language of the macro. The name must refer to a custom function or command.

        [MSDN documentation for Name.Category](https://learn.microsoft.com/en-us/office/vba/api/excel.name.category).
        */
        void SetCategory(const wxString& category);

        /**
        Returns the category for the specified name, in the language of the user, if the name refers to a custom function or command.

        [MSDN documentation for Name.CategoryLocal](https://learn.microsoft.com/en-us/office/vba/api/excel.name.categorylocal).
        */
        wxString GetCategoryLocal();

        /**
        Sets the category for the specified name, in the language of the user, if the name refers to a custom function or command.

        [MSDN documentation for Name.CategoryLocal](https://learn.microsoft.com/en-us/office/vba/api/excel.name.categorylocal).
        */
        void SetCategoryLocal(const wxString& categoryLocal);

        /**
        Returns the comment associated with the name. Since Excel 2007.

        [MSDN documentation for Name.Comment](https://learn.microsoft.com/en-us/office/vba/api/excel.name.comment).
        */
        wxString GetComment();

        /**
        Sets the comment associated with the name. Since Excel 2007.

        [MSDN documentation for Name.Comment](https://learn.microsoft.com/en-us/office/vba/api/excel.name.comment).
        */
        void SetComment(const wxString& comment);

        /**
        Returns a Long value that represents the index number of the object within the collection of similar objects.

        [MSDN documentation for Name.Index](https://learn.microsoft.com/en-us/office/vba/api/excel.name.index).
        */
        long GetIndex();

        /**
        Returns what the name refers to. Read/write XlXLMMacroType.

        [MSDN documentation for Name.MacroType](https://learn.microsoft.com/en-us/office/vba/api/excel.name.macrotype).
        */
        XlXLMMacroType GetMacroType();

        /**
        Sets what the name refers to. Read/write XlXLMMacroType.

        [MSDN documentation for Name.MacroType](https://learn.microsoft.com/en-us/office/vba/api/excel.name.macrotype).
        */
        void SetMacroType(XlXLMMacroType macroType);

        /**
        Returns a String value representing the name of the object.

        [MSDN documentation for Name.Name](https://learn.microsoft.com/en-us/office/vba/api/excel.name.name).
        */
        wxString GetName();

        /**
        Sets a String value representing the name of the object.

        [MSDN documentation for Name.Name](https://learn.microsoft.com/en-us/office/vba/api/excel.name.name).
        */
        void SetName(const wxString& name);

        /**
        Returns the name of the object, in the language of the user.

        [MSDN documentation for Name.NameLocal](https://learn.microsoft.com/en-us/office/vba/api/excel.name.namelocal).
        */
        wxString GetNameLocal();

        /**
        Sets the name of the object, in the language of the user.

        [MSDN documentation for Name.NameLocal](https://learn.microsoft.com/en-us/office/vba/api/excel.name.namelocal).
        */
        void SetNameLocal(const wxString& nameLocal);

        /**
        Returns the formula that the name is defined to refer to, in the language of the macro and in A1-style notation, beginning with an equal sign.

        [MSDN documentation for Name.RefersTo](https://learn.microsoft.com/en-us/office/vba/api/excel.name.refersto).
        */
        wxString GetRefersTo();

        /**
        Sets the formula that the name is defined to refer to, in the language of the macro and in A1-style notation, beginning with an equal sign.

        [MSDN documentation for Name.RefersTo](https://learn.microsoft.com/en-us/office/vba/api/excel.name.refersto).
        */
        void SetRefersTo(const wxString& refersTo);

        /**
        Returns the formula that the name refers to. The formula is in the language of the user, and it's in A1-style notation, beginning with an equal sign.

        [MSDN documentation for Name.RefersToLocal](https://learn.microsoft.com/en-us/office/vba/api/excel.name.referstolocal).
        */
        wxString GetRefersToLocal();

        /**
        Sets the formula that the name refers to. The formula is in the language of the user, and it's in A1-style notation, beginning with an equal sign.

        [MSDN documentation for Name.RefersToLocal](https://learn.microsoft.com/en-us/office/vba/api/excel.name.referstolocal).
        */
        void SetRefersToLocal(const wxString& refersToLocal);

        /**
        Returns the formula that the name refers to. The formula is in the language of the macro, and it's in R1C1-style notation, beginning with an equal sign.

        [MSDN documentation for Name.RefersToR1C1](https://learn.microsoft.com/en-us/office/vba/api/excel.name.referstor1c1).
        */
        wxString GetRefersToR1C1();

        /**
        Sets the formula that the name refers to. The formula is in the language of the macro, and it's in R1C1-style notation, beginning with an equal sign.

        [MSDN documentation for Name.RefersToR1C1](https://learn.microsoft.com/en-us/office/vba/api/excel.name.referstor1c1).
        */
        void SetRefersToR1C1(const wxString& refersToR1C1);

        /**
        Returns the formula that the name refers to. This formula is in the language of the user, and it's in R1C1-style notation, beginning with an equal sign.

        [MSDN documentation for Name.RefersToR1C1Local](https://learn.microsoft.com/en-us/office/vba/api/excel.name.referstor1c1local).
        */
        wxString GetRefersToR1C1Local();

        /**
        Sets the formula that the name refers to. This formula is in the language of the user, and it's in R1C1-style notation, beginning with an equal sign.

        [MSDN documentation for Name.RefersToR1C1Local](https://learn.microsoft.com/en-us/office/vba/api/excel.name.referstor1c1local).
        */
        void SetRefersToR1C1Local(const wxString& refersToR1C1Local);

        /**
        Returns the Range object referred to by a Name object.

        [MSDN documentation for Name.RefersToRange](https://learn.microsoft.com/en-us/office/vba/api/excel.name.referstorange).
        */
        wxExcelRange GetRefersToRange();

        /**
        Returns the shortcut key for a name defined as a custom Microsoft Excel 4.0 macro command.

        [MSDN documentation for Name.ShortcutKey](https://learn.microsoft.com/en-us/office/vba/api/excel.name.shortcutkey).
        */
        wxString GetShortcutKey();

        /**
        Sets the shortcut key for a name defined as a custom Microsoft Excel 4.0 macro command.

        [MSDN documentation for Name.ShortcutKey](https://learn.microsoft.com/en-us/office/vba/api/excel.name.shortcutkey).
        */
        void SetShortcutKey(const wxString& shortcutKey);

        /**
        True if the specified Name object is a valid workbook parameter. Since Excel 2007.

        [MSDN documentation for Name.ValidWorkbookParameter](https://learn.microsoft.com/en-us/office/vba/api/excel.name.validworkbookparameter).
        */
        bool GetValidWorkbookParameter();

        /**
        Returns a String value that represents the formula that the name is defined to refer to.

        [MSDN documentation for Name.Value](https://learn.microsoft.com/en-us/office/vba/api/excel.name.value).
        */
        wxString GetValue();

        /**
        Sets a String value that represents the formula that the name is defined to refer to.

        [MSDN documentation for Name.Value](https://learn.microsoft.com/en-us/office/vba/api/excel.name.value).
        */
        void SetValue(const wxString& value);

        /**
        Returns a Boolean value that determines whether the object is visible.

        [MSDN documentation for Name.Visible](https://learn.microsoft.com/en-us/office/vba/api/excel.name.visible).
        */
        bool GetVisible();

        /**
        Sets a Boolean value that determines whether the object is visible.

        [MSDN documentation for Name.Visible](https://learn.microsoft.com/en-us/office/vba/api/excel.name.visible).
        */
        void SetVisible(bool visible);

        /**
        Returns the specified Name object as a workbook parameter. Since Excel 2007.

        [MSDN documentation for Name.WorkbookParameter](https://learn.microsoft.com/en-us/office/vba/api/excel.name.workbookparameter).
        */
        bool GetWorkbookParameter();

        /**
        Sets the specified Name object as a workbook parameter. Since Excel 2007.

        [MSDN documentation for Name.WorkbookParameter](https://learn.microsoft.com/en-us/office/vba/api/excel.name.workbookparameter).
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

        [MSDN documentation for Names.Add](https://learn.microsoft.com/en-us/office/vba/api/excel.names.add).
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

        [MSDN documentation for Names.Item](https://learn.microsoft.com/en-us/office/vba/api/excel.names.item).
        */
        wxExcelName Item(const wxString& index = wxEmptyString,  const wxString& indexLocal = wxEmptyString,
                         const wxString& refersTo = wxEmptyString);
        wxExcelName Item(long index);
        wxExcelName operator[](long index);
        //@}

        // ***** PROPERTIES *****

        /**
        Returns a Long value that represents the number of objects in the collection.

        [MSDN documentation for Names.Count](https://learn.microsoft.com/en-us/office/vba/api/excel.names.count).
        */
        long GetCount();

        /**
        The iterator over the items, see wxExcelCollectionIterator.
        */
        typedef wxExcelCollectionIterator<wxExcelNames, wxExcelName> iterator;

        //@{
        /**
        Returns the iterator referring to the first item, or past the last item,
        which allows iterating over the items with a range-based for loop,
        see wxExcelCollectionIterator.
        */
        iterator begin() { return iterator(this, 1); }
        iterator end() { return iterator(this, GetCount() + 1); }
        //@}

        /**
        Returns "Names".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("Names"); }
    };


} // namespace wxAutoExcel

#endif //_WXAUTOEXCEL_NAMES_H
