/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// License:     MIT license
/////////////////////////////////////////////////////////////////////////////


#ifndef _WXAUTOEXCEL_SORTFIELDS_H
#define _WXAUTOEXCEL_SORTFIELDS_H

#include "wx/wxAutoExcel_defs.h"

#include "wx/wxAutoExcel_object.h"
#include "wx/wxAutoExcel_enums.h"


namespace wxAutoExcel {

    /**
    @brief Represents Microsoft Excel SortFields object.
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelSortField : public wxExcelObject
    {
    public:
        // ***** METHODS *****

        /**
        Removes the specified SortField object from the SortFields collection.

        [MSDN documentation for SortField.Delete](https://learn.microsoft.com/en-us/office/vba/api/excel.sortfield.delete).
        */
        void Delete();

        /**
        Modify the key value by which values are sorted in the field.

        [MSDN documentation for SortField.ModifyKey](https://learn.microsoft.com/en-us/office/vba/api/excel.sortfield.modifykey).
        */
        void ModifyKey(wxExcelRange range);

        /**
        Sets an icon for a SortField object.

        [MSDN documentation for SortField.SetIcon](https://learn.microsoft.com/en-us/office/vba/api/excel.sortfield.seticon).
        */
        void SetIcon(wxExcelIcon icon);

        // ***** PROPERTIES *****

        /**
        Specifies a custom order to sort the fields. Read/write Variant. Since Excel 2007.

        [MSDN documentation for SortField.CustomOrder](https://learn.microsoft.com/en-us/office/vba/api/excel.sortfield.customorder).
        */
        wxString GetCustomOrder();

        /**
        Specifies a custom order to sort the fields. Read/write Variant. Since Excel 2007.

        [MSDN documentation for SortField.CustomOrder](https://learn.microsoft.com/en-us/office/vba/api/excel.sortfield.customorder).
        */
        void SetCustomOrder(const wxString& customOrder);

        /**
        Specifies how to sort text in the range specified in SortField object. Read/write XlSortDataOption. Since Excel 2007.

        [MSDN documentation for SortField.DataOption](https://learn.microsoft.com/en-us/office/vba/api/excel.sortfield.dataoption).
        */
        XlSortDataOption GetDataOption();

        /**
        Specifies how to sort text in the range specified in SortField object. Read/write XlSortDataOption. Since Excel 2007.

        [MSDN documentation for SortField.DataOption](https://learn.microsoft.com/en-us/office/vba/api/excel.sortfield.dataoption).
        */
        void SetDataOption(XlSortDataOption dataOption);

        /**
        Specifies the range that is currently being sorted on. Since Excel 2007.

        [MSDN documentation for SortField.Key](https://learn.microsoft.com/en-us/office/vba/api/excel.sortfield.key).
        */
        wxExcelRange GetKey();

        /**
        Determines the sort order for the values specified in the key. Since Excel 2007.

        [MSDN documentation for SortField.Order](https://learn.microsoft.com/en-us/office/vba/api/excel.sortfield.order).
        */
        XlSortOrder GetOrder();

        /**
        Determines the sort order for the values specified in the key. Since Excel 2007.

        [MSDN documentation for SortField.Order](https://learn.microsoft.com/en-us/office/vba/api/excel.sortfield.order).
        */
        void SetOrder(XlSortOrder order);

        /**
        Specifies the priority for the sort field. Since Excel 2007.

        [MSDN documentation for SortField.Priority](https://learn.microsoft.com/en-us/office/vba/api/excel.sortfield.priority).
        */
        long GetPriority();

        /**
        Specifies the priority for the sort field. Since Excel 2007.

        [MSDN documentation for SortField.Priority](https://learn.microsoft.com/en-us/office/vba/api/excel.sortfield.priority).
        */
        void SetPriority(long priority);

        /**
        Returns what attribute of the cell to sort on . Read/write XlSortOn. Since Excel 2007.

        [MSDN documentation for SortField.SortOn](https://learn.microsoft.com/en-us/office/vba/api/excel.sortfield.sorton).
        */
        XlSortOn GetSortOn();

        /**
        Sets what attribute of the cell to sort on . Read/write XlSortOn. Since Excel 2007.

        [MSDN documentation for SortField.SortOn](https://learn.microsoft.com/en-us/office/vba/api/excel.sortfield.sorton).
        */
        void SetSortOn(XlSortOn sortOn);

        /**
        Retun the value on which the sort is performed for the specified SortField object. Since Excel 2007.

        [MSDN documentation for SortField.SortOnValue](https://learn.microsoft.com/en-us/office/vba/api/excel.sortfield.sortonvalue).
        */
        wxExcelObject GetSortOnValue();

        /**
        Returns "SortField".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("SortField"); }
    };

    /**
    @brief Represents Microsoft Excel SortFields collection.
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelSortFields : public wxExcelObject
    {
    public:
        // ***** METHODS *****

        /**
        Creates a new sort field and returns a SortFields object.

        [MSDN documentation for SortFields.Add](https://learn.microsoft.com/en-us/office/vba/api/excel.sortfields.add).
        */
        wxExcelSortField Add(wxExcelRange range, XlSortOn sortOn,
                             XlSortOrder* order = NULL, const wxString& customOrder = wxEmptyString, XlSortDataOption* dataOption = NULL);

        /**
        Clears all the SortFields objects.

        [MSDN documentation for SortFields.Clear](https://learn.microsoft.com/en-us/office/vba/api/excel.sortfields.clear).
        */
        void Clear();

        // ***** PROPERTIES *****
        /**
        Returns the number of objects in the collection. Since Excel 2007.

        [MSDN documentation for SortFields.Count](https://learn.microsoft.com/en-us/office/vba/api/excel.sortfields.count).
        */
        long GetCount();

        //@{
        /**
        Returns a SortField Represents a collection of items that can be sorted in a workbook. Since Excel 2007.

        [MSDN documentation for SortFields.Item](https://learn.microsoft.com/en-us/office/vba/api/excel.sortfields.item).
        */
        wxExcelSortField GetItem(long index);
        wxExcelSortField operator[](long index);
        //@}

        /**
        The iterator over the items, see wxExcelCollectionIterator.
        */
        typedef wxExcelCollectionIterator<wxExcelSortFields, wxExcelSortField> iterator;

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
        Returns "SortFields".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("SortFields"); }
    };


} // namespace wxAutoExcel

#endif //_WXAUTOEXCEL_TAB_H
