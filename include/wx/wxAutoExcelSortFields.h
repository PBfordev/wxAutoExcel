/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Modified by:
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
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

        [MSDN documentation for SortField.Delete](http://msdn.microsoft.com/en-us/library/bb210542).
        */
        void Delete();

        /**
        Modify the key value by which values are sorted in the field.

        [MSDN documentation for SortField.ModifyKey](http://msdn.microsoft.com/en-us/library/bb210545).
        */
        void ModifyKey(wxExcelRange range);

        /**
        Sets an icon for a SortField object.

        [MSDN documentation for SortField.SetIcon](http://msdn.microsoft.com/en-us/library/bb242073).
        */
        void SetIcon(wxExcelIcon icon);

        // ***** PROPERTIES *****

        /**
        Specifies a custom order to sort the fields. Read/write Variant. Since Excel 2007.

        [MSDN documentation for SortField.CustomOrder](http://msdn.microsoft.com/en-us/library/bb148007).
        */
        wxString GetCustomOrder();

        /**
        Specifies a custom order to sort the fields. Read/write Variant. Since Excel 2007.

        [MSDN documentation for SortField.CustomOrder](http://msdn.microsoft.com/en-us/library/bb148007).
        */
        void SetCustomOrder(const wxString& customOrder);

        /**
        Specifies how to sort text in the range specified in SortField object. Read/write XlSortDataOption. Since Excel 2007.

        [MSDN documentation for SortField.DataOption](http://msdn.microsoft.com/en-us/library/bb148010).
        */
        XlSortDataOption GetDataOption();

        /**
        Specifies how to sort text in the range specified in SortField object. Read/write XlSortDataOption. Since Excel 2007.

        [MSDN documentation for SortField.DataOption](http://msdn.microsoft.com/en-us/library/bb148010).
        */
        void SetDataOption(XlSortDataOption dataOption);

        /**
        Specifies the range that is currently being sorted on. Since Excel 2007.

        [MSDN documentation for SortField.Key](http://msdn.microsoft.com/en-us/library/bb148014).
        */
        wxExcelRange GetKey();

        /**
        Determines the sort order for the values specified in the key. Since Excel 2007.

        [MSDN documentation for SortField.Order](http://msdn.microsoft.com/en-us/library/bb148018).
        */
        XlSortOrder GetOrder();

        /**
        Determines the sort order for the values specified in the key. Since Excel 2007.

        [MSDN documentation for SortField.Order](http://msdn.microsoft.com/en-us/library/bb148018).
        */
        void SetOrder(XlSortOrder order);

        /**
        Specifies the priority for the sort field. Since Excel 2007.

        [MSDN documentation for SortField.Priority](http://msdn.microsoft.com/en-us/library/bb148021).
        */
        long GetPriority();

        /**
        Specifies the priority for the sort field. Since Excel 2007.

        [MSDN documentation for SortField.Priority](http://msdn.microsoft.com/en-us/library/bb148021).
        */
        void SetPriority(long priority);

        /**
        Returns what attribute of the cell to sort on . Read/write XlSortOn. Since Excel 2007.

        [MSDN documentation for SortField.SortOn](http://msdn.microsoft.com/en-us/library/bb148027).
        */
        XlSortOn GetSortOn();

        /**
        Sets what attribute of the cell to sort on . Read/write XlSortOn. Since Excel 2007.

        [MSDN documentation for SortField.SortOn](http://msdn.microsoft.com/en-us/library/bb148027).
        */
        void SetSortOn(XlSortOn sortOn);

        /**
        Retun the value on which the sort is performed for the specified SortField object. Since Excel 2007.

        [MSDN documentation for SortField.SortOnValue](http://msdn.microsoft.com/en-us/library/bb148031).
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

        [MSDN documentation for SortFields.Add](http://msdn.microsoft.com/en-us/library/bb210551).
        */
        wxExcelSortField Add(wxExcelRange range, XlSortOn sortOn, 
                             XlSortOrder* order = NULL, const wxString& customOrder = wxEmptyString, XlSortDataOption* dataOption = NULL);

        /**
        Clears all the SortFields objects.

        [MSDN documentation for SortFields.Clear](http://msdn.microsoft.com/en-us/library/bb210554).
        */
        void Clear();

        // ***** PROPERTIES *****
        /**
        Returns the number of objects in the collection. Since Excel 2007.

        [MSDN documentation for SortFields.Count](http://msdn.microsoft.com/en-us/library/bb148024).
        */
        long GetCount();
        
        //@{
        /**
        Returns a SortField Represents a collection of items that can be sorted in a workbook. Since Excel 2007.

        [MSDN documentation for SortFields.Item](http://msdn.microsoft.com/en-us/library/bb226046).
        */
        wxExcelSortField GetItem(long index);
        wxExcelSortField operator[](long index);
        //@}

        /**
        Returns "SortFields".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("SortFields"); }
    };


} // namespace wxAutoExcel

#endif //_WXAUTOEXCEL_TAB_H
