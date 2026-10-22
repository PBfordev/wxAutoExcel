/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// License:     MIT license
/////////////////////////////////////////////////////////////////////////////


#ifndef _WXAUTOEXCEL_CONTROLFORMAT_H
#define _WXAUTOEXCEL_CONTROLFORMAT_H

#include "wx/wxAutoExcel_defs.h"

#if WXAUTOEXCEL_USE_SHAPES

#include "wx/wxAutoExcel_object.h"

namespace wxAutoExcel {
    /**
    @brief Represents Microsoft Excel ControlFormat object.
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelControlFormat : public wxExcelObject
    {
    public:

        // ***** METHODS *****

        /**
        Adds an item to a list box or a combo box.

        [MSDN documentation for ControlFormat.AddItem](https://learn.microsoft.com/en-us/office/vba/api/excel.controlformat.additem).
        */
        void AddItem(const wxString& text, long* index);

        //@{
        /**
        Returns or sets the text entries in the specified list box or a combo box, as an array of strings, or returns or sets a single text entry. An error occurs if there are no entries in the list.

        Although Excel declares List as a method, it can only be called as a property,
        which also allows setting the entries.

        [MSDN documentation for ControlFormat.List](https://learn.microsoft.com/en-us/office/vba/api/excel.controlformat.list).
        */
        wxString List(long index);
        wxArrayString List();

        void SetList(const wxArrayString& items);
        void SetList(long index, const wxString& item);
        //@}

        /**
        Removes all entries from a Microsoft Excel list box or combo box.

        [MSDN documentation for ControlFormat.RemoveAllItems](https://learn.microsoft.com/en-us/office/vba/api/excel.controlformat.removeallitems).
        */
        void RemoveAllItems();

        /**
        Removes one or more items from a list box or combo box.

        [MSDN documentation for ControlFormat.RemoveItem](https://learn.microsoft.com/en-us/office/vba/api/excel.controlformat.removeitem).
        */
        void RemoveItem(long index, long* count = NULL);

        // ***** PROPERTIES *****


        /**
        Returns the number of list lines displayed in the drop-down portion of a combo box.

        [MSDN documentation for ControlFormat.DropDownLines](https://learn.microsoft.com/en-us/office/vba/api/excel.controlformat.dropdownlines).
        */
        long GetDropDownLines();

        /**
        Sets the number of list lines displayed in the drop-down portion of a combo box.

        [MSDN documentation for ControlFormat.DropDownLines](https://learn.microsoft.com/en-us/office/vba/api/excel.controlformat.dropdownlines).
        */
        void SetDropDownLines(long dropDownLines);

        /**
        True if the object is enabled.

        [MSDN documentation for ControlFormat.Enabled](https://learn.microsoft.com/en-us/office/vba/api/excel.controlformat.enabled).
        */
        bool GetEnabled();

        /**
        True if the object is enabled.

        [MSDN documentation for ControlFormat.Enabled](https://learn.microsoft.com/en-us/office/vba/api/excel.controlformat.enabled).
        */
        void SetEnabled(bool enabled);

        /**
        Returns the amount that the scroll box increments or decrements for a page scroll (when the user clicks in the scroll bar body region).

        [MSDN documentation for ControlFormat.LargeChange](https://learn.microsoft.com/en-us/office/vba/api/excel.controlformat.largechange).
        */
        long GetLargeChange();

        /**
        Sets the amount that the scroll box increments or decrements for a page scroll (when the user clicks in the scroll bar body region).

        [MSDN documentation for ControlFormat.LargeChange](https://learn.microsoft.com/en-us/office/vba/api/excel.controlformat.largechange).
        */
        void SetLargeChange(long largeChange);

        /**
        Returns the worksheet range linked to the control's value. If you place a value in the cell, the control takes this value. Likewise, if you change the value of the control, that value is also placed in the cell.

        [MSDN documentation for ControlFormat.LinkedCell](https://learn.microsoft.com/en-us/office/vba/api/excel.controlformat.linkedcell).
        */
        wxString GetLinkedCell();

        /**
        Sets the worksheet range linked to the control's value. If you place a value in the cell, the control takes this value. Likewise, if you change the value of the control, that value is also placed in the cell.

        [MSDN documentation for ControlFormat.LinkedCell](https://learn.microsoft.com/en-us/office/vba/api/excel.controlformat.linkedcell).
        */
        void SetLinkedCell(const wxString& linkedCell);

        /**
        Returns the number of entries in a list box or combo box. Returns 0 (zero) if there are no entries in the list.

        [MSDN documentation for ControlFormat.ListCount](https://learn.microsoft.com/en-us/office/vba/api/excel.controlformat.listcount).
        */
        long GetListCount();

        /**
        Returns the worksheet range used to fill the specified list box. Setting this property destroys any existing list in the list box.

        [MSDN documentation for ControlFormat.ListFillRange](https://learn.microsoft.com/en-us/office/vba/api/excel.controlformat.listfillrange).
        */
        wxString GetListFillRange();

        /**
        Sets the worksheet range used to fill the specified list box. Setting this property destroys any existing list in the list box.

        [MSDN documentation for ControlFormat.ListFillRange](https://learn.microsoft.com/en-us/office/vba/api/excel.controlformat.listfillrange).
        */
        void SetListFillRange(const wxString& listFillRange);

        /**
        Returns the index number of the currently selected item in a list box or combo box.

        [MSDN documentation for ControlFormat.ListIndex](https://learn.microsoft.com/en-us/office/vba/api/excel.controlformat.listindex).
        */
        long GetListIndex();

        /**
        Sets the index number of the currently selected item in a list box or combo box.

        [MSDN documentation for ControlFormat.ListIndex](https://learn.microsoft.com/en-us/office/vba/api/excel.controlformat.listindex).
        */
        void SetListIndex(long listIndex);

        /**
        True if the text in the specified object will be locked to prevent changes when the workbook is protected.

        [MSDN documentation for ControlFormat.LockedText](https://learn.microsoft.com/en-us/office/vba/api/excel.controlformat.lockedtext).
        */
        bool GetLockedText();

        /**
        True if the text in the specified object will be locked to prevent changes when the workbook is protected.

        [MSDN documentation for ControlFormat.LockedText](https://learn.microsoft.com/en-us/office/vba/api/excel.controlformat.lockedtext).
        */
        void SetLockedText(bool lockedText);

        /**
        Returns the maximum value of a scroll bar or spinner range. The scroll bar or spinner won’t take on values greater than this maximum value.

        [MSDN documentation for ControlFormat.Max](https://learn.microsoft.com/en-us/office/vba/api/excel.controlformat.max).
        */
        long GetMax();

        /**
        Sets the maximum value of a scroll bar or spinner range. The scroll bar or spinner won’t take on values greater than this maximum value.

        [MSDN documentation for ControlFormat.Max](https://learn.microsoft.com/en-us/office/vba/api/excel.controlformat.max).
        */
        void SetMax(long max);

        /**
        Returns the minimum value of a scroll bar or spinner range. The scroll bar or spinner won’t take on values less than this minimum value.

        [MSDN documentation for ControlFormat.Min](https://learn.microsoft.com/en-us/office/vba/api/excel.controlformat.min).
        */
        long GetMin();

        /**
        Sets the minimum value of a scroll bar or spinner range. The scroll bar or spinner won’t take on values less than this minimum value.

        [MSDN documentation for ControlFormat.Min](https://learn.microsoft.com/en-us/office/vba/api/excel.controlformat.min).
        */
        void SetMin(long min);

        /**
        Returns the selection mode of the specified list box. Can be one of the following constants: xlNone, xlSimple, or xlExtended.

        [MSDN documentation for ControlFormat.MultiSelect](https://learn.microsoft.com/en-us/office/vba/api/excel.controlformat.multiselect).
        */
        long GetMultiSelect();

        /**
        Sets the selection mode of the specified list box. Can be one of the following constants: xlNone, xlSimple, or xlExtended.

        [MSDN documentation for ControlFormat.MultiSelect](https://learn.microsoft.com/en-us/office/vba/api/excel.controlformat.multiselect).
        */
        void SetMultiSelect(long multiSelect);

        /**
        True if the object will be printed when the document is printed.

        [MSDN documentation for ControlFormat.PrintObject](https://learn.microsoft.com/en-us/office/vba/api/excel.controlformat.printobject).
        */
        bool GetPrintObject();

        /**
        True if the object will be printed when the document is printed.

        [MSDN documentation for ControlFormat.PrintObject](https://learn.microsoft.com/en-us/office/vba/api/excel.controlformat.printobject).
        */
        void SetPrintObject(bool printObject);

        /**
        Returns the amount that the scroll bar or spinner is incremented or decremented for a line scroll (when the user clicks an arrow).

        [MSDN documentation for ControlFormat.SmallChange](https://learn.microsoft.com/en-us/office/vba/api/excel.controlformat.smallchange).
        */
        long GetSmallChange();

        /**
        Sets the amount that the scroll bar or spinner is incremented or decremented for a line scroll (when the user clicks an arrow).

        [MSDN documentation for ControlFormat.SmallChange](https://learn.microsoft.com/en-us/office/vba/api/excel.controlformat.smallchange).
        */
        void SetSmallChange(long smallChange);

        /**
        Returns a Long value that represents the name of specified control format.

        [MSDN documentation for ControlFormat.Value](https://learn.microsoft.com/en-us/office/vba/api/excel.controlformat.value).
        */
        long GetValue();

        /**
        Sets a Long value that represents the name of specified control format.

        [MSDN documentation for ControlFormat.Value](https://learn.microsoft.com/en-us/office/vba/api/excel.controlformat.value).
        */
        void SetValue(long value);


        /**
        Returns "ControlFormat".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("ControlFormat"); }
    };


} // namespace wxAutoExcel

#endif // #if WXAUTOEXCEL_USE_SHAPES

#endif //_WXAUTOEXCEL_CONTROLFORMAT_H
