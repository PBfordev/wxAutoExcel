/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Modified by:
// RCS-ID:      $Id: $
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////


#ifndef _WXAUTOEXCEL_CONTROLFORMAT_H
#define _WXAUTOEXCEL_CONTROLFORMAT_H

#include "wx/wxAutoExcel_defs.h"

#if WXAUTOEXCEL_USE_SHAPES

#include "wx/wxAutoExcelObject.h"

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

        [MSDN documentation for ControlFormat.AddItem](http://msdn.microsoft.com/en-us/library/bb209572).
        */
        void AddItem(const wxString& text, long* index);

        //@{
        /**
        Returns or sets the text entries in the specified list box or a combo box, as an array of strings, or returns or sets a single text entry. An error occurs if there are no entries in the list.

        [MSDN documentation for ControlFormat.List](http://msdn.microsoft.com/en-us/library/bb209976).
        */
        wxString List(long index);
        wxArrayString List();
        
        ///@todo Is List a property (and should have a setter too) or a method?

        //@}        

        /**
        Removes all entries from a Microsoft Excel list box or combo box.

        [MSDN documentation for ControlFormat.RemoveAllItems](http://msdn.microsoft.com/en-us/library/bb223591).
        */
        void RemoveAllItems();

        /**
        Removes one or more items from a list box or combo box.

        [MSDN documentation for ControlFormat.RemoveItem](http://msdn.microsoft.com/en-us/library/bb223595).
        */
        void RemoveItem(long index, long* count = NULL);

        // ***** PROPERTIES *****
        

        /**
        Returns the number of list lines displayed in the drop-down portion of a combo box.

        [MSDN documentation for ControlFormat.DropDownLines](http://msdn.microsoft.com/en-us/library/bb221065).
        */
        long GetDropDownLines();

        /**
        Sets the number of list lines displayed in the drop-down portion of a combo box.

        [MSDN documentation for ControlFormat.DropDownLines](http://msdn.microsoft.com/en-us/library/bb221065).
        */
        void SetDropDownLines(long dropDownLines);

        /**
        True if the object is enabled.

        [MSDN documentation for ControlFormat.Enabled](http://msdn.microsoft.com/en-us/library/bb179520).
        */
        bool GetEnabled();

        /**
        True if the object is enabled.

        [MSDN documentation for ControlFormat.Enabled](http://msdn.microsoft.com/en-us/library/bb179520).
        */
        void SetEnabled(bool enabled);

        /**
        Returns the amount that the scroll box increments or decrements for a page scroll (when the user clicks in the scroll bar body region).

        [MSDN documentation for ControlFormat.LargeChange](http://msdn.microsoft.com/en-us/library/bb177828).
        */
        long GetLargeChange();

        /**
        Sets the amount that the scroll box increments or decrements for a page scroll (when the user clicks in the scroll bar body region).

        [MSDN documentation for ControlFormat.LargeChange](http://msdn.microsoft.com/en-us/library/bb177828).
        */
        void SetLargeChange(long largeChange);

        /**
        Returns the worksheet range linked to the control's value. If you place a value in the cell, the control takes this value. Likewise, if you change the value of the control, that value is also placed in the cell.

        [MSDN documentation for ControlFormat.LinkedCell](http://msdn.microsoft.com/en-us/library/bb179523).
        */
        wxString GetLinkedCell();

        /**
        Sets the worksheet range linked to the control's value. If you place a value in the cell, the control takes this value. Likewise, if you change the value of the control, that value is also placed in the cell.

        [MSDN documentation for ControlFormat.LinkedCell](http://msdn.microsoft.com/en-us/library/bb179523).
        */
        void SetLinkedCell(const wxString& linkedCell);

        /**
        Returns the number of entries in a list box or combo box. Returns 0 (zero) if there are no entries in the list.

        [MSDN documentation for ControlFormat.ListCount](http://msdn.microsoft.com/en-us/library/bb177914).
        */
        long GetListCount();

        /**
        Returns the worksheet range used to fill the specified list box. Setting this property destroys any existing list in the list box.

        [MSDN documentation for ControlFormat.ListFillRange](http://msdn.microsoft.com/en-us/library/bb179524).
        */
        wxString GetListFillRange();

        /**
        Sets the worksheet range used to fill the specified list box. Setting this property destroys any existing list in the list box.

        [MSDN documentation for ControlFormat.ListFillRange](http://msdn.microsoft.com/en-us/library/bb179524).
        */
        void SetListFillRange(const wxString& listFillRange);

        /**
        Returns the index number of the currently selected item in a list box or combo box.

        [MSDN documentation for ControlFormat.ListIndex](http://msdn.microsoft.com/en-us/library/bb177920).
        */
        long GetListIndex();

        /**
        Sets the index number of the currently selected item in a list box or combo box.

        [MSDN documentation for ControlFormat.ListIndex](http://msdn.microsoft.com/en-us/library/bb177920).
        */
        void SetListIndex(long listIndex);

        /**
        True if the text in the specified object will be locked to prevent changes when the workbook is protected.

        [MSDN documentation for ControlFormat.LockedText](http://msdn.microsoft.com/en-us/library/bb208700).
        */
        bool GetLockedText();

        /**
        True if the text in the specified object will be locked to prevent changes when the workbook is protected.

        [MSDN documentation for ControlFormat.LockedText](http://msdn.microsoft.com/en-us/library/bb208700).
        */
        void SetLockedText(bool lockedText);

        /**
        Returns the maximum value of a scroll bar or spinner range. The scroll bar or spinner won’t take on values greater than this maximum value.

        [MSDN documentation for ControlFormat.Max](http://msdn.microsoft.com/en-us/library/bb242043).
        */
        long GetMax();

        /**
        Sets the maximum value of a scroll bar or spinner range. The scroll bar or spinner won’t take on values greater than this maximum value.

        [MSDN documentation for ControlFormat.Max](http://msdn.microsoft.com/en-us/library/bb242043).
        */
        void SetMax(long max);

        /**
        Returns the minimum value of a scroll bar or spinner range. The scroll bar or spinner won’t take on values less than this minimum value.

        [MSDN documentation for ControlFormat.Min](http://msdn.microsoft.com/en-us/library/bb242048).
        */
        long GetMin();

        /**
        Sets the minimum value of a scroll bar or spinner range. The scroll bar or spinner won’t take on values less than this minimum value.

        [MSDN documentation for ControlFormat.Min](http://msdn.microsoft.com/en-us/library/bb242048).
        */
        void SetMin(long min);

        /**
        Returns the selection mode of the specified list box. Can be one of the following constants: xlNone, xlSimple, or xlExtended.

        [MSDN documentation for ControlFormat.MultiSelect](http://msdn.microsoft.com/en-us/library/bb208806).
        */
        long GetMultiSelect();

        /**
        Sets the selection mode of the specified list box. Can be one of the following constants: xlNone, xlSimple, or xlExtended.

        [MSDN documentation for ControlFormat.MultiSelect](http://msdn.microsoft.com/en-us/library/bb208806).
        */
        void SetMultiSelect(long multiSelect);        

        /**
        True if the object will be printed when the document is printed.

        [MSDN documentation for ControlFormat.PrintObject](http://msdn.microsoft.com/en-us/library/bb179525).
        */
        bool GetPrintObject();

        /**
        True if the object will be printed when the document is printed.

        [MSDN documentation for ControlFormat.PrintObject](http://msdn.microsoft.com/en-us/library/bb179525).
        */
        void SetPrintObject(bool printObject);

        /**
        Returns the amount that the scroll bar or spinner is incremented or decremented for a line scroll (when the user clicks an arrow).

        [MSDN documentation for ControlFormat.SmallChange](http://msdn.microsoft.com/en-us/library/bb209242).
        */
        long GetSmallChange();

        /**
        Sets the amount that the scroll bar or spinner is incremented or decremented for a line scroll (when the user clicks an arrow).

        [MSDN documentation for ControlFormat.SmallChange](http://msdn.microsoft.com/en-us/library/bb209242).
        */
        void SetSmallChange(long smallChange);

        /**
        Returns a Long value that represents the name of specified control format.

        [MSDN documentation for ControlFormat.Value](http://msdn.microsoft.com/en-us/library/bb214529).
        */
        long GetValue();

        /**
        Sets a Long value that represents the name of specified control format.

        [MSDN documentation for ControlFormat.Value](http://msdn.microsoft.com/en-us/library/bb214529).
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
