/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// License:     MIT license
/////////////////////////////////////////////////////////////////////////////


#ifndef _WXAUTOEXCEL_PAGEBREAKS_H
#define _WXAUTOEXCEL_PAGEBREAKS_H

#include "wx/wxAutoExcel_defs.h"

#include "wx/wxAutoExcel_object.h"
#include "wx/wxAutoExcel_enums.h"

namespace wxAutoExcel {

    /**
    @brief Represents Microsoft Excel HPageBreak or VPageBreak object.
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelPageBreak : public wxExcelObject
    {
    public:
        // ***** METHODS *****

        /**
        Deletes the object.

        [MSDN documentation for HPageBreak.Delete](https://learn.microsoft.com/en-us/office/vba/api/excel.hpagebreak.delete).
        [MSDN documentation for VPageBreak.Delete](https://learn.microsoft.com/en-us/office/vba/api/excel.vpagebreak.delete).
        */
        void Delete();

        /**
        Drags a page break out of the print area.

        [MSDN documentation for HPageBreak.DragOff](https://learn.microsoft.com/en-us/office/vba/api/excel.hpagebreak.dragoff).
        [MSDN documentation for VPageBreak.DragOff](https://learn.microsoft.com/en-us/office/vba/api/excel.vpagebreak.dragoff).
        */
        void DragOff(XlDirection direction, long regionIndex);

        // ***** PROPERTIES *****


        /**
        Returns the type of the specified page break: full-screen or only within a print area. Can be either of the following XlPageBreakExtent constants: xlPageBreakFull or xlPageBreakPartial.

        [MSDN documentation for HPageBreak.Extent](https://learn.microsoft.com/en-us/office/vba/api/excel.hpagebreak.extent).
        [MSDN documentation for VPageBreak.Extent](https://learn.microsoft.com/en-us/office/vba/api/excel.vpagebreak.extent).
        */
        XlPageBreakExtent GetExtent();

        /**
        Returns the cell (a Range object) that defines the page-break location. Horizontal page breaks are aligned with the top edge of the location cell; vertical page breaks are aligned with the left edge of the location cell.

        [MSDN documentation for HPageBreak.Location](https://learn.microsoft.com/en-us/office/vba/api/excel.hpagebreak.location).
        [MSDN documentation for VPageBreak.Location](https://learn.microsoft.com/en-us/office/vba/api/excel.vpagebreak.location).
        */
        wxExcelRange GetLocation();

        /**
        Sets the cell (a Range object) that defines the page-break location. Horizontal page breaks are aligned with the top edge of the location cell; vertical page breaks are aligned with the left edge of the location cell.

        [MSDN documentation for HPageBreak.Location](https://learn.microsoft.com/en-us/office/vba/api/excel.hpagebreak.location).
        [MSDN documentation for VPageBreak.Location](https://learn.microsoft.com/en-us/office/vba/api/excel.vpagebreak.location).
        */
        void SetLocation(wxExcelRange location);

        /**
        Returns a XlPageBreak value that represents the page break type.

        [MSDN documentation for HPageBreak.Type](https://learn.microsoft.com/en-us/office/vba/api/excel.hpagebreak.type).
        [MSDN documentation for VPageBreak.Type](https://learn.microsoft.com/en-us/office/vba/api/excel.vpagebreak.type).
        */
        XlPageBreak GetType();

        /**
        Sets a XlPageBreak value that represents the page break type.

        [MSDN documentation for HPageBreak.Type](https://learn.microsoft.com/en-us/office/vba/api/excel.hpagebreak.type).
        [MSDN documentation for VPageBreak.Type](https://learn.microsoft.com/en-us/office/vba/api/excel.vpagebreak.type).
        */
        void SetType(XlPageBreak type);

        /**
        Returns "PageBreak".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("PageBreak"); }
    };


    /**
    @brief Represents Microsoft Excel HPageBreaks or VPageBreaks collection.
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelPageBreaks : public wxExcelObject
    {
    public:
        // ***** METHODS *****

        /**
        Adds a horizontal page break.

        [MSDN documentation for HPageBreaks.Add](https://learn.microsoft.com/en-us/office/vba/api/excel.hpagebreaks.add).
        [MSDN documentation for VPageBreaks.Add](https://learn.microsoft.com/en-us/office/vba/api/excel.vpagebreaks.add).
        */
        wxExcelPageBreak Add(wxExcelRange before);

        // ***** PROPERTIES *****


        /**
        Returns a Long value that represents the number of objects in the collection.

        [MSDN documentation for HPageBreaks.Count](https://learn.microsoft.com/en-us/office/vba/api/excel.hpagebreaks.count).
        [MSDN documentation for VPageBreaks.Count](https://learn.microsoft.com/en-us/office/vba/api/excel.vpagebreaks.count).
        */
        long GetCount();

        //@{
        /**
        Returns a single object from a collection.

        [MSDN documentation for HPageBreaks.Item](https://learn.microsoft.com/en-us/office/vba/api/excel.hpagebreaks.item).
        [MSDN documentation for VPageBreaks.Item](https://learn.microsoft.com/en-us/office/vba/api/excel.vpagebreaks.item).
        */
        wxExcelPageBreak GetItem(long index);
        wxExcelPageBreak operator[](long index);
        //@}

        /**
        Returns "PageBreaks".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("PageBreaks"); }
    };


} // namespace wxAutoExcel

#endif //_WXAUTOEXCEL_PAGEBREAKS_H
